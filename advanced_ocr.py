# advanced_ocr.py - Version 100% offline
import cv2
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import re
from typing import List, Dict, Tuple
import logging
import io

logger = logging.getLogger(__name__)

class AdvancedOCRProcessor:
    def __init__(self):
        # Configuration Tesseract optimisée pour tickets
        self.tesseract_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzàáâäéèêëíìîïóòôöúùûü.,:-€$£¥₹/°% '
        self.tesseract_config_numbers = r'--oem 3 --psm 8 -c tessedit_char_whitelist=0123456789.,-€$£¥ '
    
    def preprocess_image(self, image_bytes: bytes) -> np.ndarray:
        """Pipeline de préprocessing complète offline"""
        # Conversion PIL
        pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        
        # Auto-orientation basique (sans EXIF, juste détection)
        pil_img = self._auto_orient_basic(pil_img)
        
        # Conversion OpenCV
        img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        
        # Pipeline de nettoyage
        img = self._enhance_image_quality(img)
        
        return img
    
    def _auto_orient_basic(self, pil_img: Image.Image) -> Image.Image:
        """Orientation basique sans EXIF - détection par analyse d'image"""
        try:
            # Convertir pour analyse
            gray = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2GRAY)
            
            # Essayer les 4 orientations et choisir celle avec le plus de texte horizontal
            orientations = [0, 90, 180, 270]
            scores = []
            
            for angle in orientations:
                if angle > 0:
                    rotated = pil_img.rotate(angle, expand=True)
                    test_gray = cv2.cvtColor(np.array(rotated), cv2.COLOR_RGB2GRAY)
                else:
                    test_gray = gray
                
                # Score basé sur les lignes horizontales
                score = self._score_text_orientation(test_gray)
                scores.append((angle, score))
            
            # Prendre la meilleure orientation
            best_angle = max(scores, key=lambda x: x[1])[0]
            if best_angle > 0:
                pil_img = pil_img.rotate(best_angle, expand=True)
                logger.info(f"Image auto-orientée de {best_angle}°")
                
        except Exception as e:
            logger.warning(f"Erreur auto-orientation: {e}")
        
        return pil_img
    
    def _score_text_orientation(self, gray: np.ndarray) -> float:
        """Score une image selon la probabilité d'orientation correcte du texte"""
        try:
            # Détecter les contours
            edges = cv2.Canny(gray, 50, 150, apertureSize=3)
            
            # Lignes de Hough pour détecter l'horizontalité
            lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=50)
            
            if lines is None:
                return 0.0
            
            horizontal_score = 0
            for rho, theta in lines[:20]:
                # Angle par rapport à l'horizontale
                angle = abs(np.degrees(theta) - 90)
                if angle > 90:
                    angle = 180 - angle
                
                # Score plus élevé pour les lignes proches de l'horizontal
                if angle < 10:  # Quasi-horizontal
                    horizontal_score += 1
            
            return horizontal_score
            
        except:
            return 0.0
    
    def _enhance_image_quality(self, img: np.ndarray) -> np.ndarray:
        """Amélioration qualité image pour OCR"""
        try:
            # Redimensionner si trop petit
            h, w = img.shape[:2]
            if min(h, w) < 800:
                scale = 800 / min(h, w)
                new_w, new_h = int(w * scale), int(h * scale)
                img = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_CUBIC)
            
            # Débruitage
            img = cv2.fastNlMeansDenoisingColored(img, None, 10, 10, 7, 21)
            
            # Amélioration contraste
            lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
            l, a, b = cv2.split(lab)
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
            l = clahe.apply(l)
            img = cv2.cvtColor(cv2.merge([l, a, b]), cv2.COLOR_LAB2BGR)
            
            # Affûtage
            kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
            img = cv2.filter2D(img, -1, kernel)
            
            return img
            
        except Exception as e:
            logger.warning(f"Erreur amélioration image: {e}")
            return img
    
    def extract_text_tesseract(self, img: np.ndarray) -> List[Dict]:
        """Extraction avec Tesseract local"""
        try:
            # Conversion en PIL pour Tesseract
            pil_img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
            
            # Extraction avec données de position
            data = pytesseract.image_to_data(
                pil_img, 
                config=self.tesseract_config,
                output_type=pytesseract.Output.DICT
            )
            
            lines = []
            n_boxes = len(data['level'])
            
            for i in range(n_boxes):
                confidence = int(data['conf'][i])
                text = data['text'][i].strip()
                
                # Filtrer les détections faibles
                if confidence > 30 and len(text) > 0:
                    x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                    
                    lines.append({
                        "text": text,
                        "confidence": confidence / 100.0,  # Normaliser 0-1
                        "bbox": [x, y, x + w, y + h],
                        "area": w * h
                    })
            
            # Tri spatial
            lines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))
            
            logger.info(f"Tesseract: {len(lines)} éléments détectés")
            return lines
            
        except Exception as e:
            logger.error(f"Erreur Tesseract: {e}")
            return []
    
    def extract_structured_info(self, ocr_lines: List[Dict]) -> Dict:
        """Extraction d'infos structurées (identique à avant)"""
        info = {
            "merchant": None,
            "date": None,
            "total": None,
            "currency": "EUR",
            "items": [],
            "vat_rate": None,
            "confidence": 0.0
        }
        
        if not ocr_lines:
            return info
        
        # Reconstruire le texte par lignes spatiales
        text_lines = self._reconstruct_text_lines(ocr_lines)
        
        # Calculer confiance moyenne
        confidences = [line["confidence"] for line in ocr_lines if line["confidence"] > 0.3]
        info["confidence"] = sum(confidences) / len(confidences) if confidences else 0.0
        
        # Détection commerçant (première ligne substantielle)
        for line_text in text_lines[:5]:  # Top 5 lignes
            clean_text = line_text.strip()
            if len(clean_text) > 3 and sum(c.isalpha() for c in clean_text) > 2:
                info["merchant"] = clean_text
                break
        
        # Détection date
        full_text = " ".join(text_lines)
        date_patterns = [
            r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.](?:\d{4}|\d{2}))',
            r'(\d{1,2}\s+(?:jan|fév|mar|avr|mai|jun|jul|aoû|sep|oct|nov|déc)\s+\d{2,4})',
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, full_text, re.IGNORECASE)
            if match:
                try:
                    from dateparser import parse as parse_date
                    parsed_date = parse_date(match.group(1), languages=['fr', 'en'])
                    if parsed_date:
                        info["date"] = parsed_date.date().isoformat()
                        break
                except ImportError:
                    # Parsing manuel si dateparser indisponible
                    date_str = match.group(1)
                    if self._is_valid_date_format(date_str):
                        info["date"] = date_str
                    break
                except:
                    continue
        
        # Détection montants
        amounts = self._extract_amounts(text_lines)
        if amounts:
            # Choisir le plus gros montant comme total probable
            info["total"] = max(amounts)
        
        # Détection TVA
        tva_match = re.search(r'(?:tva|vat|tax).*?(\d{1,2}[,.]?\d{0,2})\s*%', 
                             full_text, re.IGNORECASE)
        if tva_match:
            try:
                info["vat_rate"] = float(tva_match.group(1).replace(',', '.'))
            except:
                pass
        
        return info
    
    def _reconstruct_text_lines(self, ocr_lines: List[Dict]) -> List[str]:
        """Reconstitue les lignes de texte à partir des éléments OCR"""
        if not ocr_lines:
            return []
        
        # Grouper par hauteur Y approximative (tolérance de quelques pixels)
        line_groups = []
        tolerance = 10
        
        for element in ocr_lines:
            y_center = (element["bbox"][1] + element["bbox"][3]) // 2
            
            # Trouver un groupe existant proche
            placed = False
            for group in line_groups:
                group_y = sum((e["bbox"][1] + e["bbox"][3]) // 2 for e in group) // len(group)
                if abs(y_center - group_y) <= tolerance:
                    group.append(element)
                    placed = True
                    break
            
            if not placed:
                line_groups.append([element])
        
        # Trier chaque groupe par position X et reconstituer le texte
        text_lines = []
        for group in line_groups:
            group.sort(key=lambda e: e["bbox"][0])
            line_text = " ".join(e["text"] for e in group)
            text_lines.append(line_text)
        
        return text_lines
    
    def _extract_amounts(self, text_lines: List[str]) -> List[float]:
        """Extraction des montants depuis les lignes de texte"""
        amounts = []
        
        # Patterns pour montants
        amount_patterns = [
            r'(\d+[,.]?\d{0,2})\s*€',
            r'(\d+[,.]?\d{0,2})\s*EUR',
            r'€\s*(\d+[,.]?\d{0,2})',
            r'(\d+[,.]?\d{0,2})\s*$',
        ]
        
        for line in text_lines:
            for pattern in amount_patterns:
                matches = re.findall(pattern, line)
                for match in matches:
                    try:
                        amount = float(match.replace(',', '.'))
                        if 0.1 <= amount <= 10000:  # Filtres réalistes
                            amounts.append(amount)
                    except:
                        continue
        
        return amounts
    
    def _is_valid_date_format(self, date_str: str) -> bool:
        """Validation basique d'un format de date"""
        patterns = [
            r'\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}',
            r'\d{1,2}\s+\w+\s+\d{2,4}'
        ]
        return any(re.match(pattern, date_str) for pattern in patterns)

    def process_ticket_image(self, image_bytes: bytes, filename: str) -> Dict:
        """Pipeline complète offline"""
        try:
            # Préprocessing
            img_processed = self.preprocess_image(image_bytes)
            
            # OCR
            ocr_lines = self.extract_text_tesseract(img_processed)
            
            # Extraction structurée
            structured_info = self.extract_structured_info(ocr_lines)
            
            # Reconstituer le texte complet
            full_text = "\n".join([line["text"] for line in ocr_lines])
            
            return {
                "filename": filename,
                "raw_text": full_text,
                "extraction_method": "tesseract_advanced",
                "ocr_lines": ocr_lines,
                "lines_detected": len(ocr_lines),
                "average_confidence": structured_info["confidence"],
                **structured_info
            }
            
        except Exception as e:
            logger.error(f"Erreur pipeline OCR offline: {e}")
            return {
                "filename": filename,
                "error": str(e),
                "raw_text": "",
                "extraction_method": "failed"
            }
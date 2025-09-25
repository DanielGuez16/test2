# advanced_ocr.py
import cv2
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
from paddleocr import PaddleOCR
import re
from rapidfuzz import fuzz, process
from dateparser import parse as parse_date
from price_parser import Price
from typing import List, Dict, Tuple
import logging
import io

logger = logging.getLogger(__name__)

class AdvancedOCRProcessor:
    def __init__(self):
        self.ocr = PaddleOCR(lang='fr', use_angle_cls=True, show_log=False)
    
    def preprocess_image(self, image_bytes: bytes) -> np.ndarray:
        """Pipeline de préprocessing complète sans piexif"""
        # Conversion PIL -> OpenCV
        pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        
        # Correction orientation avec Pillow natif
        pil_img = self._auto_orient_pillow(pil_img)
        
        # Conversion BGR pour OpenCV
        img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
        
        # Crop des bordures
        h, w = img.shape[:2]
        crop_margin = min(h, w) // 50
        if crop_margin > 5:
            img = img[crop_margin:-crop_margin, crop_margin:-crop_margin]
        
        # Débruitage bilatéral
        img = cv2.bilateralFilter(img, 7, 50, 50)
        
        # Amélioration contraste avec CLAHE
        img = self._enhance_contrast(img)
        
        # Affûtage (unsharp mask)
        img = self._sharpen_image(img)
        
        # Détection et correction d'inclinaison
        img = self._correct_skew(img)
        
        return img
    
    def _auto_orient_pillow(self, pil_img: Image.Image) -> Image.Image:
        """Correction orientation avec Pillow natif (sans piexif)"""
        try:
            # Utiliser Pillow's getexif() (disponible depuis Pillow 6.0)
            exif = pil_img.getexif()
            
            # Code d'orientation EXIF
            orientation_key = 274  # Code standard pour orientation
            
            if orientation_key in exif:
                orientation = exif[orientation_key]
                
                if orientation == 3:
                    pil_img = pil_img.rotate(180, expand=True)
                elif orientation == 6:
                    pil_img = pil_img.rotate(270, expand=True)
                elif orientation == 8:
                    pil_img = pil_img.rotate(90, expand=True)
                
                logger.info(f"Image orientée : orientation EXIF {orientation}")
            
        except Exception as e:
            logger.warning(f"Impossible de lire EXIF pour orientation: {e}")
        
        return pil_img
    
    def _enhance_contrast(self, img: np.ndarray) -> np.ndarray:
        """Amélioration contraste avec CLAHE sur canal L"""
        lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)
        
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        l_enhanced = clahe.apply(l)
        
        return cv2.cvtColor(cv2.merge([l_enhanced, a, b]), cv2.COLOR_LAB2BGR)
    
    def _sharpen_image(self, img: np.ndarray) -> np.ndarray:
        """Affûtage avec unsharp mask"""
        blur = cv2.GaussianBlur(img, (0, 0), 1.0)
        return cv2.addWeighted(img, 1.5, blur, -0.5, 0)
    
    def _correct_skew(self, img: np.ndarray) -> np.ndarray:
        """Correction automatique d'inclinaison"""
        try:
            # Conversion en niveaux de gris
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # Binarisation
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Détection des lignes avec HoughLines
            edges = cv2.Canny(binary, 50, 150, apertureSize=3)
            lines = cv2.HoughLines(edges, 1, np.pi/180, threshold=100)
            
            if lines is not None:
                angles = []
                for rho, theta in lines[:20]:  # Prendre les 20 premières lignes
                    angle = np.degrees(theta) - 90
                    if -45 < angle < 45:  # Filtrer les angles raisonnables
                        angles.append(angle)
                
                if angles:
                    # Angle médian pour éviter les outliers
                    skew_angle = np.median(angles)
                    
                    # Corriger seulement si l'angle est significatif (> 0.5°)
                    if abs(skew_angle) > 0.5:
                        h, w = img.shape[:2]
                        center = (w // 2, h // 2)
                        
                        # Matrice de rotation
                        M = cv2.getRotationMatrix2D(center, skew_angle, 1.0)
                        
                        # Calculer les nouvelles dimensions
                        cos = np.abs(M[0, 0])
                        sin = np.abs(M[0, 1])
                        new_w = int((h * sin) + (w * cos))
                        new_h = int((h * cos) + (w * sin))
                        
                        # Ajuster la translation
                        M[0, 2] += (new_w / 2) - center[0]
                        M[1, 2] += (new_h / 2) - center[1]
                        
                        # Appliquer la rotation
                        img = cv2.warpAffine(img, M, (new_w, new_h), 
                                           flags=cv2.INTER_CUBIC,
                                           borderMode=cv2.BORDER_REPLICATE)
                        
                        logger.info(f"Image redressée de {skew_angle:.2f}°")
            
        except Exception as e:
            logger.warning(f"Erreur correction inclinaison: {e}")
        
        return img
    
    def extract_text_paddleocr(self, img: np.ndarray) -> List[Dict]:
        """Extraction texte avec PaddleOCR"""
        try:
            results = self.ocr.ocr(img, cls=True)
            
            lines = []
            for result_page in results:
                if result_page is None:
                    continue
                    
                for detection in result_page:
                    box, (text, confidence) = detection
                    
                    # Calculer bounding box
                    x_coords = [point[0] for point in box]
                    y_coords = [point[1] for point in box]
                    x1, y1 = int(min(x_coords)), int(min(y_coords))
                    x2, y2 = int(max(x_coords)), int(max(y_coords))
                    
                    lines.append({
                        "text": text,
                        "confidence": float(confidence),
                        "bbox": [x1, y1, x2, y2],
                        "area": (x2 - x1) * (y2 - y1)
                    })
            
            # Tri top-to-bottom, left-to-right
            lines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))
            
            logger.info(f"PaddleOCR: {len(lines)} lignes détectées")
            return lines
            
        except Exception as e:
            logger.error(f"Erreur PaddleOCR: {e}")
            return []

    def extract_structured_info(self, ocr_lines: List[Dict]) -> Dict:
        """Extraction d'informations structurées depuis les lignes OCR"""
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
        
        # Calculer confiance moyenne
        confidences = [line["confidence"] for line in ocr_lines if line["confidence"] > 0.3]
        info["confidence"] = sum(confidences) / len(confidences) if confidences else 0.0
        
        # Détection commerçant (lignes du haut avec lettres)
        top_lines = [l for l in ocr_lines if l["bbox"][1] < 150]
        for line in top_lines:
            text = line["text"].strip()
            if len(text) > 3 and sum(c.isalpha() for c in text) > len(text) * 0.5:
                if not info["merchant"] or line["confidence"] > 0.8:
                    info["merchant"] = text
                break
        
        # Détection date
        date_pattern = r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.](?:\d{4}|\d{2}))'
        for line in ocr_lines:
            match = re.search(date_pattern, line["text"])
            if match:
                try:
                    parsed_date = parse_date(match.group(1), languages=['fr', 'en'])
                    if parsed_date:
                        info["date"] = parsed_date.date().isoformat()
                        break
                except:
                    continue
        
        # Détection montants et total
        amounts = []
        for line in ocr_lines:
            try:
                price = Price.fromstring(line["text"])
                if price.amount and float(price.amount.replace(',', '.')) > 0:
                    amount_value = float(price.amount.replace(',', '.'))
                    amounts.append({
                        "value": amount_value,
                        "currency": price.currency or "EUR",
                        "text": line["text"],
                        "confidence": line["confidence"],
                        "is_total": any(word in line["text"].lower() 
                                      for word in ["total", "ttc", "à payer", "amount"])
                    })
            except:
                continue
        
        # Choisir le total (priorité aux lignes marquées "total", sinon max)
        if amounts:
            total_candidates = [a for a in amounts if a["is_total"]]
            if total_candidates:
                best_total = max(total_candidates, key=lambda x: x["confidence"])
            else:
                best_total = max(amounts, key=lambda x: x["value"])
            
            info["total"] = best_total["value"]
            info["currency"] = best_total["currency"]
        
        # Détection TVA
        for line in ocr_lines:
            vat_match = re.search(r'(?:tva|vat).*?(\d{1,2}[,.]?\d{0,2})\s*%', 
                                line["text"].lower())
            if vat_match:
                try:
                    info["vat_rate"] = float(vat_match.group(1).replace(',', '.'))
                    break
                except:
                    continue
        
        return info

    def process_ticket_image(self, image_bytes: bytes, filename: str) -> Dict:
        """Pipeline complète de traitement d'image de ticket"""
        try:
            # Préprocessing
            img_processed = self.preprocess_image(image_bytes)
            
            # OCR
            ocr_lines = self.extract_text_paddleocr(img_processed)
            
            # Extraction structurée
            structured_info = self.extract_structured_info(ocr_lines)
            
            # Reconstituer le texte complet
            full_text = "\n".join([line["text"] for line in ocr_lines])
            
            return {
                "filename": filename,
                "raw_text": full_text,
                "extraction_method": "paddleocr_advanced",
                "ocr_lines": ocr_lines,
                "lines_detected": len(ocr_lines),
                "average_confidence": structured_info["confidence"],
                **structured_info  # Merge structured info
            }
            
        except Exception as e:
            logger.error(f"Erreur pipeline OCR: {e}")
            return {
                "filename": filename,
                "error": str(e),
                "raw_text": "",
                "extraction_method": "failed"
            }
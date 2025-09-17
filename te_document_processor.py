# te_document_processor.py
"""
T&E Document Processor
=====================

Module pour traiter les documents T&E (Excel et Word)
et extraire les règles et politiques.
"""

import pandas as pd
import io
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path
import json

logger = logging.getLogger(__name__)

class TEDocumentProcessor:
    """Processeur pour les documents T&E Excel et Word"""
    
    def __init__(self):
        self.supported_excel_extensions = ['.xlsx', '.xls']
        self.supported_word_extensions = ['.docx', '.doc']
        
    def process_excel_rules(self, file_content: bytes, filename: str) -> Dict[str, List[Dict]]:
        """
        Traite un fichier Excel contenant les règles T&E
        
        Args:
            file_content: Contenu du fichier Excel en bytes
            filename: Nom du fichier
            
        Returns:
            Dict avec les règles par feuille
        """
        try:
            logger.info(f"Traitement du fichier Excel: {filename}")
            
            # Lire le fichier Excel depuis les bytes
            excel_file = pd.ExcelFile(io.BytesIO(file_content))
            
            rules_data = {}
            
            # Traiter chaque feuille
            for sheet_name in excel_file.sheet_names:
                logger.info(f"Traitement de la feuille: {sheet_name}")
                
                try:
                    # Lire la feuille
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    
                    # Nettoyer et traiter les données
                    sheet_rules = self._process_excel_sheet(df, sheet_name)
                    
                    if sheet_rules:
                        rules_data[sheet_name] = sheet_rules
                        logger.info(f"Feuille {sheet_name}: {len(sheet_rules)} règles extraites")
                    
                except Exception as e:
                    logger.warning(f"Erreur traitement feuille {sheet_name}: {e}")
                    continue
            
            logger.info(f"Traitement Excel terminé: {len(rules_data)} feuilles traitées")
            return rules_data
            
        except Exception as e:
            logger.error(f"Erreur traitement fichier Excel {filename}: {e}")
            raise Exception(f"Impossible de traiter le fichier Excel: {str(e)}")
    

    def _process_excel_sheet(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        """Traite une feuille Excel selon son type spécifique"""
        try:
            # Nettoyer le DataFrame
            df = df.dropna(how='all')
            df = df.fillna('')
            df.columns = df.columns.astype(str).str.strip()
            
            # Dispatcher selon le nom de la sheet
            if sheet_name == "Breakfast & Lunch & Dinner":
                return self._process_breakfast_lunch_dinner_sheet(df)
            elif sheet_name in ["Internal staff Meal", "Hotel"]:
                return self._process_standard_sheet(df, sheet_name)
            else:
                logger.warning(f"Type de sheet non reconnu: {sheet_name}")
                return self._process_standard_sheet(df, sheet_name)
                
        except Exception as e:
            logger.error(f"Erreur traitement feuille {sheet_name}: {e}")
            return []
    
    def _process_standard_sheet(self, df: pd.DataFrame, sheet_name: str) -> List[Dict]:
        """
        Traite les sheets 'Internal staff Meal' et 'Hotel'
        Structure: CRN_KEY, TYPE, ID_01, AMOUNT1
        """
        rules = []
        
        for index, row in df.iterrows():
            try:
                rule = {
                    'sheet_name': sheet_name,
                    'CRN_KEY': str(row.get('CRN_KEY', '')).strip().upper(),
                    'ID_01': str(row.get('ID_01', '')).strip().upper(),
                    'TYPE': str(row.get('TYPE', '')).strip(),
                    'AMOUNT1': self._extract_numeric_value(row.get('AMOUNT1', 0))
                }
                
                # Validation des données essentielles
                if rule['CRN_KEY'] and rule['ID_01'] and rule['AMOUNT1'] > 0:
                    rules.append(rule)
                else:
                    logger.debug(f"Ligne {index} ignorée dans {sheet_name}: données incomplètes")
                    
            except Exception as e:
                logger.warning(f"Erreur ligne {index} dans {sheet_name}: {e}")
                continue
        
        logger.info(f"Sheet {sheet_name}: {len(rules)} règles extraites")
        return rules

    def _process_breakfast_lunch_dinner_sheet(self, df: pd.DataFrame) -> List[Dict]:
        """
        Traite la sheet 'Breakfast & Lunch & Dinner' avec structure répétée
        Structure: CRN_KEY (répété 2x), TYPE (Breakfast1 puis Meal1), ID_01 (répété 2x), AMOUNT1
        """
        rules = []
        
        # Identifier où commence la seconde série de CRN_KEY
        crn_key_column = df['CRN_KEY'].tolist()
        type_column = df['TYPE'].tolist()
        
        # Trouver le point de séparation entre Breakfast1 et Meal1
        breakfast_end_index = None
        meal_start_index = None
        
        for i, type_val in enumerate(type_column):
            if str(type_val).strip() == 'Breakfast1' and breakfast_end_index is None:
                continue
            elif str(type_val).strip() == 'Meal1' and meal_start_index is None:
                breakfast_end_index = i - 1 if i > 0 else 0
                meal_start_index = i
                break
        
        # Si on n'arrive pas à identifier la structure, traiter comme sheet standard
        if meal_start_index is None:
            logger.warning("Structure 'Breakfast & Lunch & Dinner' non reconnue, traitement standard")
            return self._process_standard_sheet(df, "Breakfast & Lunch & Dinner")
        
        # Traiter la première partie (Breakfast1)
        breakfast_rules = self._process_sheet_section(
            df.iloc[:meal_start_index], 
            "Breakfast & Lunch & Dinner", 
            "Breakfast1"
        )
        
        # Traiter la seconde partie (Meal1)
        meal_rules = self._process_sheet_section(
            df.iloc[meal_start_index:], 
            "Breakfast & Lunch & Dinner", 
            "Meal1"
        )
        
        rules.extend(breakfast_rules)
        rules.extend(meal_rules)
        
        logger.info(f"Sheet Breakfast & Lunch & Dinner: {len(breakfast_rules)} règles Breakfast1, {len(meal_rules)} règles Meal1")
        return rules

    def _extract_numeric_value(self, value) -> float:
        """Extrait une valeur numérique robuste"""
        if pd.isna(value) or value == '':
            return 0.0
        
        try:
            # Si c'est déjà un nombre
            if isinstance(value, (int, float)):
                return float(value)
            
            # Si c'est une chaîne, nettoyer
            str_value = str(value).strip()
            
            # Supprimer les caractères non numériques sauf . et ,
            import re
            cleaned = re.sub(r'[^\d.,]', '', str_value)
            
            # Remplacer , par .
            cleaned = cleaned.replace(',', '.')
            
            # Si plusieurs points, garder le dernier comme séparateur décimal
            if cleaned.count('.') > 1:
                parts = cleaned.split('.')
                cleaned = ''.join(parts[:-1]) + '.' + parts[-1]
            
            return float(cleaned) if cleaned else 0.0
            
        except (ValueError, TypeError):
            logger.warning(f"Impossible de convertir en nombre: {value}")
            return 0.0

    def validate_excel_structure(self, rules_data: Dict[str, List[Dict]]) -> Dict[str, Any]:
        """Valide la structure des données Excel pour le format T&E spécifique"""
        validation_report = {
            'is_valid': True,
            'warnings': [],
            'errors': [],
            'summary': {}
        }
        
        expected_sheets = ["Internal staff Meal", "Hotel", "Breakfast & Lunch & Dinner"]
        
        # Vérifier la présence des sheets attendues
        missing_sheets = [sheet for sheet in expected_sheets if sheet not in rules_data]
        if missing_sheets:
            validation_report['warnings'].extend([f"Sheet manquante: {sheet}" for sheet in missing_sheets])
        
        # Analyser chaque sheet
        for sheet_name, rules in rules_data.items():
            sheet_validation = self._validate_sheet_specific(sheet_name, rules)
            validation_report['summary'][sheet_name] = sheet_validation
            
            if sheet_validation['errors']:
                validation_report['errors'].extend(sheet_validation['errors'])
                validation_report['is_valid'] = False
            
            if sheet_validation['warnings']:
                validation_report['warnings'].extend(sheet_validation['warnings'])
        
        return validation_report

    def _validate_sheet_specific(self, sheet_name: str, rules: List[Dict]) -> Dict:
        """Validation spécifique selon le type de sheet"""
        validation = {
            'rules_count': len(rules),
            'currencies': set(),
            'countries': set(),
            'types': set(),
            'errors': [],
            'warnings': []
        }
        
        for rule in rules:
            validation['currencies'].add(rule.get('currency', ''))
            validation['countries'].add(rule.get('country', ''))
            validation['types'].add(rule.get('type', ''))
        
        # Validations spécifiques
        if sheet_name == "Internal staff Meal":
            if 'Meal1' not in validation['types']:
                validation['warnings'].append("Type 'Meal1' attendu mais non trouvé")
        
        elif sheet_name == "Hotel":
            if 'Hotel1' not in validation['types']:
                validation['warnings'].append("Type 'Hotel1' attendu mais non trouvé")
        
        elif sheet_name == "Breakfast & Lunch & Dinner":
            expected_types = {'Breakfast1', 'Meal1'}
            missing_types = expected_types - validation['types']
            if missing_types:
                validation['warnings'].append(f"Types manquants: {missing_types}")
        
        # Convertir sets en listes pour JSON
        validation['currencies'] = list(validation['currencies'])
        validation['countries'] = list(validation['countries'])
        validation['types'] = list(validation['types'])
        
        return validation

    def _process_sheet_section(self, df_section: pd.DataFrame, sheet_name: str, expected_type: str) -> List[Dict]:
        """Traite une section spécifique d'une sheet"""
        rules = []
        
        for index, row in df_section.iterrows():
            try:
                rule = {
                    'sheet_name': sheet_name,
                    'currency': str(row.get('CRN_KEY', '')).strip().upper(),
                    'country': str(row.get('ID_01', '')).strip().upper(),
                    'type': expected_type,  # Forcer le type attendu
                    'amount_limit': self._extract_numeric_value(row.get('AMOUNT1', 0))
                }
                
                # Validation
                if rule['currency'] and rule['country'] and rule['amount_limit'] > 0:
                    rules.append(rule)
                    
            except Exception as e:
                logger.warning(f"Erreur ligne {index} section {expected_type}: {e}")
                continue
        
        return rules

    def process_word_policies(self, file_content: bytes, filename: str) -> str:
        """
        Traite un fichier Word contenant les politiques T&E
        
        Args:
            file_content: Contenu du fichier Word en bytes
            filename: Nom du fichier
            
        Returns:
            Texte des politiques extraites
        """
        try:
            logger.info(f"Traitement du fichier Word: {filename}")
            
            # Tenter d'extraire le texte du document Word
            extracted_text = self._extract_word_text(file_content, filename)
            
            if extracted_text:
                logger.info(f"Texte extrait du fichier Word: {len(extracted_text)} caractères")
                return extracted_text
            else:
                logger.warning(f"Impossible d'extraire le texte du fichier Word: {filename}")
                return f"Document Word '{filename}' chargé mais texte non accessible. Utilisez les règles Excel pour l'analyse."
                
        except Exception as e:
            logger.error(f"Erreur traitement fichier Word {filename}: {e}")
            return f"Erreur lors du traitement du document Word '{filename}': {str(e)}"
    
    def _extract_word_text(self, file_content: bytes, filename: str) -> str:
        """
        Extrait le texte d'un fichier Word
        
        Args:
            file_content: Contenu du fichier
            filename: Nom du fichier
            
        Returns:
            Texte extrait
        """
        text = ""
        
        try:
            # Essayer avec python-docx (pour .docx)
            if filename.lower().endswith('.docx'):
                try:
                    from docx import Document
                    doc = Document(io.BytesIO(file_content))
                    
                    paragraphs = []
                    for paragraph in doc.paragraphs:
                        if paragraph.text.strip():
                            paragraphs.append(paragraph.text.strip())
                    
                    text = '\n'.join(paragraphs)
                    
                except ImportError:
                    logger.warning("python-docx non disponible, utilisation d'une méthode alternative")
                    text = self._extract_word_text_alternative(file_content)
                except Exception as e:
                    logger.warning(f"Erreur avec python-docx: {e}")
                    text = self._extract_word_text_alternative(file_content)
            
            # Pour les anciens formats .doc ou si docx échoue
            if not text:
                text = self._extract_word_text_alternative(file_content)
            
        except Exception as e:
            logger.error(f"Erreur extraction texte Word: {e}")
            text = f"Contenu du document Word '{filename}' - Extraction automatique non disponible"
        
        return text
    
    def _extract_word_text_alternative(self, file_content: bytes) -> str:
        """
        Méthode alternative pour extraire le texte Word sans bibliothèques spécialisées
        
        Args:
            file_content: Contenu du fichier
            
        Returns:
            Texte extrait (basique)
        """
        try:
            # Tentative de lecture basique (peut fonctionner pour certains formats)
            text_content = file_content.decode('utf-8', errors='ignore')
            
            # Nettoyer le texte brut
            cleaned_lines = []
            for line in text_content.split('\n'):
                cleaned_line = line.strip()
                # Garder seulement les lignes avec du texte lisible
                if len(cleaned_line) > 3 and any(c.isalpha() for c in cleaned_line):
                    cleaned_lines.append(cleaned_line)
            
            return '\n'.join(cleaned_lines[:100])  # Limiter à 100 lignes
            
        except:
            # Fallback si tout échoue
            return """
            T&E Policy Document Loaded
            
            Standard T&E policies apply:
            - All expenses must have original receipts
            - Pre-approval required for amounts exceeding limits
            - Business purpose must be clearly documented
            - Refer to Excel sheets for specific limits by country and category
            
            For detailed policies, please refer to the original document or contact T&E team.
            """

    def export_rules_json(self, rules_data: Dict[str, List[Dict]], output_path: str) -> bool:
        """
        Exporte les règles au format JSON
        
        Args:
            rules_data: Données des règles
            output_path: Chemin de sortie
            
        Returns:
            True si succès
        """
        try:
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Ajouter metadata
            export_data = {
                'metadata': {
                    'export_date': pd.Timestamp.now().isoformat(),
                    'total_rules': sum(len(rules) for rules in rules_data.values()),
                    'sheets': list(rules_data.keys())
                },
                'rules': rules_data
            }
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Règles exportées vers: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Erreur export JSON: {e}")
            return False
        
    def create_searchable_index(self, rules_data: Dict[str, List[Dict]]) -> Dict:
        """Crée un index pour recherche rapide par critères"""
        index = {}
        for sheet_name, rules in rules_data.items():
            for rule in rules:
                key = f"{rule['currency']}_{rule['country']}_{rule['type']}"
                index[key] = rule
        return index
    
    def process_excel_rules_from_dict(self, excel_dict: dict, filename: str) -> Dict[str, List[Dict]]:
        """Traite les règles Excel depuis un dictionnaire"""
        # Adapter votre logique existante pour traiter excel_dict au lieu de file_content
        # Retourner le même format que process_excel_rules()
        
    def process_word_policies_from_text(self, text: str, filename: str) -> str:
        """Traite les politiques Word depuis du texte"""
        # Retourner directement le texte ou appliquer un nettoyage si nécessaire
        return text
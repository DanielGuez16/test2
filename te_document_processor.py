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
        """
        Traite une feuille Excel spécifique
        
        Args:
            df: DataFrame pandas de la feuille
            sheet_name: Nom de la feuille
            
        Returns:
            Liste des règles extraites
        """
        rules = []
        
        try:
            # Nettoyer le DataFrame
            df = df.dropna(how='all')  # Supprimer les lignes complètement vides
            df = df.fillna('')  # Remplacer NaN par chaîne vide
            
            # Standardiser les noms de colonnes
            df.columns = df.columns.astype(str).str.strip()
            
            # Colonnes attendues dans les fichiers T&E
            expected_columns = ['CRN_KEY', 'ID_01', 'TYPE', 'AMOUNT1', 'AMOUNT2']
            
            # Vérifier les colonnes disponibles
            available_columns = df.columns.tolist()
            logger.info(f"Colonnes disponibles dans {sheet_name}: {available_columns}")
            
            # Mapper les colonnes si nécessaire
            column_mapping = self._map_columns(available_columns, expected_columns)
            
            # Traiter chaque ligne
            for index, row in df.iterrows():
                try:
                    rule = self._extract_rule_from_row(row, column_mapping, sheet_name)
                    if rule:
                        rules.append(rule)
                except Exception as e:
                    logger.warning(f"Erreur ligne {index} dans {sheet_name}: {e}")
                    continue
            
            return rules
            
        except Exception as e:
            logger.error(f"Erreur traitement feuille {sheet_name}: {e}")
            return []
    
    def _map_columns(self, available_columns: List[str], expected_columns: List[str]) -> Dict[str, str]:
        """
        Mappe les colonnes disponibles aux colonnes attendues
        
        Args:
            available_columns: Colonnes disponibles dans le fichier
            expected_columns: Colonnes attendues
            
        Returns:
            Mapping des colonnes
        """
        mapping = {}
        
        # Mapping exact d'abord
        for col in expected_columns:
            if col in available_columns:
                mapping[col] = col
        
        # Mapping approximatif ensuite
        approximations = {
            'CRN_KEY': ['currency', 'devise', 'crn', 'curr'],
            'ID_01': ['country', 'pays', 'location', 'id'],
            'TYPE': ['type', 'category', 'categorie', 'description'],
            'AMOUNT1': ['amount', 'limite', 'limit', 'montant', 'max'],
            'AMOUNT2': ['amount2', 'limite2', 'limit2', 'montant2']
        }
        
        for expected, variations in approximations.items():
            if expected not in mapping:
                for available in available_columns:
                    available_lower = available.lower().strip()
                    if any(var.lower() in available_lower for var in variations):
                        mapping[expected] = available
                        break
        
        logger.info(f"Mapping colonnes: {mapping}")
        return mapping
    
    def _extract_rule_from_row(self, row: pd.Series, column_mapping: Dict[str, str], sheet_name: str) -> Optional[Dict]:
        """
        Extrait une règle depuis une ligne de données
        
        Args:
            row: Ligne de données pandas
            column_mapping: Mapping des colonnes
            sheet_name: Nom de la feuille
            
        Returns:
            Règle extraite ou None
        """
        try:
            rule = {
                'sheet_name': sheet_name,
                'CRN_KEY': '',
                'ID_01': '',
                'TYPE': '',
                'AMOUNT1': 0,
                'AMOUNT2': 0
            }
            
            # Extraire les valeurs selon le mapping
            for expected_col, actual_col in column_mapping.items():
                if actual_col in row.index:
                    value = row[actual_col]
                    
                    # Traitement spécial pour les montants
                    if expected_col in ['AMOUNT1', 'AMOUNT2']:
                        try:
                            # Nettoyer et convertir en nombre
                            if pd.isna(value) or value == '':
                                rule[expected_col] = 0
                            else:
                                # Supprimer les caractères non numériques sauf . et ,
                                cleaned_value = str(value).replace(',', '.').strip()
                                # Extraire le nombre
                                import re
                                number_match = re.search(r'[\d.]+', cleaned_value)
                                if number_match:
                                    rule[expected_col] = float(number_match.group())
                                else:
                                    rule[expected_col] = 0
                        except:
                            rule[expected_col] = 0
                    else:
                        # Traitement pour les chaînes
                        rule[expected_col] = str(value).strip() if not pd.isna(value) else ''
            
            # Vérifier que la règle est valide
            if rule['CRN_KEY'] and rule['AMOUNT1'] > 0:
                return rule
            else:
                return None
                
        except Exception as e:
            logger.warning(f"Erreur extraction règle: {e}")
            return None
    
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
    
    def validate_excel_structure(self, rules_data: Dict[str, List[Dict]]) -> Dict[str, Any]:
        """
        Valide la structure des données Excel extraites
        
        Args:
            rules_data: Données extraites
            
        Returns:
            Rapport de validation
        """
        validation_report = {
            'is_valid': True,
            'warnings': [],
            'errors': [],
            'summary': {}
        }
        
        try:
            total_rules = 0
            currencies = set()
            countries = set()
            
            for sheet_name, rules in rules_data.items():
                sheet_summary = {
                    'rules_count': len(rules),
                    'currencies': set(),
                    'countries': set(),
                    'has_amounts': 0
                }
                
                for rule in rules:
                    total_rules += 1
                    
                    # Collecter les devises
                    if rule.get('CRN_KEY'):
                        currencies.add(rule['CRN_KEY'])
                        sheet_summary['currencies'].add(rule['CRN_KEY'])
                    
                    # Collecter les pays
                    if rule.get('ID_01'):
                        countries.add(rule['ID_01'])
                        sheet_summary['countries'].add(rule['ID_01'])
                    
                    # Vérifier les montants
                    if rule.get('AMOUNT1', 0) > 0:
                        sheet_summary['has_amounts'] += 1
                
                validation_report['summary'][sheet_name] = {
                    'rules_count': sheet_summary['rules_count'],
                    'currencies': list(sheet_summary['currencies']),
                    'countries': list(sheet_summary['countries']),
                    'has_amounts': sheet_summary['has_amounts']
                }
                
                # Vérifications
                if sheet_summary['rules_count'] == 0:
                    validation_report['warnings'].append(f"Feuille '{sheet_name}' ne contient aucune règle valide")
                
                if sheet_summary['has_amounts'] == 0:
                    validation_report['warnings'].append(f"Feuille '{sheet_name}' ne contient aucun montant valide")
            
            validation_report['summary']['total'] = {
                'rules_count': total_rules,
                'currencies': list(currencies),
                'countries': list(countries),
                'sheets_count': len(rules_data)
            }
            
            # Vérifications globales
            if total_rules == 0:
                validation_report['is_valid'] = False
                validation_report['errors'].append("Aucune règle valide trouvée dans le fichier Excel")
            
            if len(currencies) == 0:
                validation_report['warnings'].append("Aucune devise trouvée dans les règles")
            
            logger.info(f"Validation Excel: {total_rules} règles, {len(currencies)} devises, {len(countries)} pays")
            
        except Exception as e:
            validation_report['is_valid'] = False
            validation_report['errors'].append(f"Erreur lors de la validation: {str(e)}")
        
        return validation_report
    
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


def test_te_processor():
    """Teste le processeur T&E avec des données simulées"""
    processor = TEDocumentProcessor()
    
    print("=== Test TEDocumentProcessor ===")
    
    # Créer des données Excel simulées
    test_data = {
        'Internal staff Meal': [
            {'CRN_KEY': 'EUR', 'ID_01': 'FR', 'TYPE': 'Lunch', 'AMOUNT1': 25.0, 'AMOUNT2': 0},
            {'CRN_KEY': 'USD', 'ID_01': 'US', 'TYPE': 'Lunch', 'AMOUNT1': 30.0, 'AMOUNT2': 0},
        ],
        'Hotel': [
            {'CRN_KEY': 'EUR', 'ID_01': 'FR', 'TYPE': 'Standard', 'AMOUNT1': 150.0, 'AMOUNT2': 0},
            {'CRN_KEY': 'USD', 'ID_01': 'US', 'TYPE': 'Standard', 'AMOUNT1': 180.0, 'AMOUNT2': 0},
        ]
    }
    
    # Valider la structure
    validation = processor.validate_excel_structure(test_data)
    print(f"Validation: {validation['is_valid']}")
    print(f"Règles total: {validation['summary']['total']['rules_count']}")
    print(f"Devises: {validation['summary']['total']['currencies']}")
    print(f"Avertissements: {len(validation['warnings'])}")
    
    # Test processing Word
    test_word_content = b"Test policy document content"
    word_result = processor.process_word_policies(test_word_content, "test_policy.docx")
    print(f"Texte Word extrait: {len(word_result)} caractères")


if __name__ == "__main__":
    test_te_processor()
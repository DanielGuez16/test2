# ticket_analyzer.py
"""
Ticket Analyzer for T&E System
==============================

Analyse les tickets de frais en utilisant le système RAG et l'IA
pour valider la conformité avec les politiques T&E.
"""

import logging
from typing import Dict, List, Any, Optional
from datetime import datetime
import re
import json

from context_builder import TEContextBuilder
from llm_connector import LLMConnector

logger = logging.getLogger(__name__)

class TicketAnalyzer:
    """Analyseur de tickets T&E avec IA et RAG"""
    
    def __init__(self, rag_system, llm_connector: LLMConnector):
        self.rag_system = rag_system
        self.llm_connector = llm_connector
        self.context_builder = TEContextBuilder()
    
    def analyze_ticket(self, ticket_info: dict, user_question: str = "") -> dict:
        """
        Analyse complète d'un ticket T&E
        
        Args:
            ticket_info: Informations extraites du ticket par IA
            user_question: Question spécifique de l'utilisateur
            
        Returns:
            dict: Résultat complet de l'analyse au format business
        """
        logger.info(f"Début analyse ticket: {ticket_info.get('filename', 'N/A')}")
        
        try:
            # 1. Extraire les critères d'analyse
            analysis_criteria = self._extract_analysis_criteria(ticket_info)
            
            # 2. Rechercher les règles pertinentes via RAG
            relevant_rules = self._find_relevant_rules(analysis_criteria)
            
            # 3. Obtenir le contexte des politiques
            policies_context = self._get_policies_context(analysis_criteria)
            
            # 4. Validation basique contre les règles trouvées
            basic_validation = self._perform_basic_validation(ticket_info, relevant_rules)
            
            # 5. Préparer le contexte pour l'IA
            context = self.context_builder.build_context(
                'ticket_analysis',
                ticket_info=ticket_info,
                relevant_rules=relevant_rules,
                policies_context=policies_context
            )
            
            # 6. Préparer le prompt utilisateur
            prompt = self.context_builder.build_prompt_for_ticket_analysis(user_question, ticket_info)
            
            # 7. Obtenir l'analyse IA
            ai_response = self.llm_connector.get_llm_response(prompt, context)
            
            # 8. Structurer le résultat final au FORMAT BUSINESS
            analysis_result = {
                'result': 'PASS' if basic_validation['is_valid'] else 'FAIL',
                'expense_type': self._determine_expense_type(ticket_info, relevant_rules),
                'justification': self._build_justification(relevant_rules, basic_validation, ticket_info),
                'comment': ai_response,  # Commentaire IA complet
                'confidence_score': self._calculate_confidence(relevant_rules, ticket_info),
                'applied_rules': relevant_rules,  # Pour debug
                'timestamp': datetime.now().isoformat()
            }
            
            logger.info(f"Analyse terminée avec succès - Result: {analysis_result['result']}")
            return analysis_result
            
        except Exception as e:
            logger.error(f"Erreur lors de l'analyse du ticket: {str(e)}")
            return {
                'result': 'FAIL',
                'expense_type': 'Erreur d\'analyse',
                'justification': f'Erreur technique: {str(e)}',
                'comment': f"Erreur lors de l'analyse: {str(e)}",
                'confidence_score': 0.0,
                'applied_rules': [],
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }
    
    def answer_general_question(self, user_question: str, te_rules_summary: dict) -> dict:
        """
        Répond à une question générale sur les politiques T&E
        """
        try:
            # Obtenir le contexte général
            policies_context = self._get_general_policies_context(user_question)
            
            # Construire le contexte
            context = self.context_builder.build_context(
                'general_query',
                te_rules_summary=te_rules_summary,
                policies_excerpt=policies_context
            )
            
            # Construire le prompt
            prompt = self.context_builder.build_prompt_for_general_query(user_question)
            
            # Obtenir la réponse IA
            ai_response = self.llm_connector.get_llm_response(prompt, context)
            
            return {
                'ai_response': ai_response,
                'question_type': 'general_policy',
                'timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            logger.error(f"Erreur réponse question générale: {str(e)}")
            return {
                'ai_response': f"Erreur lors du traitement de votre question: {str(e)}",
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }
    
    def _extract_analysis_criteria(self, ticket_info: dict) -> dict:
        """Extrait les critères d'analyse du ticket"""
        return {
            'currency': ticket_info.get('currency', 'EUR'),
            'country': ticket_info.get('country_code', 'FR'),
            'expense_type': self._map_category_to_expense_type(ticket_info.get('category', 'unknown')),
            'amount': ticket_info.get('amount', 0),
            'category': ticket_info.get('category', 'meal')
        }
    
    def _map_category_to_expense_type(self, category: str) -> str:
        """Mappe une catégorie IA vers un type de dépense Excel"""
        if not category:
            return 'Meal1'
        
        category_lower = category.lower()
        
        mapping = {
            'hotel': 'Hotel1',
            'accommodation': 'Hotel1',
            'lodging': 'Hotel1',
            'meal': 'Meal1',
            'restaurant': 'Meal1',
            'food': 'Meal1',
            'breakfast': 'Breakfast1',
            'lunch': 'Meal1',
            'dinner': 'Meal1'
        }
        
        return mapping.get(category_lower, 'Meal1')
    
    def _find_relevant_rules(self, criteria: dict) -> List[Dict]:
        """Trouve les règles pertinentes via le système RAG"""
        try:
            if hasattr(self.rag_system, 'search_relevant_rules') and self.rag_system:
                # Utiliser le système RAG
                query = f"{criteria['expense_type']} {criteria['country']} {criteria['currency']}"
                return self.rag_system.search_relevant_rules(query, filters=criteria)
            else:
                # Fallback: recherche directe dans les données Excel
                return self._fallback_rule_search_with_excel_data(criteria)
        except Exception as e:
            logger.warning(f"Erreur recherche règles RAG: {e}, utilisation fallback")
            return self._fallback_rule_search_with_excel_data(criteria)
    
    def _fallback_rule_search_with_excel_data(self, criteria: dict) -> List[Dict]:
        """Recherche de règles en fallback"""
        from run import te_documents
        
        rules = []
        excel_rules = te_documents.get("excel_rules")
        
        if not excel_rules:
            logger.warning("Aucune donnée Excel chargée")
            return []
        
        currency = criteria.get("currency", "")
        country = criteria.get("country", "")
        expense_type = criteria.get("expense_type", "")
        
        # Mapper les types de dépenses aux sheets Excel
        sheet_mapping = {
            "Hotel1": "Hotel",
            "Meal1": "Internal staff Meal",
            "Breakfast1": "Breakfast & Lunch & Dinner"
        }
        
        # Chercher dans la sheet appropriée
        target_sheets = []
        if expense_type in sheet_mapping:
            target_sheets.append(sheet_mapping[expense_type])
        else:
            target_sheets = list(excel_rules.keys())
        
        for sheet_name in target_sheets:
            if sheet_name not in excel_rules:
                continue
                
            sheet_rules = excel_rules[sheet_name]
            for rule in sheet_rules:
                rule_matches = True
                
                if currency and rule.get("currency") != currency:
                    rule_matches = False
                
                if country and rule.get("country") != country:
                    rule_matches = False
                
                if sheet_name == "Breakfast & Lunch & Dinner":
                    if expense_type == "Meal1" and rule.get("type") not in ["Meal1", "Breakfast1"]:
                        rule_matches = False
                    elif expense_type == "Breakfast1" and rule.get("type") != "Breakfast1":
                        rule_matches = False
                else:
                    if expense_type and rule.get("type") != expense_type:
                        rule_matches = False
                
                if rule_matches:
                    formatted_rule = {
                        "sheet_name": sheet_name,
                        "currency": rule.get("currency"),
                        "country": rule.get("country"),
                        "type": rule.get("type"),
                        "amount_limit": rule.get("amount_limit", 0)
                    }
                    rules.append(formatted_rule)
        
        # Si pas de correspondance exacte, chercher partiellement
        if not rules:
            for sheet_name, sheet_rules in excel_rules.items():
                for rule in sheet_rules:
                    if currency and rule.get("currency") == currency:
                        formatted_rule = {
                            "sheet_name": sheet_name,
                            "currency": rule.get("currency"),
                            "country": rule.get("country"),
                            "type": rule.get("type"),
                            "amount_limit": rule.get("amount_limit", 0)
                        }
                        rules.append(formatted_rule)
                        
                        if len(rules) >= 5:
                            break
                if len(rules) >= 5:
                    break
        
        logger.info(f"Fallback search trouvé {len(rules)} règles pour {criteria}")
        return rules[:10]
    
    def _ai_extract_ticket_info(self, raw_text: str, filename: str) -> dict:
        """Extrait TOUTES les informations du ticket via IA + RAG"""
        try:
            # Construire le prompt d'extraction
            extraction_prompt = f"""
            Analyse ce ticket et extrait les informations suivantes au format JSON :
            {{
                "amount": float ou null,
                "currency": "EUR/USD/etc" ou null,
                "category": "hotel/meal/transport/flight" ou "unknown",
                "subcategory": "breakfast/lunch/dinner/accommodation" ou null,
                "date": "YYYY-MM-DD" ou null,
                "vendor": "nom établissement" ou null,
                "location": "ville/pays" ou null,
                "country_code": "FR/US/etc" ou null,
                "confidence": float entre 0 et 1
            }}
            
            Ticket à analyser:
            {raw_text[:2000]}
            """
            
            # Obtenir le contexte des règles T&E
            context = self.context_builder.build_context(
                'ticket_extraction',
                te_rules_summary=self._get_rules_summary()
            )
            
            # Appel IA
            ai_response = self.llm_connector.get_llm_response(extraction_prompt, context)
            
            # Parser la réponse JSON
            try:
                extracted_info = json.loads(ai_response)
            except:
                # Fallback si JSON malformé
                extracted_info = self._parse_ai_response_fallback(ai_response)
            
            # Ajouter métadonnées
            extracted_info.update({
                "filename": filename,
                "raw_text": raw_text[:1500],
                "extraction_method": "ai_rag"
            })
            
            return extracted_info
            
        except Exception as e:
            logger.error(f"Erreur extraction IA: {e}")
            # Fallback vers méthode basique
            return self._fallback_extraction(raw_text, filename)

    def _get_rules_summary(self) -> dict:
        """Résumé des règles pour le contexte IA"""
        from run import te_documents
        summary = {}
        
        if te_documents.get("excel_rules"):
            for sheet_name, rules in te_documents["excel_rules"].items():
                currencies = set(rule.get("currency") for rule in rules if rule.get("currency"))
                countries = set(rule.get("country") for rule in rules if rule.get("country"))
                summary[sheet_name] = {
                    "rules_count": len(rules),
                    "currencies": list(currencies),
                    "countries": list(countries)
                }
        
        return summary

    def _parse_ai_response_fallback(self, ai_response: str) -> dict:
        """Parse la réponse IA si JSON échoue"""
        info = {
            "amount": None,
            "currency": None,
            "category": "unknown",
            "confidence": 0.5
        }
        
        # Extraire montant avec regex
        amount_match = re.search(r'"amount":\s*([0-9.]+)', ai_response)
        if amount_match:
            try:
                info["amount"] = float(amount_match.group(1))
            except ValueError:
                pass
        
        # Extraire devise
        currency_match = re.search(r'"currency":\s*"([A-Z]{3})"', ai_response)
        if currency_match:
            info["currency"] = currency_match.group(1)
        
        # Extraire catégorie
        category_match = re.search(r'"category":\s*"([^"]+)"', ai_response)
        if category_match:
            info["category"] = category_match.group(1)
        
        return info

    def _fallback_extraction(self, raw_text: str, filename: str) -> dict:
        """Fallback si l'extraction IA échoue"""
        return {
            "filename": filename,
            "raw_text": raw_text[:1500],
            "amount": None,
            "currency": None,
            "category": "unknown",
            "extraction_method": "fallback",
            "error": "AI extraction failed"
        }

    def ai_extract_ticket_info(self, raw_text: str, filename: str) -> dict:
        """Méthode publique pour extraction IA"""
        return self._ai_extract_ticket_info(raw_text, filename)

    def _get_policies_context(self, criteria: dict) -> str:
        """Obtient le contexte des politiques pertinentes"""
        try:
            if hasattr(self.rag_system, 'search_policies') and self.rag_system:
                category = criteria.get('category', 'general') or 'general'
                if category == 'unknown':
                    category = 'general'
                return self.rag_system.search_policies(category)
            else:
                # Fallback: utiliser les données Word depuis run.py
                from run import te_documents
                word_policies = te_documents.get("word_policies", "")
                
                if word_policies:
                    category = criteria.get('category', '').lower()
                    relevant_sections = []
                    
                    paragraphs = word_policies.split('\n\n')
                    for paragraph in paragraphs:
                        if len(paragraph) > 50:
                            if category in paragraph.lower():
                                relevant_sections.append(paragraph.strip())
                    
                    if relevant_sections:
                        return '\n\n'.join(relevant_sections[:3])
                    else:
                        return word_policies[:500] + "..."
                
                return "Politiques T&E standard - Les frais doivent être justifiés et dans les limites approuvées."
        except Exception as e:
            logger.warning(f"Erreur récupération contexte politiques: {e}")
            return "Contexte politiques non disponible."
    
    def _get_general_policies_context(self, question: str) -> str:
        """Obtient le contexte général des politiques pour une question"""
        question_lower = question.lower()
        
        if any(word in question_lower for word in ['hotel', 'accommodation', 'lodging']):
            return "Politiques hôtels: Pré-approbation requise pour séjours > 3 nuits..."
        elif any(word in question_lower for word in ['meal', 'restaurant', 'food', 'dining']):
            return "Politiques repas: Limites par pays et devise, justificatifs requis..."
        else:
            return "Politiques générales T&E applicables..."
    
    def _perform_basic_validation(self, ticket_info: dict, relevant_rules: List[Dict]) -> dict:
        """Effectue une validation basique contre les règles"""
        validation = {
            'is_valid': False,
            'status': 'pending_review',
            'issues': [],
            'within_limits': False
        }
        
        amount = ticket_info.get('amount', 0)
        if not amount or amount <= 0:
            validation['issues'].append("Montant non détecté ou invalide")
            return validation
        
        # Vérifier contre les règles trouvées
        if relevant_rules:
            for rule in relevant_rules:
                limit = rule.get('amount_limit', 0)
                if amount <= limit:
                    validation['is_valid'] = True
                    validation['status'] = 'approved'
                    validation['within_limits'] = True
                    break
                else:
                    validation['issues'].append(f"Montant {amount} dépasse la limite de {limit}")
        else:
            validation['issues'].append("Aucune règle applicable trouvée")
        
        return validation
    
    def _determine_expense_type(self, ticket_info: dict, rules: List[Dict]) -> str:
        """Détermine le type de dépense en français business"""
        category = ticket_info.get('category', 'unknown')
        if not category:
            category = 'unknown'
            
        category_lower = category.lower()
        subcategory = ticket_info.get('subcategory', '').lower()
        
        # Mapping vers des termes business clairs
        expense_mapping = {
            'hotel': 'Hébergement',
            'accommodation': 'Hébergement',
            'meal': 'Repas',
            'restaurant': 'Repas',
            'food': 'Repas',
            'breakfast': 'Petit-déjeuner',
            'lunch': 'Déjeuner',
            'dinner': 'Dîner',
            'transport': 'Transport',
            'taxi': 'Taxi',
            'car_rental': 'Location de voiture',
            'flight': 'Vol',
            'train': 'Train',
            'unknown': 'Dépense non catégorisée'
        }
        
        # Priorité aux sous-catégories
        if subcategory and subcategory in expense_mapping:
            return expense_mapping[subcategory]
        elif category_lower and category_lower in expense_mapping:
            return expense_mapping[category_lower]
        else:
            return 'Dépense non catégorisée'

    def _build_justification(self, rules: List[Dict], validation: dict, ticket_info: dict) -> str:
        """Construit la justification basée sur les règles"""
        if not rules:
            return "Aucune règle applicable trouvée dans les documents T&E"
        
        if validation['is_valid']:
            rule = rules[0]  # Première règle applicable
            amount = ticket_info.get('amount', 0)
            return f"Conforme à la règle '{rule.get('sheet_name', '')}': montant {amount} {rule.get('currency', '')} ≤ limite {rule.get('amount_limit', 0)} {rule.get('currency', '')} pour {rule.get('country', '')}"
        else:
            issues = validation.get('issues', [])
            return f"Non conforme: {'; '.join(issues)}"
    
    def _calculate_confidence(self, rules: List[Dict], ticket_info: dict) -> float:
        """Calcule un score de confiance pour l'analyse"""
        confidence = 0.5  # Base
        
        # +0.3 si des règles ont été trouvées
        if rules:
            confidence += 0.3
        
        # +0.1 si montant détecté
        if ticket_info.get('amount', 0) > 0:
            confidence += 0.1
        
        # +0.1 si devise détectée
        if ticket_info.get('currency'):
            confidence += 0.1
        
        return min(confidence, 1.0)


def test_ticket_analyzer():
    """Teste l'analyseur de tickets"""
    from llm_connector import LLMConnector
    
    # Mock RAG system pour les tests
    class MockRAGSystem:
        def search_relevant_rules(self, query, filters=None):
            return [{
                'type': 'Hotel1',
                'country': 'FR',
                'currency': 'EUR',
                'amount_limit': 200,
                'sheet_name': 'Hotel'
            }]
    
    # Initialiser
    rag_system = MockRAGSystem()
    llm_connector = LLMConnector()
    analyzer = TicketAnalyzer(rag_system, llm_connector)
    
    # Ticket de test
    ticket_info = {
        'filename': 'hotel_paris.pdf',
        'amount': 150,
        'currency': 'EUR',
        'category': 'hotel',
        'location': 'Paris',
        'vendor': 'Hotel Marais'
    }
    
    # Analyser
    result = analyzer.analyze_ticket(ticket_info, "Ce montant est-il acceptable?")
    
    print("=== TEST TICKET ANALYZER ===")
    print(f"Result: {result['result']}")
    print(f"Expense Type: {result['expense_type']}")
    print(f"Justification: {result['justification']}")
    print(f"Confidence: {result['confidence_score']}")


if __name__ == "__main__":
    test_ticket_analyzer()
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

from context_builder import TEContextBuilder
from llm_connector import LLMConnector

logger = logging.getLogger(__name__)

class TicketAnalyzer:
    """Analyseur de tickets T&E avec IA et RAG"""
    
    def __init__(self, rag_system, llm_connector: LLMConnector):
        self.rag_system = rag_system
        self.llm_connector = llm_connector
        self.context_builder = TEContextBuilder()
        
        # Mappings pour normaliser les données
        self.country_mappings = {
            'france': 'FR', 'paris': 'FR', 'lyon': 'FR',
            'germany': 'DE', 'berlin': 'DE', 'munich': 'DE',
            'australia': 'AU', 'sydney': 'AU', 'melbourne': 'AU',
            'united states': 'US', 'usa': 'US', 'new york': 'US',
            'uae': 'AE', 'dubai': 'AE', 'abu dhabi': 'AE',
            'switzerland': 'CH', 'zurich': 'CH', 'geneva': 'CH'
        }
        
        self.expense_type_mappings = {
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
    
    def analyze_ticket(self, ticket_info: dict, user_question: str = "") -> dict:
        """
        Analyse complète d'un ticket T&E
        
        Args:
            ticket_info: Informations extraites du ticket
            user_question: Question spécifique de l'utilisateur
            
        Returns:
            dict: Résultat complet de l'analyse
        """
        logger.info(f"Début analyse ticket: {ticket_info.get('filename', 'N/A')}")
        
        try:
            # 1. Normaliser et enrichir les données du ticket
            normalized_ticket = self._normalize_ticket_data(ticket_info)
            
            # 2. Extraire les critères d'analyse
            analysis_criteria = self._extract_analysis_criteria(normalized_ticket)
            
            # 3. Rechercher les règles pertinentes via RAG
            relevant_rules = self._find_relevant_rules(analysis_criteria)
            
            # 4. Obtenir le contexte des politiques
            policies_context = self._get_policies_context(analysis_criteria)
            
            # 5. Validation basique contre les règles trouvées
            basic_validation = self._perform_basic_validation(normalized_ticket, relevant_rules)
            
            # 6. Préparer le contexte pour l'IA
            context = self.context_builder.build_context(
                'ticket_analysis',
                ticket_info=normalized_ticket,
                relevant_rules=relevant_rules,
                policies_context=policies_context
            )
            
            # 7. Préparer le prompt utilisateur
            prompt = self.context_builder.build_prompt_for_ticket_analysis(user_question, normalized_ticket)
            
            # 8. Obtenir l'analyse IA
            ai_response = self.llm_connector.get_llm_response(prompt, context)
            
            # 9. Structurer le résultat final
            analysis_result = {
                'ai_response': ai_response,
                'basic_validation': basic_validation,
                'applied_rules': relevant_rules,
                'analysis_criteria': analysis_criteria,
                'policies_context': policies_context[:200] + "..." if len(policies_context) > 200 else policies_context,
                'confidence_score': self._calculate_confidence(relevant_rules, normalized_ticket),
                'recommendations': self._generate_recommendations(normalized_ticket, relevant_rules, basic_validation),
                'timestamp': datetime.now().isoformat()
            }
            
            logger.info(f"Analyse terminée avec succès - Confidence: {analysis_result['confidence_score']}")
            return analysis_result
            
        except Exception as e:
            logger.error(f"Erreur lors de l'analyse du ticket: {str(e)}")
            return {
                'ai_response': f"Erreur lors de l'analyse: {str(e)}",
                'basic_validation': {'is_valid': False, 'status': 'error'},
                'applied_rules': [],
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }
    
    def answer_general_question(self, user_question: str, te_rules_summary: dict) -> dict:
        """
        Répond à une question générale sur les politiques T&E
        
        Args:
            user_question: Question de l'utilisateur
            te_rules_summary: Résumé des règles disponibles
            
        Returns:
            dict: Réponse structurée
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
    
    def _normalize_ticket_data(self, ticket_info: dict) -> dict:
        """Normalise et enrichit les données du ticket"""
        normalized = ticket_info.copy()
        
        # Normaliser le pays
        location = ticket_info.get('location', '').lower()
        vendor = ticket_info.get('vendor', '').lower()
        description = ticket_info.get('description', '').lower()
        
        # Chercher le pays dans différents champs
        for text in [location, vendor, description]:
            for location_name, country_code in self.country_mappings.items():
                if location_name in text:
                    normalized['country_code'] = country_code
                    break
            if 'country_code' in normalized:
                break
        
        # Normaliser le type de dépense
        category = ticket_info.get('category', '').lower()
        if category in self.expense_type_mappings:
            normalized['expense_type'] = self.expense_type_mappings[category]
        else:
            normalized['expense_type'] = 'Meal1'  # Défaut
        
        # Normaliser la devise
        currency = ticket_info.get('currency', 'EUR').upper()
        normalized['currency'] = currency
        
        return normalized
    
    def _extract_analysis_criteria(self, ticket_info: dict) -> dict:
        """Extrait les critères d'analyse du ticket"""
        return {
            'currency': ticket_info.get('currency', 'EUR'),
            'country': ticket_info.get('country_code', 'FR'),
            'expense_type': ticket_info.get('expense_type', 'Meal1'),
            'amount': ticket_info.get('amount', 0),
            'category': ticket_info.get('category', 'meal')
        }
    
    def _find_relevant_rules(self, criteria: dict) -> List[Dict]:
        """Trouve les règles pertinentes via le système RAG"""
        try:
            if hasattr(self.rag_system, 'search_relevant_rules') and self.rag_system:
                # Utiliser le système RAG
                query = f"{criteria['expense_type']} {criteria['country']} {criteria['currency']}"
                return self.rag_system.search_relevant_rules(query, filters=criteria)
            else:
                # Fallback: recherche directe dans les données Excel chargées depuis run.py
                return self._fallback_rule_search_with_excel_data(criteria)
        except Exception as e:
            logger.warning(f"Erreur recherche règles RAG: {e}, utilisation fallback")
            return self._fallback_rule_search_with_excel_data(criteria)
    
    def _fallback_rule_search_with_excel_data(self, criteria: dict) -> List[Dict]:
        """Recherche de règles en fallback en accédant aux données Excel depuis run.py"""
        from run import te_documents  # Import direct depuis run.py
        
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
            # Chercher dans toutes les sheets
            target_sheets = list(excel_rules.keys())
        
        for sheet_name in target_sheets:
            if sheet_name not in excel_rules:
                continue
                
            sheet_rules = excel_rules[sheet_name]
            for rule in sheet_rules:
                # Vérifier correspondance avec les critères
                rule_matches = True
                
                if currency and rule.get("currency") != currency:
                    rule_matches = False
                
                if country and rule.get("country") != country:
                    rule_matches = False
                
                # Pour Breakfast & Lunch & Dinner, être plus flexible sur le type
                if sheet_name == "Breakfast & Lunch & Dinner":
                    if expense_type == "Meal1" and rule.get("type") not in ["Meal1", "Breakfast1"]:
                        rule_matches = False
                    elif expense_type == "Breakfast1" and rule.get("type") != "Breakfast1":
                        rule_matches = False
                else:
                    if expense_type and rule.get("type") != expense_type:
                        rule_matches = False
                
                if rule_matches:
                    # Convertir au format attendu
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
                    partial_match = False
                    
                    if currency and rule.get("currency") == currency:
                        partial_match = True
                    elif country and rule.get("country") == country:
                        partial_match = True
                    
                    if partial_match:
                        formatted_rule = {
                            "sheet_name": sheet_name,
                            "currency": rule.get("currency"),
                            "country": rule.get("country"),
                            "type": rule.get("type"),
                            "amount_limit": rule.get("amount_limit", 0)
                        }
                        rules.append(formatted_rule)
                        
                        # Limiter le nombre de résultats partiels
                        if len(rules) >= 5:
                            break
                if len(rules) >= 5:
                    break
        
        logger.info(f"Fallback search trouvé {len(rules)} règles pour {criteria}")
        return rules[:10]  # Limiter à 10 résultats
    
    def _get_policies_context(self, criteria: dict) -> str:
        """Obtient le contexte des politiques pertinentes"""
        try:
            if hasattr(self.rag_system, 'search_policies') and self.rag_system:
                return self.rag_system.search_policies(criteria['category'])
            else:
                # Fallback: utiliser les données Word depuis run.py
                from run import te_documents
                word_policies = te_documents.get("word_policies", "")
                
                if word_policies:
                    # Recherche simple par mots-clés dans les politiques
                    category = criteria.get('category', '').lower()
                    relevant_sections = []
                    
                    paragraphs = word_policies.split('\n\n')
                    for paragraph in paragraphs:
                        if len(paragraph) > 50:  # Ignorer les paragraphes trop courts
                            if category in paragraph.lower():
                                relevant_sections.append(paragraph.strip())
                    
                    if relevant_sections:
                        return '\n\n'.join(relevant_sections[:3])  # Max 3 sections
                    else:
                        return word_policies[:500] + "..."  # Début du document
                
                return "Politiques T&E standard - Les frais doivent être justifiés et dans les limites approuvées."
        except Exception as e:
            logger.warning(f"Erreur récupération contexte politiques: {e}")
            return "Contexte politiques non disponible."
    
    def _get_general_policies_context(self, question: str) -> str:
        """Obtient le contexte général des politiques pour une question"""
        # Analyser la question pour identifier les sections pertinentes
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
        if amount <= 0:
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
    
    def _generate_recommendations(self, ticket_info: dict, rules: List[Dict], validation: dict) -> List[str]:
        """Génère des recommandations basées sur l'analyse"""
        recommendations = []
        
        if validation['is_valid']:
            recommendations.append("✅ Dépense conforme aux politiques T&E")
        else:
            recommendations.append("⚠️ Dépense nécessite une révision")
        
        if validation['issues']:
            recommendations.append("📋 Vérifier les justificatifs requis")
        
        if not rules:
            recommendations.append("📞 Contacter l'équipe T&E pour clarification")
        
        return recommendations


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
    print(f"Status: {result['basic_validation']['status']}")
    print(f"Valid: {result['basic_validation']['is_valid']}")
    print(f"Confidence: {result['confidence_score']}")
    print(f"Rules applied: {len(result['applied_rules'])}")


if __name__ == "__main__":
    test_ticket_analyzer()
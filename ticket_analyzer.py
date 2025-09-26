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

    def analyze_ticket(self, ticket_info: dict, te_documents: dict, user_question: str = "") -> dict:
        """Analyse complète d'un ticket T&E avec IA directe (sans RAG)"""
        logger.info(f"Début analyse ticket: {ticket_info.get('filename', 'N/A')}")
        
        try:
            # Construire le contexte complet avec TOUTES les données
            full_context = self._build_complete_context(te_documents)

            # Construire le prompt d'analyse STRUCTURÉ avec raisonnement en étapes
            prompt = f"""
            QUESTION: {user_question if user_question else "Analyze this T&E expense according to company policies."}

            TICKET INFORMATION:
            - Amount: {ticket_info.get('amount', 'Not detected')} {ticket_info.get('currency', 'N/A')}
            - Category: {ticket_info.get('category', 'Unknown')}
            - Date: {ticket_info.get('date', 'Not detected')}
            - Vendor: {ticket_info.get('vendor', 'Not detected')}
            - Location: {ticket_info.get('location', 'Not detected')}

            ANALYSIS REQUIREMENTS:
            You must respond in this EXACT format:

            DECISION: [PASS or FAIL - exactly one word]
            EXPENSE_TYPE: [Hotel/Meal/Transport/etc.]
            AMOUNT: [NUM CURRENCY]
            ANALYSIS: [Your detailed analysis here]

            MANDATORY REASONING PROCESS:
            STEP 1 - CONSOLIDATED LIMITS CHECK:
            - Identify the exact country, currency, and expense type from the ticket
            - Search CONSOLIDATED LIMITS for matching rule (country + currency + type)
            - Compare ticket amount vs limit amount
            - State: "LIMITS CHECK: PASS/FAIL - [specific rule and amounts]"

            STEP 2 - TRAVEL & ENTERTAINMENT PROCEDURES CHECK:  
            - Search TRAVEL & ENTERTAINMENT PROCEDURES for relevant policies
            - Check if expense complies with general T&E procedures (timing, documentation, business purpose, etc.)
            - Quote the complete relevant policy section, not just the title
            - State: "PROCEDURES CHECK: PASS/FAIL - [specific policy quoted]"

            STEP 3 - FINAL DECISION:
            - DECISION = PASS only if BOTH Step 1 AND Step 2 are PASS
            - DECISION = FAIL if either Step 1 OR Step 2 is FAIL
            - If no applicable rule found in either source: DECISION = FAIL

            Rules:
            1. DECISION must be exactly "PASS" or "FAIL" - nothing else
            2. Follow the 3-step reasoning process above
            3. Quote exact amounts, currencies, and complete policy text
            4. Be specific about which rules apply from both sources
            5. Always check BOTH consolidated limits AND procedures
            6. If missing information prevents proper analysis: DECISION = FAIL

            Provide a professional business analysis following the 3-step process.
            """
            
            # Appel direct à l'IA avec contexte complet
            ai_response = self.llm_connector.get_llm_response(prompt, full_context)
            
            # Parser la réponse structurée
            parsed_response = self._parse_structured_response(ai_response)
            
            return {
                'result': parsed_response['decision'],
                'expense_type': parsed_response['expense_type'],
                'justification': parsed_response['analysis'],
                'comment': parsed_response['analysis'],
                'confidence_score': 0.95,
                'applied_rules': [],
                'timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            logger.error(f"Erreur lors de l'analyse du ticket: {str(e)}")
            return {
                'result': 'FAIL',
                'expense_type': 'Analysis Error',
                'justification': f'Technical error during analysis: {str(e)}',
                'comment': f"Error during analysis: {str(e)}",
                'confidence_score': 0.0,
                'timestamp': datetime.now().isoformat()
            }

    def _parse_structured_response(self, ai_response: str) -> dict:
        """Parse une réponse IA structurée avec montant"""
        import re
        
        # Patterns pour extraire chaque section
        decision_pattern = r'DECISION:\s*(PASS|FAIL)'
        expense_type_pattern = r'EXPENSE_TYPE:\s*([^\n]+)'
        amount_pattern = r'AMOUNT:\s*([^\n]+)'
        analysis_pattern = r'ANALYSIS:\s*(.*?)(?=\n\n|\Z)'
        
        # Extraction avec regex
        decision_match = re.search(decision_pattern, ai_response, re.IGNORECASE)
        expense_type_match = re.search(expense_type_pattern, ai_response, re.IGNORECASE)
        amount_match = re.search(amount_pattern, ai_response, re.IGNORECASE)
        analysis_match = re.search(analysis_pattern, ai_response, re.IGNORECASE | re.DOTALL)
        
        # Résultats avec fallbacks sécurisés
        decision = decision_match.group(1).upper() if decision_match else 'FAIL'
        expense_type = expense_type_match.group(1).strip() if expense_type_match else 'Unknown'
        amount = amount_match.group(1).strip() if amount_match else 'Not detected'
        analysis = analysis_match.group(1).strip() if analysis_match else ai_response
        
        # Validation finale
        if decision not in ['PASS', 'FAIL']:
            decision = 'FAIL'
        
        return {
            'decision': decision,
            'expense_type': expense_type,
            'amount': amount,
            'analysis': analysis
        }

    def _extract_expense_type(self, ticket_info: dict) -> str:
        """Extrait le type de dépense de manière simple"""
        category = ticket_info.get('category', 'unknown').lower()
        
        mapping = {
            'hotel': 'Hotel',
            'accommodation': 'Hotel', 
            'meal': 'Meal',
            'restaurant': 'Meal',
            'breakfast': 'Breakfast',
            'lunch': 'Meal',
            'dinner': 'Meal',
            'transport': 'Transport',
            'flight': 'Flight'
        }
        
        return mapping.get(category, 'Unknown')

    def answer_general_question(self, user_question: str, te_documents: dict) -> dict:
        """Répond à une question générale avec contexte complet"""
        try:
            # Contexte complet pour question générale
            full_context = self._build_complete_context(te_documents)
            
            prompt = f"""
            QUESTION: {user_question}
            === MANDATORY INSTRUCTIONS ===
            1. ALWAYS extract and quote exact content from COMPLETE T&E RULES AND POLICIES - never just reference sections
            2. For amount/price/expense queries: identify exact country(or continent or region)/currency/type and provide the specific limit amount
            3. For policy questions: extract and explain the actual content from relevant sections
            4. NEVER say 'refer to section X' without explaining what that section contains
            5. Give precise answers with specific numbers, currencies, and details
            6. Search through ALL data in the context to find relevant information
            """
            ai_response = self.llm_connector.get_llm_response(prompt, full_context)
            
            return {
                'ai_response': ai_response,
                'question_type': 'general_policy',
                'timestamp': datetime.now().isoformat()
            }
            
        except Exception as e:
            logger.error(f"Erreur réponse question générale: {str(e)}")
            return {
                'ai_response': f"Error processing your question: {str(e)}",
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

    def _ai_extract_ticket_info(self, raw_text: str, filename: str) -> dict:
        """Extrait TOUTES les informations du ticket via IA + RAG"""
        try:
            # Construire le prompt d'extraction
            extraction_prompt = f"""
            Parse this ticket and extract information in VALID JSON format with double quotes only:
            {{
                "amount": 57.0,
                "currency": "EUR", 
                "category": "meal",
                "subcategory": "dinner",
                "date": "2025-09-12",
                "vendor": "pizzeria",
                "location": "Rome", 
                "country_code": "IT",
                "confidence": 0.8
            }}

            IMPORTANT: Use double quotes for strings, not single quotes.
            Return ONLY the JSON, no extra text.
            
            Ticket to analyze :
            {raw_text[:2000]}
            """
            
            # Obtenir le contexte des règles T&E
            context = "You are an expert in T&E analysis. You are given a ticket/raw text and your task is to extract precise information from it. You must return the extracted information in JSON format. You must return the JSON only, no context or additional text and information."

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

    def ai_extract_ticket_info(self, raw_text: str, filename: str) -> dict:
        """Méthode publique pour extraction IA"""
        return self._ai_extract_ticket_info(raw_text, filename)

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
            "subcategory": None,
            "date": None,
            "vendor": None,
            "location": None,
            "country_code": None,
            "confidence": 0.5
        }
        
        # Extraire tous les champs avec regex
        patterns = {
            "amount": r'"amount":\s*([0-9.]+)',
            "currency": r'"currency":\s*[\'"]([A-Z]{3})[\'"]',
            "category": r'"category":\s*[\'"]([^\'",]+)[\'"]',
            "subcategory": r'"subcategory":\s*[\'"]([^\'",]+)[\'"]',
            "date": r'"date":\s*[\'"]([^\'",]+)[\'"]', 
            "vendor": r'"vendor":\s*[\'"]([^\'",]+)[\'"]',
            "location": r'"location":\s*[\'"]([^\'",]+)[\'"]',
            "country_code": r'"country_code":\s*[\'"]([^\'",]+)[\'"]',
            "confidence": r'"confidence":\s*[\'"]?([0-9.]+)[\'"]?'
        }
        
        for field, pattern in patterns.items():
            match = re.search(pattern, ai_response)
            if match:
                value = match.group(1)
                if field in ["amount", "confidence"]:
                    try:
                        info[field] = float(value)
                    except ValueError:
                        pass
                else:
                    info[field] = value
        
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

    def _build_complete_context(self, te_documents: dict) -> str:
        """Construit un contexte complet avec TOUTES les données T&E"""        
        context_parts = []
        
        context_parts.append("=== COMPLETE T&E RULES AND POLICIES ===")
        context_parts.append("You are a T&E expense analysis expert. You have to access to complete TRAVEL & ENTERTAINMENT PROCEDURES and CONSOLIDATED LIMITS below.")
        context_parts.append("")

        
        # TOUTES les règles Excel - 3 sheets complètes
        if te_documents.get("excel_rules"):
            context_parts.append("CONSOLIDATED LIMITS:")

            # Define a mapping for the rule types
            type_mapping = {
                "Hotel1": "Hotel",
                "Breakfast1": "Breakfast",
                "Meal1": "Meal"
            }
            
            for sheet_name, rules in te_documents["excel_rules"].items():
                if sheet_name == "Breakfast & Lunch & Dinner":
                    sheet_name = "Internal Staff Meal"
                context_parts.append(f"\n{sheet_name} Sheet:")
                context_parts.append(f"Total rules: {len(rules)}")
                
                for rule in rules:
                    rule_type = type_mapping.get(rule.get('type', 'N/A'), rule.get('type', 'N/A'))
                    rule_line = f"- For a travel in {rule.get('country', 'N/A')} with a expense for {rule_type}, the limit is {rule.get('amount_limit', 0)} {rule.get('currency', 'N/A')}\n"
                    context_parts.append(rule_line)
            
            context_parts.append("")
        
        # TOUT le document Word - 16 pages complètes
        if te_documents.get("word_policies"):
            context_parts.append("TRAVEL & ENTERTAINMENT PROCEDURES:")
            context_parts.append(te_documents["word_policies"])
            context_parts.append("")
        
        context_parts.append("=== END OF COMPLETE T&E DATA ===")
        
        return "\n".join(context_parts)
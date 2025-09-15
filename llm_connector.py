# llm_connector.py
"""
Simulateur LLM pour le développement
==================================

Fournit des réponses simulées pour tester l'interface
sans nécessiter une vraie API IA
"""

import random
import time
from datetime import datetime

class LLMConnector:
    def __init__(self):
        """Initialise le connecteur LLM simulé"""
        self.response_templates = {
            "travel_expense": [
                "Based on the T&E policies, this expense appears to be **valid**. The amount of {amount} for {category} in {location} is within the approved limits for your grade level.",
                "I found a potential issue with this expense. The amount of {amount} exceeds the maximum allowed for {category} in {location} according to section 3.2 of the T&E policy.",
                "This expense requires additional documentation. While the amount is acceptable, please provide the original receipt as per policy section 4.1.",
            ],
            "policy_question": [
                "According to the T&E policy document, the reimbursement rules for {topic} are as follows: Maximum amounts vary by country and employee grade. For specific limits, please refer to the Excel sheet column AMOUNT1.",
                "The policy states that {topic} expenses must be approved in advance for amounts exceeding the threshold. Check the ID_01 column in your Excel file for country-specific rules.",
                "Based on the current T&E guidelines, {topic} is covered under the following conditions: Pre-approval required for amounts over the standard limit, original receipts mandatory.",
            ],
            "general": [
                "I understand you're asking about {topic}. Let me analyze the T&E documents to provide you with accurate information.",
                "Based on the uploaded T&E policies, I can help you understand the rules for {topic}. Here's what I found in the documentation:",
                "Your question about {topic} relates to several sections of the T&E policy. Let me break down the relevant rules for you:",
            ]
        }
        
        self.policy_facts = [
            "Hotel expenses are limited based on CRN_KEY (currency) and ID_01 (country code)",
            "Meal allowances vary by location and are categorized as Breakfast, Lunch, or Dinner", 
            "Internal staff meals have different limits than external business meals",
            "All expenses must have original receipts for amounts above 25 EUR/USD equivalent",
            "Pre-approval is required for hotel stays exceeding 200 EUR per night",
            "Taxi receipts are mandatory for distances over 10km or amounts over 50 EUR"
        ]
        
    def get_llm_response(self, user_prompt, context_prompt="", modelID="gpt-4o-mini-2024-07-18", temperature=0.1):
        """
        Simule une réponse IA basée sur le prompt utilisateur
        
        Args:
            user_prompt: Question de l'utilisateur
            context_prompt: Contexte des documents T&E
            modelID: ID du modèle (ignoré en simulation)
            temperature: Température (ignorée en simulation)
        
        Returns:
            str: Réponse simulée contextuelle
        """
        # Simulation du délai de traitement
        time.sleep(random.uniform(1, 3))
        
        # Analyse du type de question
        question_type = self._analyze_question_type(user_prompt)
        
        # Extraction d'informations du prompt
        extracted_info = self._extract_info_from_prompt(user_prompt, context_prompt)
        
        # Génération de la réponse
        response = self._generate_contextual_response(question_type, extracted_info, user_prompt)
        
        return response
    
    def _analyze_question_type(self, prompt):
        """Détermine le type de question posée"""
        prompt_lower = prompt.lower()
        
        # Mots-clés pour validation de tickets
        expense_keywords = ["ticket", "reçu", "facture", "expense", "receipt", "invoice", "valid", "approve"]
        if any(keyword in prompt_lower for keyword in expense_keywords):
            return "travel_expense"
        
        # Mots-clés pour questions de politique
        policy_keywords = ["policy", "politique", "règle", "rule", "limit", "limite", "allowance", "remboursement"]
        if any(keyword in prompt_lower for keyword in policy_keywords):
            return "policy_question"
            
        return "general"
    
    def _extract_info_from_prompt(self, user_prompt, context_prompt):
        """Extrait des informations pertinentes des prompts"""
        info = {
            "amount": None,
            "currency": None,
            "category": None,
            "location": None,
            "topic": None
        }
        
        # Extraction d'informations numériques (montants)
        import re
        amounts = re.findall(r'(\d+(?:\.\d{2})?)\s*(EUR|USD|AED|CHF|AUD)', user_prompt)
        if amounts:
            info["amount"] = f"{amounts[0][0]} {amounts[0][1]}"
            info["currency"] = amounts[0][1]
        
        # Extraction de catégories
        categories = ["hotel", "meal", "taxi", "transport", "breakfast", "lunch", "dinner"]
        for category in categories:
            if category in user_prompt.lower():
                info["category"] = category
                break
        
        # Extraction de lieux (codes pays)
        countries = ["FR", "DE", "UK", "US", "AE", "AU", "CH", "BE"]
        for country in countries:
            if country in user_prompt.upper():
                info["location"] = country
                break
        
        # Extraction du sujet principal
        if not info["category"]:
            # Essayer d'extraire le sujet principal
            words = user_prompt.split()
            if len(words) > 3:
                info["topic"] = " ".join(words[:3])
            else:
                info["topic"] = "travel expenses"
        else:
            info["topic"] = info["category"]
        
        return info
    
    def _generate_contextual_response(self, question_type, info, original_prompt):
        """Génère une réponse contextuelle basée sur l'analyse"""
        
        # Sélectionner un template approprié
        templates = self.response_templates.get(question_type, self.response_templates["general"])
        base_response = random.choice(templates)
        
        # Formater avec les informations extraites
        try:
            formatted_response = base_response.format(
                amount=info.get("amount", "the specified amount"),
                currency=info.get("currency", "EUR"),
                category=info.get("category", "this expense category"),
                location=info.get("location", "this location"),
                topic=info.get("topic", "your request")
            )
        except KeyError:
            # Fallback si le formatage échoue
            formatted_response = base_response.replace("{topic}", info.get("topic", "your request"))
        
        # Ajouter des faits pertinents sur les politiques
        if question_type in ["travel_expense", "policy_question"]:
            additional_facts = random.sample(self.policy_facts, min(2, len(self.policy_facts)))
            formatted_response += "\n\n**Relevant policy information:**\n"
            for fact in additional_facts:
                formatted_response += f"- {fact}\n"
        
        # Ajouter une référence aux documents sources
        formatted_response += "\n\n*Source: T&E Policy Documents (Excel sheets: Internal staff Meal, Hotel, Breakfast & Lunch & Dinner)*"
        
        # Ajouter une conclusion appropriée
        if question_type == "travel_expense":
            if "valid" in formatted_response or "acceptable" in formatted_response:
                formatted_response += "\n\n✅ **Recommendation:** Expense can be approved for reimbursement."
            else:
                formatted_response += "\n\n⚠️ **Recommendation:** Please review this expense before approval."
        
        return formatted_response
    
    def _simulate_document_analysis(self, ticket_info, te_rules):
        """
        Simule l'analyse d'un ticket contre les règles T&E
        
        Args:
            ticket_info: Informations extraites du ticket
            te_rules: Règles T&E des documents Excel/Word
        
        Returns:
            dict: Résultat de l'analyse
        """
        # Cette méthode pourra être utilisée plus tard pour une analyse plus poussée
        return {
            "is_valid": random.choice([True, False]),
            "confidence": random.uniform(0.7, 0.95),
            "issues": random.choice([[], ["Missing receipt"], ["Amount exceeds limit"]]),
            "recommendations": ["Provide original receipt", "Get manager approval"]
        }

# Fonction utilitaire pour tester le connecteur
def test_llm_connector():
    """Teste le connecteur LLM avec différents types de questions"""
    connector = LLMConnector()
    
    test_prompts = [
        "I have a hotel receipt for 150 EUR in Paris. Is this valid?",
        "What are the meal allowance limits for Germany?", 
        "Can you explain the travel expense policy?",
        "This taxi receipt is for 45 USD in New York, can I claim it?"
    ]
    
    print("=== Test du LLM Connector ===\n")
    
    for i, prompt in enumerate(test_prompts, 1):
        print(f"Test {i}: {prompt}")
        response = connector.get_llm_response(prompt)
        print(f"Réponse: {response}\n")
        print("-" * 50 + "\n")

if __name__ == "__main__":
    test_llm_connector()
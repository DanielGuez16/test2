# context_builder.py
"""
Context Builder for T&E Analysis
================================

Prépare les contextes enrichis pour l'analyse IA des tickets T&E
en utilisant les données RAG et les informations structurées.
"""

from typing import Dict, List, Any
import json
from datetime import datetime

class TEContextBuilder:
    """Constructeur de contexte pour l'analyse T&E"""
    
    def __init__(self):
        self.context_templates = {
            'ticket_analysis': self._build_ticket_analysis_context,
            'general_query': self._build_general_query_context,
            'policy_question': self._build_policy_question_context
        }
    
    def build_context(self, context_type: str, **kwargs) -> str:
        """
        Construit un contexte selon le type demandé
        
        Args:
            context_type: Type de contexte ('ticket_analysis', 'general_query', 'policy_question')
            **kwargs: Données spécifiques au type de contexte
            
        Returns:
            str: Contexte formaté pour l'IA
        """
        if context_type not in self.context_templates:
            raise ValueError(f"Type de contexte non supporté: {context_type}")
        
        return self.context_templates[context_type](**kwargs)
    
    def _build_ticket_analysis_context(self, ticket_info: dict, relevant_rules: List[Dict], 
                                     policies_context: str = "", **kwargs) -> str:
        """Construit le contexte pour l'analyse d'un ticket"""
        context_parts = []
        
        # En-tête du contexte
        context_parts.append("=== ANALYSE TICKET T&E - RÉGION APAC ===")
        context_parts.append(f"Date d'analyse: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        context_parts.append("")
        
        # Informations du ticket
        context_parts.append("TICKET À ANALYSER:")
        context_parts.append(f"- Fichier: {ticket_info.get('filename', 'N/A')}")
        context_parts.append(f"- Montant: {ticket_info.get('amount', 'N/A')} {ticket_info.get('currency', 'N/A')}")
        context_parts.append(f"- Catégorie détectée: {ticket_info.get('category', 'N/A')}")
        context_parts.append(f"- Date: {ticket_info.get('date', 'N/A')}")
        context_parts.append(f"- Lieu/Pays: {ticket_info.get('location', 'N/A')}")
        context_parts.append(f"- Établissement: {ticket_info.get('vendor', 'N/A')}")
        context_parts.append("")
        
        # Règles applicables trouvées par RAG
        if relevant_rules:
            context_parts.append("RÈGLES T&E APPLICABLES:")
            for rule in relevant_rules:
                rule_desc = f"- {rule.get('type', 'N/A')} en {rule.get('country', 'N/A')} "
                rule_desc += f"({rule.get('currency', 'N/A')}): limite {rule.get('amount_limit', 0)}"
                
                # Ajouter sheet d'origine si disponible
                if 'sheet_name' in rule:
                    rule_desc += f" [Source: {rule['sheet_name']}]"
                
                context_parts.append(rule_desc)
            context_parts.append("")
        else:
            context_parts.append("ATTENTION: Aucune règle spécifique trouvée pour ce ticket")
            context_parts.append("")
        
        # Contexte des politiques Word
        if policies_context:
            context_parts.append("EXTRAITS DES POLITIQUES T&E:")
            context_parts.append(policies_context[:800] + "..." if len(policies_context) > 800 else policies_context)
            context_parts.append("")
        
        # Instructions pour l'IA
        context_parts.append("INSTRUCTIONS D'ANALYSE:")
        context_parts.append("1. Vérifier la conformité du montant avec les limites autorisées")
        context_parts.append("2. Valider la catégorie de dépense et le pays")
        context_parts.append("3. Identifier les documents manquants ou problèmes potentiels")
        context_parts.append("4. Fournir une recommandation claire (Approuvé/Rejeté/En attente)")
        context_parts.append("5. Justifier la décision en citant les règles applicables")
        
        return "\n".join(context_parts)
    
    def _build_general_query_context(self, te_rules_summary: dict, policies_excerpt: str = "", **kwargs) -> str:
        """Construit le contexte pour une question générale sur les politiques T&E"""
        context_parts = []
        
        context_parts.append("=== ASSISTANT T&E - RÉGION APAC ===")
        context_parts.append("")
        
        # Résumé des règles disponibles
        if te_rules_summary:
            context_parts.append("RÈGLES T&E DISPONIBLES:")
            for sheet_name, rules_count in te_rules_summary.items():
                context_parts.append(f"- {sheet_name}: {rules_count} règles")
            
            # Devises et pays couverts
            if 'currencies' in kwargs:
                context_parts.append(f"- Devises couvertes: {', '.join(kwargs['currencies'])}")
            if 'countries' in kwargs:
                context_parts.append(f"- Pays couverts: {', '.join(kwargs['countries'])}")
            context_parts.append("")
        
        # Extrait des politiques
        if policies_excerpt:
            context_parts.append("EXTRAITS DES POLITIQUES:")
            context_parts.append(policies_excerpt)
            context_parts.append("")
        
        context_parts.append("Tu es un assistant spécialisé en politiques T&E pour la région APAC.")
        context_parts.append("Réponds de manière précise en te basant sur les règles et politiques fournies.")
        
        return "\n".join(context_parts)
    
    def _build_policy_question_context(self, question_topic: str, relevant_policies: str, 
                                     related_rules: List[Dict] = None, **kwargs) -> str:
        """Construit le contexte pour une question spécifique sur les politiques"""
        context_parts = []
        
        context_parts.append(f"=== QUESTION POLITIQUE T&E: {question_topic.upper()} ===")
        context_parts.append("")
        
        # Politiques pertinentes
        context_parts.append("POLITIQUES PERTINENTES:")
        context_parts.append(relevant_policies)
        context_parts.append("")
        
        # Règles associées si disponibles
        if related_rules:
            context_parts.append("RÈGLES ASSOCIÉES:")
            for rule in related_rules:
                context_parts.append(f"- {rule.get('type', 'N/A')}: {rule.get('amount_limit', 0)} {rule.get('currency', 'N/A')} en {rule.get('country', 'N/A')}")
            context_parts.append("")
        
        context_parts.append("Réponds de manière détaillée en expliquant les règles et leurs applications pratiques.")
        
        return "\n".join(context_parts)
    
    def build_prompt_for_ticket_analysis(self, user_question: str, ticket_info: dict) -> str:
        """Construit le prompt utilisateur pour l'analyse de ticket"""
        if user_question.strip():
            prompt = f"Question spécifique: {user_question}\n\n"
        else:
            prompt = "Analyse ce ticket T&E selon les politiques de l'entreprise.\n\n"
        
        prompt += "INFORMATIONS DU TICKET:\n"
        prompt += f"- Montant: {ticket_info.get('amount', 'Non détecté')} {ticket_info.get('currency', 'N/A')}\n"
        prompt += f"- Catégorie: {ticket_info.get('category', 'Non déterminée')}\n"
        prompt += f"- Date: {ticket_info.get('date', 'Non détectée')}\n"
        prompt += f"- Établissement: {ticket_info.get('vendor', 'Non détecté')}\n\n"
        
        prompt += "Merci de fournir:\n"
        prompt += "1. Statut de validation (Approuvé/Rejeté/Révision nécessaire)\n"
        prompt += "2. Justification basée sur les règles T&E\n"
        prompt += "3. Actions recommandées si nécessaire"
        
        return prompt
    
    def build_prompt_for_general_query(self, user_question: str) -> str:
        """Construit le prompt utilisateur pour une question générale"""
        return f"Question sur les politiques T&E: {user_question}\n\nMerci de répondre en te basant sur les règles et politiques fournies."


def test_context_builder():
    """Teste le constructeur de contexte"""
    builder = TEContextBuilder()
    
    # Test contexte analyse ticket
    ticket_info = {
        'filename': 'receipt_hotel_paris.pdf',
        'amount': 180,
        'currency': 'EUR',
        'category': 'hotel',
        'date': '2024-03-15',
        'vendor': 'Hotel Le Marais'
    }
    
    relevant_rules = [{
        'type': 'Hotel1',
        'country': 'FR',
        'currency': 'EUR',
        'amount_limit': 200,
        'sheet_name': 'Hotel'
    }]
    
    context = builder.build_context(
        'ticket_analysis',
        ticket_info=ticket_info,
        relevant_rules=relevant_rules,
        policies_context="Les hôtels doivent être pré-approuvés pour les séjours > 3 nuits..."
    )
    
    print("=== TEST CONTEXT BUILDER ===")
    print(context[:500] + "...")
    
    # Test prompt
    prompt = builder.build_prompt_for_ticket_analysis(
        "Ce montant est-il dans les limites autorisées?",
        ticket_info
    )
    
    print(f"\nPROMPT: {prompt}")


if __name__ == "__main__":
    test_context_builder()
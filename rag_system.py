import re
import logging
from typing import List, Dict, Any, Optional
from sentence_transformers import SentenceTransformer
import chromadb
from chromadb.config import Settings
import hashlib

logger = logging.getLogger(__name__)

class TERAGSystem:
    def __init__(self):
        # Initialiser le modèle d'embedding
        try:
            self.model = SentenceTransformer('all-MiniLM-L6-v2')
            logger.info("Modèle d'embedding initialisé avec succès")
        except Exception as e:
            logger.warning(f"Erreur chargement modèle embedding: {e}, utilisation fallback")
            self.model = None
        
        # Initialiser ChromaDB en mode persistant local
        try:
            self.chroma_client = chromadb.PersistentClient(path="./data/chroma_db")
            logger.info("ChromaDB initialisé en mode persistant")
        except Exception as e:
            logger.warning(f"Erreur ChromaDB: {e}, utilisation mode mémoire")
            self.chroma_client = chromadb.Client()
        
        self.rules_collection = None
        self.policies_collection = None
        
        # Index en mémoire pour recherche rapide
        self.rules_index = {}
        self.policies_chunks = []
    
    def index_excel_rules(self, rules_data: Dict[str, List[Dict]]):
        """Indexe les règles Excel pour recherche vectorielle et structurée"""
        try:
            logger.info("Début indexation des règles Excel")
            
            # Créer ou récupérer la collection pour les règles
            collection_name = "te_excel_rules"
            try:
                self.rules_collection = self.chroma_client.create_collection(
                    name=collection_name,
                    metadata={"hnsw:space": "cosine"}
                )
            except Exception:
                # Collection existe déjà
                self.rules_collection = self.chroma_client.get_collection(collection_name)
                self.rules_collection.delete()  # Nettoyer pour réindexer
                self.rules_collection = self.chroma_client.create_collection(
                    name=collection_name,
                    metadata={"hnsw:space": "cosine"}
                )
            
            # Préparer les documents pour l'indexation vectorielle
            documents = []
            metadatas = []
            ids = []
            
            rule_id = 0
            for sheet_name, rules in rules_data.items():
                for rule in rules:
                    # Créer un document textuel pour la recherche vectorielle
                    doc_text = self._create_rule_document(rule, sheet_name)
                    documents.append(doc_text)
                    
                    # Métadonnées structurées pour filtrage
                    metadata = {
                        "sheet_name": sheet_name,
                        "currency": rule.get("currency", ""),
                        "country": rule.get("country", ""),
                        "type": rule.get("type", ""),
                        "amount_limit": rule.get("amount_limit", 0)
                    }
                    metadatas.append(metadata)
                    ids.append(f"rule_{rule_id}")
                    
                    # Index rapide pour recherche directe
                    key = f"{rule.get('currency', '')}_{rule.get('country', '')}_{rule.get('type', '')}"
                    self.rules_index[key] = rule
                    
                    rule_id += 1
            
            # Indexer dans ChromaDB si embeddings disponibles
            if self.model and documents:
                try:
                    embeddings = self.model.encode(documents).tolist()
                    self.rules_collection.add(
                        embeddings=embeddings,
                        documents=documents,
                        metadatas=metadatas,
                        ids=ids
                    )
                    logger.info(f"Indexé {len(documents)} règles dans ChromaDB")
                except Exception as e:
                    logger.warning(f"Erreur indexation ChromaDB: {e}")
            
            logger.info(f"Indexation Excel terminée: {len(documents)} règles indexées")
            
        except Exception as e:
            logger.error(f"Erreur indexation règles Excel: {e}")
    
    def index_word_policies(self, policies_text: str):
        """Indexe les politiques Word par sections"""
        try:
            logger.info("Début indexation des politiques Word")
            
            # Découper le texte en sections logiques
            sections = self._chunk_policies_text(policies_text)
            
            # Créer ou récupérer la collection pour les politiques
            collection_name = "te_word_policies"
            try:
                self.policies_collection = self.chroma_client.create_collection(
                    name=collection_name,
                    metadata={"hnsw:space": "cosine"}
                )
            except Exception:
                self.policies_collection = self.chroma_client.get_collection(collection_name)
                self.policies_collection.delete()
                self.policies_collection = self.chroma_client.create_collection(
                    name=collection_name,
                    metadata={"hnsw:space": "cosine"}
                )
            
            # Indexer les sections si embeddings disponibles
            if self.model and sections:
                try:
                    documents = [section["text"] for section in sections]
                    metadatas = [{"section_type": section["type"], "keywords": section["keywords"]} for section in sections]
                    ids = [f"policy_{i}" for i in range(len(sections))]
                    
                    embeddings = self.model.encode(documents).tolist()
                    self.policies_collection.add(
                        embeddings=embeddings,
                        documents=documents,
                        metadatas=metadatas,
                        ids=ids
                    )
                    logger.info(f"Indexé {len(sections)} sections de politiques")
                except Exception as e:
                    logger.warning(f"Erreur indexation politiques: {e}")
            
            # Stocker en mémoire pour fallback
            self.policies_chunks = sections
            
        except Exception as e:
            logger.error(f"Erreur indexation politiques: {e}")
    
    def search_relevant_rules(self, query: str, filters: Dict = None) -> List[Dict]:
        """Recherche les règles pertinentes pour une requête"""
        try:
            relevant_rules = []
            
            # 1. Recherche directe par critères exacts
            if filters:
                exact_rules = self._search_exact_rules(filters)
                relevant_rules.extend(exact_rules)
            
            # 2. Recherche vectorielle si disponible
            if self.model and self.rules_collection:
                try:
                    vector_rules = self._search_vector_rules(query, filters)
                    relevant_rules.extend(vector_rules)
                except Exception as e:
                    logger.warning(f"Erreur recherche vectorielle: {e}")
            
            # 3. Recherche par mots-clés en fallback
            if not relevant_rules:
                keyword_rules = self._search_keyword_rules(query, filters)
                relevant_rules.extend(keyword_rules)
            
            # Dédoublonner et trier par pertinence
            unique_rules = self._deduplicate_rules(relevant_rules)
            
            logger.info(f"Trouvé {len(unique_rules)} règles pour: {query}")
            return unique_rules[:10]  # Limiter à 10 résultats
            
        except Exception as e:
            logger.error(f"Erreur recherche règles: {e}")
            return []
    
    def search_policies(self, topic: str) -> str:
        """Recherche dans les politiques par sujet"""
        try:
            relevant_sections = []
            
            # Recherche vectorielle si disponible
            if self.model and self.policies_collection:
                try:
                    results = self.policies_collection.query(
                        query_embeddings=[self.model.encode([topic]).tolist()[0]],
                        n_results=3
                    )
                    relevant_sections.extend(results["documents"][0])
                except Exception as e:
                    logger.warning(f"Erreur recherche vectorielle politiques: {e}")
            
            # Recherche par mots-clés en fallback
            if not relevant_sections:
                for chunk in self.policies_chunks:
                    if any(keyword.lower() in topic.lower() for keyword in chunk.get("keywords", [])):
                        relevant_sections.append(chunk["text"])
            
            return "\n\n".join(relevant_sections[:3]) if relevant_sections else "Aucune politique spécifique trouvée."
            
        except Exception as e:
            logger.error(f"Erreur recherche politiques: {e}")
            return "Erreur lors de la recherche dans les politiques."
    
    def _create_rule_document(self, rule: Dict, sheet_name: str) -> str:
        """Crée un document textuel pour une règle"""
        doc_parts = [
            f"Sheet: {sheet_name}",
            f"Type de dépense: {rule.get('type', 'Non spécifié')}",
            f"Pays: {rule.get('country', 'Non spécifié')}",
            f"Devise: {rule.get('currency', 'Non spécifié')}",
            f"Limite: {rule.get('amount_limit', 0)} {rule.get('currency', '')}"
        ]
        return " | ".join(doc_parts)
    
    def _chunk_policies_text(self, text: str) -> List[Dict]:
        """Découpe le texte des politiques en sections"""
        sections = []
        
        # Découpage par paragraphes
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
        
        for i, paragraph in enumerate(paragraphs):
            if len(paragraph) > 50:  # Ignorer les paragraphes trop courts
                # Extraire des mots-clés
                keywords = self._extract_keywords(paragraph)
                
                # Déterminer le type de section
                section_type = self._classify_section(paragraph)
                
                sections.append({
                    "text": paragraph,
                    "type": section_type,
                    "keywords": keywords,
                    "index": i
                })
        
        return sections
    
    def _extract_keywords(self, text: str) -> List[str]:
        """Extrait des mots-clés d'un texte"""
        keywords = []
        text_lower = text.lower()
        
        # Mots-clés T&E courants
        te_keywords = [
            "hotel", "meal", "transport", "flight", "taxi", "restaurant",
            "breakfast", "lunch", "dinner", "accommodation", "expense",
            "receipt", "approval", "limit", "policy", "reimbursement"
        ]
        
        for keyword in te_keywords:
            if keyword in text_lower:
                keywords.append(keyword)
        
        return keywords
    
    def _classify_section(self, text: str) -> str:
        """Classifie une section de texte"""
        text_lower = text.lower()
        
        if any(word in text_lower for word in ["hotel", "accommodation", "lodging"]):
            return "hotel_policy"
        elif any(word in text_lower for word in ["meal", "restaurant", "food", "dining"]):
            return "meal_policy"
        elif any(word in text_lower for word in ["transport", "taxi", "flight", "travel"]):
            return "transport_policy"
        elif any(word in text_lower for word in ["receipt", "documentation", "proof"]):
            return "documentation_policy"
        else:
            return "general_policy"
    
    def _search_exact_rules(self, filters: Dict) -> List[Dict]:
        """Recherche exacte par critères"""
        rules = []
        
        currency = filters.get("currency", "")
        country = filters.get("country", "")
        expense_type = filters.get("expense_type", "")
        
        # Recherche directe dans l'index
        key = f"{currency}_{country}_{expense_type}"
        if key in self.rules_index:
            rules.append(self.rules_index[key])
        
        # Recherche partielle
        for indexed_key, rule in self.rules_index.items():
            parts = indexed_key.split("_")
            if len(parts) >= 3:
                if (currency and parts[0] == currency) or \
                   (country and parts[1] == country) or \
                   (expense_type and parts[2] == expense_type):
                    if rule not in rules:
                        rules.append(rule)
        
        return rules
    
    def _search_vector_rules(self, query: str, filters: Dict = None) -> List[Dict]:
        """Recherche vectorielle dans les règles"""
        rules = []
        
        try:
            # Construire les filtres ChromaDB
            where_clause = {}
            if filters:
                if filters.get("currency"):
                    where_clause["currency"] = filters["currency"]
                if filters.get("country"):
                    where_clause["country"] = filters["country"]
            
            # Recherche vectorielle
            results = self.rules_collection.query(
                query_embeddings=[self.model.encode([query]).tolist()[0]],
                n_results=5,
                where=where_clause if where_clause else None
            )
            
            # Convertir les résultats
            for metadata in results["metadatas"][0]:
                rules.append({
                    "sheet_name": metadata.get("sheet_name"),
                    "currency": metadata.get("currency"),
                    "country": metadata.get("country"),
                    "type": metadata.get("type"),
                    "amount_limit": metadata.get("amount_limit")
                })
        
        except Exception as e:
            logger.warning(f"Erreur recherche vectorielle: {e}")
        
        return rules
    
    def _search_keyword_rules(self, query: str, filters: Dict = None) -> List[Dict]:
        """Recherche par mots-clés en fallback"""
        rules = []
        query_lower = query.lower()
        
        for rule in self.rules_index.values():
            # Vérifier correspondance avec les filtres
            if filters:
                if filters.get("currency") and rule.get("currency") != filters["currency"]:
                    continue
                if filters.get("country") and rule.get("country") != filters["country"]:
                    continue
            
            # Vérifier correspondance avec la requête
            rule_text = f"{rule.get('type', '')} {rule.get('currency', '')} {rule.get('country', '')}".lower()
            if any(word in rule_text for word in query_lower.split()):
                rules.append(rule)
        
        return rules
    
    def _deduplicate_rules(self, rules: List[Dict]) -> List[Dict]:
        """Supprime les doublons dans les règles"""
        seen = set()
        unique_rules = []
        
        for rule in rules:
            rule_signature = f"{rule.get('currency')}_{rule.get('country')}_{rule.get('type')}_{rule.get('amount_limit')}"
            if rule_signature not in seen:
                seen.add(rule_signature)
                unique_rules.append(rule)
        
        return unique_rules
    
    def get_stats(self) -> Dict:
        """Retourne les statistiques du système RAG"""
        return {
            "rules_indexed": len(self.rules_index),
            "policies_chunks": len(self.policies_chunks),
            "embedding_model_available": self.model is not None,
            "chroma_available": self.rules_collection is not None
        }
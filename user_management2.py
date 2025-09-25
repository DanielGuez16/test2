# user_management.py - Version Clean
"""
User Management System - Version simplifiée et robuste
====================================================
Gestion des utilisateurs, logs, analyses et feedbacks avec SharePoint uniquement.
"""

from typing import Dict, List, Optional
from datetime import datetime
import hashlib
import logging
import pandas as pd
from io import BytesIO
from sharepoint_connector import SharePointClient

# Configuration des chemins SharePoint
SHAREPOINT_LOGS_PATH = "Chatbot/logs/activity_logs.xlsx"
SHAREPOINT_ANALYSIS_PATH = "Chatbot/logs/analysis_history.xlsx" 
SHAREPOINT_FEEDBACK_PATH = "Chatbot/logs/feedback_data.xlsx"

# Base de données utilisateurs (gardée comme demandé)
USERS_DB = {
    "daniel.guez@natixis.com": {
        "username": "daniel.guez@natixis.com",
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
        "full_name": "Daniel GUEZ",
        "role": "admin",
        "created_at": "2024-01-01T00:00:00"
    },
    "franck.pokou-ext@natixis.com": {
        "username": "franck.pokou-ext@natixis.com", 
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
        "full_name": "Franck POKOU",
        "role": "admin",
        "created_at": "2024-01-01T00:00:00"
    },
    "juvenalamos.ido@natixis.com": {
        "username": "juvenalamos.ido@natixis.com",
        "password_hash": hashlib.sha256("admin123".encode()).hexdigest(), 
        "full_name": "Juvenal Amos IDO",
        "role": "admin",
        "created_at": "2024-01-01T00:00:00"
    },
    "user.te@natixis.com": {
        "username": "user.te@natixis.com",
        "password_hash": hashlib.sha256("user123".encode()).hexdigest(),
        "full_name": "User ALM",
        "role": "user",
        "created_at": "2024-01-01T00:00:00"
    }
}

logger = logging.getLogger(__name__)

def authenticate_user(username: str, password: str) -> Optional[Dict]:
    """Authentifie un utilisateur"""
    user = USERS_DB.get(username)
    if user:
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        if user["password_hash"] == password_hash:
            return user
    return None

# ============================================================================
# 1. GESTION DES LOGS D'ACTIVITÉ
# ============================================================================

def log_activity(username: str, action: str, details: str = ""):
    """
    Enregistre une activité utilisateur dans SharePoint
    
    Args:
        username: Nom d'utilisateur
        action: Type d'action (LOGIN, LOGOUT, ANALYSIS, etc.)
        details: Détails de l'action
    """
    try:
        client = SharePointClient()
        
        # Créer l'entrée de log
        log_entry = {
            "timestamp": datetime.now().isoformat(),
            "username": username,
            "action": action,
            "details": details
        }
        
        # Lire le fichier existant ou créer un DataFrame vide
        try:
            binary_content = client.read_binary_file(SHAREPOINT_LOGS_PATH)
            logs_df = pd.read_excel(BytesIO(binary_content))
            logs_df = logs_df.fillna("")
            
            # Vérifier les colonnes
            expected_columns = ["timestamp", "username", "action", "details"]
            if list(logs_df.columns) != expected_columns:
                logger.warning("Colonnes incorrectes dans le fichier logs, recréation")
                logs_df = pd.DataFrame(columns=expected_columns)
                
        except Exception as e:
            logger.info(f"Création nouveau fichier logs: {e}")
            logs_df = pd.DataFrame(columns=["timestamp", "username", "action", "details"])
        
        # Ajouter le nouveau log
        new_log_df = pd.DataFrame([log_entry])
        logs_df = pd.concat([logs_df, new_log_df], ignore_index=True)
        
        # Limiter à 1000 entrées (garder les plus récentes)
        if len(logs_df) > 1000:
            logs_df = logs_df.tail(1000).reset_index(drop=True)
        
        # Sauvegarder dans SharePoint
        excel_buffer = BytesIO()
        logs_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        
        client.save_binary_in_sharepoint(excel_buffer.getvalue(), SHAREPOINT_LOGS_PATH)
        logger.info(f"LOG: {username} - {action} - {details}")
        
    except Exception as e:
        logger.error(f"Erreur sauvegarde log: {e}")

def get_logs(limit: int = 100) -> List[Dict]:
    """
    Récupère les logs d'activité depuis SharePoint
    
    Args:
        limit: Nombre maximum de logs à retourner
        
    Returns:
        Liste des logs sous forme de dictionnaires
    """
    try:
        client = SharePointClient()
        binary_content = client.read_binary_file(SHAREPOINT_LOGS_PATH)
        
        logs_df = pd.read_excel(BytesIO(binary_content))
        logs_df = logs_df.fillna("")
        
        if logs_df.empty:
            return []
        
        # Trier par timestamp décroissant et limiter
        logs_df = logs_df.sort_values('timestamp', ascending=False)
        limited_logs = logs_df.head(limit)
        
        # Convertir en liste de dictionnaires
        return limited_logs.to_dict('records')
        
    except Exception as e:
        logger.error(f"Erreur lecture logs: {e}")
        return []

def get_logs_stats() -> Dict:
    """
    Calcule les statistiques des logs
    
    Returns:
        Dictionnaire avec les statistiques
    """
    try:
        client = SharePointClient()
        binary_content = client.read_binary_file(SHAREPOINT_LOGS_PATH)
        
        logs_df = pd.read_excel(BytesIO(binary_content))
        logs_df = logs_df.fillna("")
        
        if logs_df.empty:
            return {"total": 0, "users": 0, "actions": 0}
        
        # Calculer les statistiques
        total_logs = len(logs_df)
        unique_users = logs_df['username'].nunique() if 'username' in logs_df.columns else 0
        unique_actions = logs_df['action'].nunique() if 'action' in logs_df.columns else 0
        
        # Dates de premier et dernier log
        first_log = None
        last_log = None
        if 'timestamp' in logs_df.columns and not logs_df.empty:
            sorted_logs = logs_df.sort_values('timestamp')
            first_log = sorted_logs.iloc[0]['timestamp']
            last_log = sorted_logs.iloc[-1]['timestamp']
        
        return {
            "total": total_logs,
            "users": unique_users,
            "actions": unique_actions,
            "first_log": first_log,
            "last_log": last_log
        }
        
    except Exception as e:
        logger.error(f"Erreur stats logs: {e}")
        return {"total": 0, "users": 0, "actions": 0}

def clear_all_logs() -> bool:
    """
    Efface tous les logs (admin seulement)
    
    Returns:
        True si succès, False sinon
    """
    try:
        client = SharePointClient()
        
        # Créer un DataFrame vide avec les bonnes colonnes
        empty_df = pd.DataFrame(columns=["timestamp", "username", "action", "details"])
        
        # Sauvegarder le fichier vide
        excel_buffer = BytesIO()
        empty_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        
        client.save_binary_in_sharepoint(excel_buffer.getvalue(), SHAREPOINT_LOGS_PATH)
        logger.info("Tous les logs ont été effacés")
        
        return True
        
    except Exception as e:
        logger.error(f"Erreur effacement logs: {e}")
        return False

# ============================================================================
# 2. GESTION DE L'HISTORIQUE DES ANALYSES
# ============================================================================

def save_analysis_history(username: str, ticket_filename: str, question: str, 
                         analysis_result: dict, ticket_info: dict):
    """
    Sauvegarde une analyse dans l'historique SharePoint
    
    Args:
        username: Nom d'utilisateur
        ticket_filename: Nom du fichier ticket
        question: Question posée
        analysis_result: Résultat de l'analyse
        ticket_info: Informations du ticket
    """
    try:
        client = SharePointClient()
        
        # Créer l'entrée d'analyse
        analysis_entry = {
            "timestamp": datetime.now().isoformat(),
            "username": username,
            "ticket_filename": ticket_filename,
            "question": question,
            "result": analysis_result.get("result", "UNKNOWN"),
            "expense_type": analysis_result.get("expense_type", "Unknown"),
            "justification": analysis_result.get("justification", "")[:1000],  # Limiter la taille
            "amount": ticket_info.get("amount", ""),
            "currency": ticket_info.get("currency", ""), 
            "vendor": ticket_info.get("vendor", ""),
            "confidence": ticket_info.get("confidence", 0)
        }
        
        # Lire le fichier existant ou créer un DataFrame vide
        try:
            binary_content = client.read_binary_file(SHAREPOINT_ANALYSIS_PATH)
            analysis_df = pd.read_excel(BytesIO(binary_content))
            analysis_df = analysis_df.fillna("")
            
            expected_columns = ["timestamp", "username", "ticket_filename", "question",
                              "result", "expense_type", "justification", "amount", 
                              "currency", "vendor", "confidence"]
            if list(analysis_df.columns) != expected_columns:
                logger.warning("Colonnes incorrectes dans le fichier analysis, recréation")
                analysis_df = pd.DataFrame(columns=expected_columns)
                
        except Exception as e:
            logger.info(f"Création nouveau fichier analysis: {e}")
            analysis_df = pd.DataFrame(columns=[
                "timestamp", "username", "ticket_filename", "question",
                "result", "expense_type", "justification", "amount",
                "currency", "vendor", "confidence"
            ])
        
        # Ajouter la nouvelle analyse
        new_analysis_df = pd.DataFrame([analysis_entry])
        analysis_df = pd.concat([analysis_df, new_analysis_df], ignore_index=True)
        
        # Limiter à 1000 analyses (garder les plus récentes)
        if len(analysis_df) > 1000:
            analysis_df = analysis_df.tail(1000).reset_index(drop=True)
        
        # Sauvegarder dans SharePoint
        excel_buffer = BytesIO()
        analysis_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        
        client.save_binary_in_sharepoint(excel_buffer.getvalue(), SHAREPOINT_ANALYSIS_PATH)
        logger.info(f"ANALYSIS: {username} - {ticket_filename}")
        
    except Exception as e:
        logger.error(f"Erreur sauvegarde analysis: {e}")

def get_analysis_history(limit: int = 100) -> List[Dict]:
    """
    Récupère l'historique des analyses depuis SharePoint
    
    Args:
        limit: Nombre maximum d'analyses à retourner
        
    Returns:
        Liste des analyses avec la structure attendue par le frontend
    """
    try:
        client = SharePointClient()
        binary_content = client.read_binary_file(SHAREPOINT_ANALYSIS_PATH)
        
        analysis_df = pd.read_excel(BytesIO(binary_content))
        analysis_df = analysis_df.fillna("")
        
        if analysis_df.empty:
            return []
        
        # Trier par timestamp décroissant et limiter
        analysis_df = analysis_df.sort_values('timestamp', ascending=False)
        limited_df = analysis_df.head(limit)
        
        # Reconstituer la structure pour le frontend
        analyses = []
        for _, row in limited_df.iterrows():
            analysis = {
                "timestamp": row.get("timestamp", ""),
                "user": row.get("username", ""),
                "ticket_filename": row.get("ticket_filename", ""),
                "question": row.get("question", ""),
                "analysis_result": {
                    "result": row.get("result", ""),
                    "expense_type": row.get("expense_type", ""),
                    "justification": row.get("justification", "")
                },
                "ticket_info": {
                    "amount": row.get("amount", ""),
                    "currency": row.get("currency", ""),
                    "vendor": row.get("vendor", ""),
                    "confidence": row.get("confidence", 0)
                }
            }
            analyses.append(analysis)
        
        return analyses
        
    except Exception as e:
        logger.error(f"Erreur lecture analysis history: {e}")
        return []

# ============================================================================
# 3. GESTION DES FEEDBACKS
# ============================================================================

def save_feedback(username: str, rating: int, comment: str = "", 
                 issue_type: str = "", analysis_id: str = ""):
    """
    Sauvegarde un feedback dans SharePoint
    
    Args:
        username: Nom d'utilisateur
        rating: Note de 1 à 5
        comment: Commentaire optionnel
        issue_type: Type de problème
        analysis_id: ID de l'analyse concernée
    """
    try:
        client = SharePointClient()
        
        # Créer l'entrée de feedback
        feedback_entry = {
            "timestamp": datetime.now().isoformat(),
            "username": username,
            "rating": rating,
            "comment": comment[:500] if comment else "",  # Limiter la taille
            "issue_type": issue_type,
            "analysis_id": analysis_id
        }
        
        # Lire le fichier existant ou créer un DataFrame vide
        try:
            binary_content = client.read_binary_file(SHAREPOINT_FEEDBACK_PATH)
            feedback_df = pd.read_excel(BytesIO(binary_content))
            feedback_df = feedback_df.fillna("")
            
            expected_columns = ["timestamp", "username", "rating", "comment", 
                              "issue_type", "analysis_id"]
            if list(feedback_df.columns) != expected_columns:
                logger.warning("Colonnes incorrectes dans le fichier feedback, recréation")
                feedback_df = pd.DataFrame(columns=expected_columns)
                
        except Exception as e:
            logger.info(f"Création nouveau fichier feedback: {e}")
            feedback_df = pd.DataFrame(columns=[
                "timestamp", "username", "rating", "comment", 
                "issue_type", "analysis_id"
            ])
        
        # Ajouter le nouveau feedback
        new_feedback_df = pd.DataFrame([feedback_entry])
        feedback_df = pd.concat([feedback_df, new_feedback_df], ignore_index=True)
        
        # Limiter à 500 feedbacks (garder les plus récents)
        if len(feedback_df) > 500:
            feedback_df = feedback_df.tail(500).reset_index(drop=True)
        
        # Sauvegarder dans SharePoint
        excel_buffer = BytesIO()
        feedback_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        
        client.save_binary_in_sharepoint(excel_buffer.getvalue(), SHAREPOINT_FEEDBACK_PATH)
        logger.info(f"FEEDBACK: {username} - Rating {rating}/5")
        
    except Exception as e:
        logger.error(f"Erreur sauvegarde feedback: {e}")

def get_feedback_stats() -> Dict:
    """
    Calcule les statistiques des feedbacks
    
    Returns:
        Dictionnaire avec les statistiques
    """
    try:
        client = SharePointClient()
        binary_content = client.read_binary_file(SHAREPOINT_FEEDBACK_PATH)
        
        feedback_df = pd.read_excel(BytesIO(binary_content))
        feedback_df = feedback_df.fillna("")
        
        if feedback_df.empty:
            return {
                "total_feedback": 0,
                "average_rating": 0.0,
                "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0},
                "common_issues": {}
            }
        
        # Calculer les statistiques
        total_feedback = len(feedback_df)
        
        # Ratings valides (entre 1 et 5)
        valid_ratings = [r for r in feedback_df.get('rating', []) 
                        if pd.notna(r) and 1 <= r <= 5]
        average_rating = round(sum(valid_ratings) / len(valid_ratings), 2) if valid_ratings else 0.0
        
        # Distribution des ratings
        rating_distribution = {i: 0 for i in range(1, 6)}
        for rating in valid_ratings:
            rating_distribution[int(rating)] += 1
        
        # Problèmes les plus fréquents
        issues = [str(i) for i in feedback_df.get('issue_type', []) 
                 if pd.notna(i) and str(i).strip()]
        from collections import Counter
        common_issues = dict(Counter(issues))
        
        return {
            "total_feedback": total_feedback,
            "average_rating": average_rating,
            "rating_distribution": rating_distribution,
            "common_issues": common_issues
        }
        
    except Exception as e:
        logger.error(f"Erreur stats feedback: {e}")
        return {
            "total_feedback": 0,
            "average_rating": 0.0,
            "rating_distribution": {1: 0, 2: 0, 3: 0, 4: 0, 5: 0},
            "common_issues": {}
        }

# ============================================================================
# FONCTION DE NETTOYAGE COMPLET (ADMIN SEULEMENT)
# ============================================================================

def clear_all_data() -> bool:
    """
    Efface toutes les données (logs, analyses, feedbacks) - ADMIN SEULEMENT
    
    Returns:
        True si succès, False sinon
    """
    try:
        success_count = 0
        
        # Effacer les logs
        if clear_all_logs():
            success_count += 1
        
        # Effacer l'historique des analyses
        try:
            client = SharePointClient()
            empty_df = pd.DataFrame(columns=[
                "timestamp", "username", "ticket_filename", "question",
                "result", "expense_type", "justification", "amount",
                "currency", "vendor", "confidence"
            ])
            excel_buffer = BytesIO()
            empty_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            client.save_binary_in_sharepoint(excel_buffer.getvalue(), SHAREPOINT_ANALYSIS_PATH)
            success_count += 1
        except Exception as e:
            logger.error(f"Erreur effacement analysis: {e}")
        
        # Effacer les feedbacks
        try:
            client = SharePointClient()
            empty_df = pd.DataFrame(columns=[
                "timestamp", "username", "rating", "comment",
                "issue_type", "analysis_id"
            ])
            excel_buffer = BytesIO()
            empty_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)
            client.save_binary_in_sharepoint(excel_buffer.getvalue(), SHAREPOINT_FEEDBACK_PATH)
            success_count += 1
        except Exception as e:
            logger.error(f"Erreur effacement feedback: {e}")
        
        logger.info(f"Nettoyage terminé: {success_count}/3 fichiers effacés")
        return success_count == 3
        
    except Exception as e:
        logger.error(f"Erreur nettoyage complet: {e}")
        return False

# ============================================================================
# FONCTIONS DE TEST (OPTIONNELLES)
# ============================================================================

def test_system():
    """Teste toutes les fonctions du système"""
    print("=== Test du système User Management ===")
    
    # Test authentification
    print("1. Test authentification...")
    user = authenticate_user("daniel.guez@natixis.com", "admin123")
    print(f"Authentification: {'OK' if user else 'ERREUR'}")
    
    # Test logs
    print("2. Test logs...")
    log_activity("test@example.com", "TEST", "Test du système")
    logs = get_logs(5)
    print(f"Logs: {len(logs)} entrées récupérées")
    
    # Test analyses
    print("3. Test analyses...")
    save_analysis_history(
        "test@example.com", "test_ticket.pdf", "Test question",
        {"result": "PASS", "expense_type": "Test"}, 
        {"amount": 50, "currency": "EUR"}
    )
    analyses = get_analysis_history(5)
    print(f"Analyses: {len(analyses)} entrées récupérées")
    
    # Test feedbacks
    print("4. Test feedbacks...")
    save_feedback("test@example.com", 5, "Test feedback", "no_issue")
    stats = get_feedback_stats()
    print(f"Feedbacks: {stats['total_feedback']} entrées")
    
    print("=== Test terminé ===")

if __name__ == "__main__":
    test_system()
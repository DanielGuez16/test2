#!/usr/bin/env python3
"""
Script de diagnostic et démarrage pour T&E Chatbot
=================================================

Ce script teste les dépendances, corrige les problèmes courants
et lance l'application avec un diagnostic complet.
"""

import sys
import os
import subprocess
import importlib
from pathlib import Path
import json

def print_section(title):
    """Affiche une section avec style"""
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")

def print_status(message, status="INFO"):
    """Affiche un message avec statut"""
    symbols = {
        "OK": "✅",
        "ERROR": "❌", 
        "WARNING": "⚠️",
        "INFO": "ℹ️"
    }
    print(f"{symbols.get(status, 'ℹ️')} {message}")

def check_python_version():
    """Vérifie la version Python"""
    print_section("VÉRIFICATION PYTHON")
    
    version = sys.version_info
    if version.major == 3 and version.minor >= 7:
        print_status(f"Python {version.major}.{version.minor}.{version.micro} - Compatible", "OK")
        return True
    else:
        print_status(f"Python {version.major}.{version.minor} - Version trop ancienne (requis: 3.7+)", "ERROR")
        return False

def install_package(package_name, import_name=None):
    """Installe un package si nécessaire"""
    if import_name is None:
        import_name = package_name
    
    try:
        importlib.import_module(import_name)
        print_status(f"{package_name} - Déjà installé", "OK")
        return True
    except ImportError:
        print_status(f"Installation de {package_name}...", "INFO")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print_status(f"{package_name} - Installé avec succès", "OK")
            return True
        except subprocess.CalledProcessError:
            print_status(f"Échec installation {package_name}", "ERROR")
            return False

def check_dependencies():
    """Vérifie et installe les dépendances"""
    print_section("VÉRIFICATION DES DÉPENDANCES")
    
    # Dépendances obligatoires
    required_deps = [
        ("fastapi", "fastapi"),
        ("uvicorn", "uvicorn"),
        ("pandas", "pandas"),
        ("python-multipart", "multipart"),
        ("jinja2", "jinja2")
    ]
    
    # Dépendances optionnelles (pour fonctionnalités avancées)
    optional_deps = [
        ("pillow", "PIL"),
        ("pytesseract", "pytesseract"),
        ("pypdf2", "PyPDF2"),
        ("python-docx", "docx")
    ]
    
    all_ok = True
    
    print("Dépendances obligatoires:")
    for package, import_name in required_deps:
        if not install_package(package, import_name):
            all_ok = False
    
    print("\nDépendances optionnelles:")
    optional_ok = 0
    for package, import_name in optional_deps:
        if install_package(package, import_name):
            optional_ok += 1
    
    print_status(f"Fonctionnalités avancées: {optional_ok}/{len(optional_deps)} disponibles", "INFO")
    
    return all_ok

def check_file_structure():
    """Vérifie la structure des fichiers"""
    print_section("VÉRIFICATION STRUCTURE FICHIERS")
    
    required_files = [
        "run.py",
        "llm_connector.py", 
        "te_document_processor.py",
        "user_management.py",
        "templates/te_index.html",
        "templates/login.html",
        "static/js/te_main.js"
    ]
    
    required_dirs = [
        "data",
        "templates", 
        "static",
        "static/js",
        "static/css",
        "uploads",
        "te_documents"
    ]
    
    all_ok = True
    
    # Vérifier fichiers
    for file_path in required_files:
        if Path(file_path).exists():
            print_status(f"Fichier {file_path} - Présent", "OK")
        else:
            print_status(f"Fichier {file_path} - MANQUANT", "ERROR")
            all_ok = False
    
    # Créer dossiers manquants
    for dir_path in required_dirs:
        if not Path(dir_path).exists():
            Path(dir_path).mkdir(parents=True, exist_ok=True)
            print_status(f"Dossier {dir_path} - Créé", "INFO")
        else:
            print_status(f"Dossier {dir_path} - Présent", "OK")
    
    return all_ok

def fix_encoding_issues():
    """Corrige les problèmes d'encodage dans les fichiers"""
    print_section("CORRECTION ENCODAGE")
    
    files_to_fix = [
        "llm_connector.py",
        "run.py", 
        "te_document_processor.py",
        "user_management.py"
    ]
    
    encoding_fixes = {
        'Ã©': 'é',
        'Ã ': 'à', 
        'Ã¨': 'è',
        'Ã´': 'ô',
        'Ã§': 'ç',
        'Ã‰': 'É'
    }
    
    for file_path in files_to_fix:
        if Path(file_path).exists():
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                fixed_content = content
                fixes_made = 0
                
                for wrong, correct in encoding_fixes.items():
                    if wrong in fixed_content:
                        fixed_content = fixed_content.replace(wrong, correct)
                        fixes_made += 1
                
                if fixes_made > 0:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(fixed_content)
                    print_status(f"{file_path} - {fixes_made} corrections d'encodage", "OK")
                else:
                    print_status(f"{file_path} - Aucune correction nécessaire", "OK")
                    
            except Exception as e:
                print_status(f"Erreur lors de la correction de {file_path}: {e}", "WARNING")

def fix_code_issues():
    """Corrige les problèmes de code identifiés"""
    print_section("CORRECTION PROBLÈMES CODE")
    
    # Correction dans run.py
    run_py_path = Path("run.py")
    if run_py_path.exists():
        try:
            with open(run_py_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Correction du calcul des règles Excel
            old_line = 'logger.info(f"Documents T&E chargés: {len(excel_rules)} règles Excel'
            new_line = 'total_rules = sum(len(rules) for rules in excel_rules.values())\n        logger.info(f"Documents T&E chargés: {total_rules} règles Excel'
            
            if old_line in content and 'total_rules = sum' not in content:
                content = content.replace(old_line, new_line)
                
                with open(run_py_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                print_status("run.py - Correction calcul règles Excel", "OK")
            else:
                print_status("run.py - Aucune correction nécessaire", "OK")
                
        except Exception as e:
            print_status(f"Erreur correction run.py: {e}", "WARNING")

def create_requirements_txt():
    """Crée le fichier requirements.txt"""
    print_section("CRÉATION REQUIREMENTS.TXT")
    
    requirements = [
        "fastapi>=0.68.0",
        "uvicorn[standard]>=0.15.0", 
        "pandas>=1.3.0",
        "python-multipart>=0.0.5",
        "jinja2>=3.0.0",
        "pillow>=8.0.0",
        "pytesseract>=0.3.8", 
        "pypdf2>=1.26.0",
        "python-docx>=0.8.11"
    ]
    
    with open("requirements.txt", "w") as f:
        f.write("\n".join(requirements))
    
    print_status("requirements.txt créé", "OK")

def test_basic_imports():
    """Teste les imports de base"""
    print_section("TEST IMPORTS MODULES")
    
    modules_to_test = [
        ("llm_connector", "LLMConnector"),
        ("te_document_processor", "TEDocumentProcessor"), 
        ("user_management", "authenticate_user")
    ]
    
    all_ok = True
    
    for module_name, item_name in modules_to_test:
        try:
            module = importlib.import_module(module_name)
            if hasattr(module, item_name):
                print_status(f"{module_name}.{item_name} - Import OK", "OK")
            else:
                print_status(f"{module_name}.{item_name} - Attribut manquant", "ERROR")
                all_ok = False
        except Exception as e:
            print_status(f"{module_name} - Erreur import: {e}", "ERROR")
            all_ok = False
    
    return all_ok

def test_user_auth():
    """Teste le système d'authentification"""
    print_section("TEST AUTHENTIFICATION")
    
    try:
        from user_management import authenticate_user, USERS_DB
        
        # Test utilisateur valide
        user = authenticate_user("admin.te@company.com", "admin123")
        if user and user.get("role") == "admin":
            print_status("Authentification admin - OK", "OK")
        else:
            print_status("Authentification admin - ÉCHEC", "ERROR")
            return False
        
        # Test utilisateur invalide
        invalid_user = authenticate_user("fake@test.com", "wrong")
        if invalid_user is None:
            print_status("Rejet utilisateur invalide - OK", "OK")
        else:
            print_status("Rejet utilisateur invalide - ÉCHEC", "ERROR")
            return False
        
        print_status(f"Utilisateurs configurés: {len(USERS_DB)}", "INFO")
        return True
        
    except Exception as e:
        print_status(f"Erreur test authentification: {e}", "ERROR")
        return False

def test_llm_connector():
    """Teste le connecteur LLM"""
    print_section("TEST CONNECTEUR LLM")
    
    try:
        from llm_connector import LLMConnector
        
        connector = LLMConnector()
        response = connector.get_llm_response("Test question about T&E policy")
        
        if response and len(response) > 10:
            print_status("Connecteur LLM - OK", "OK")
            print_status(f"Réponse test: {response[:100]}...", "INFO")
            return True
        else:
            print_status("Connecteur LLM - Réponse vide", "ERROR")
            return False
            
    except Exception as e:
        print_status(f"Erreur test LLM: {e}", "ERROR")
        return False

def create_test_files():
    """Crée des fichiers de test si nécessaires"""
    print_section("CRÉATION FICHIERS TEST")
    
    # Créer un fichier de données T&E simulé
    test_data_dir = Path("data")
    test_data_dir.mkdir(exist_ok=True)
    
    # Fichier de configuration test
    config_test = {
        "app_name": "T&E Chatbot APAC",
        "version": "1.0.0",
        "test_mode": True,
        "created_by": "startup_diagnostic.py"
    }
    
    with open("data/config_test.json", "w") as f:
        json.dump(config_test, f, indent=2)
    
    print_status("Fichiers de test créés", "OK")

def run_application():
    """Lance l'application"""
    print_section("LANCEMENT APPLICATION")
    
    try:
        print_status("Démarrage du serveur FastAPI...", "INFO")
        print_status("Interface: http://localhost:8000", "INFO")
        print_status("Login admin: admin.te@company.com / admin123", "INFO")
        print_status("Login user: user.singapore@company.com / user123", "INFO")
        print("")
        print_status("Appuyez sur Ctrl+C pour arrêter", "INFO")
        
        # Import et lancement
        from run import app
        import uvicorn
        
        uvicorn.run(
            app,
            host="0.0.0.0", 
            port=8000,
            reload=False,
            log_level="info"
        )
        
    except KeyboardInterrupt:
        print_status("\nApplication arrêtée par l'utilisateur", "INFO")
    except Exception as e:
        print_status(f"Erreur lancement: {e}", "ERROR")
        return False
    
    return True

def main():
    """Fonction principale de diagnostic"""
    print("🚀 T&E Chatbot - Diagnostic et Démarrage")
    print("=" * 60)
    
    # Étapes de diagnostic
    steps = [
        ("Version Python", check_python_version),
        ("Dépendances", check_dependencies), 
        ("Structure fichiers", check_file_structure),
        ("Encodage", fix_encoding_issues),
        ("Problèmes code", fix_code_issues),
        ("Requirements.txt", create_requirements_txt),
        ("Imports modules", test_basic_imports),
        ("Authentification", test_user_auth),
        ("Connecteur LLM", test_llm_connector),
        ("Fichiers test", create_test_files)
    ]
    
    all_passed = True
    
    for step_name, step_func in steps:
        try:
            result = step_func()
            if result is False:
                all_passed = False
        except Exception as e:
            print_status(f"Erreur lors de {step_name}: {e}", "ERROR")
            all_passed = False
    
    # Résumé
    print_section("RÉSUMÉ DIAGNOSTIC")
    
    if all_passed:
        print_status("Tous les tests ont réussi!", "OK")
        print_status("L'application devrait fonctionner correctement", "OK")
        
        # Proposer le lancement
        response = input("\n🚀 Lancer l'application maintenant? (y/N): ")
        if response.lower() in ['y', 'yes', 'oui']:
            run_application()
        else:
            print_status("Pour lancer manuellement: python run.py", "INFO")
    else:
        print_status("Certains tests ont échoué", "WARNING")
        print_status("L'application peut fonctionner partiellement", "WARNING")
        print_status("Corrigez les erreurs puis relancez ce script", "INFO")

if __name__ == "__main__":
    main()
def read_docx_file_as_text(self, binary_content: bytes) -> str:
    """
    Lit un fichier DOCX depuis du contenu binaire et retourne le texte
    
    Args:
        binary_content: Contenu binaire du fichier DOCX
        
    Returns:
        str: Texte extrait du document
    """
    try:
        from docx import Document
        import io
        
        # Créer un objet Document depuis les bytes
        doc = Document(io.BytesIO(binary_content))
        
        # Extraire tout le texte des paragraphes
        paragraphs = []
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():  # Ignorer les paragraphes vides
                paragraphs.append(paragraph.text.strip())
        
        # Joindre tous les paragraphes
        full_text = '\n\n'.join(paragraphs)
        
        return full_text
        
    except Exception as e:
        raise Exception(f"Erreur lecture fichier DOCX: {str(e)}")
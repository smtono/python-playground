import win32com.client

def update_fields(doc_path):
    """
    Open a Word document, update all fields, and save the document.
    
    Parameters:
    - doc_path (str): The path to the Word document.
    """
    # Initialize the Word application
    word_app = win32com.client.Dispatch("Word.Application")

    # Update fields
    doc = word_app.Documents.Open(doc_path)
    doc.Fields.Update()

    # Save and close the document
    doc.Save()
    doc.Close()
    word_app.Quit()

# TOC
# Load the document
doc_path = r'C:\Users\smtho\workspaces\vscode\python-playground\src\word\populated_document.docx'
doc = Document(doc_path)

# Insert a paragraph for the TOC at the beginning of the document
toc_paragraph = doc.paragraphs[0].insert_paragraph_before()
insert_toc(toc_paragraph)

# Save the document
new_doc_path = 'document_with_toc.docx'
doc.save(new_doc_path)

# TABLE
csv_file = r'C:\Users\smtho\workspaces\vscode\python-playground\src\word\username.csv'
document = Document(r'C:\Users\smtho\workspaces\vscode\python-playground\src\word\table.docx')

new_document = Document()
table = document.tables[0]
copy_table(table, new_document, csv_file)

new_document.save(r'C:\Users\smtho\workspaces\vscode\python-playground\src\word\new.docx')
pathlib.Path.unlink(pathlib.Path(r'C:\Users\smtho\workspaces\vscode\python-playground\src\word\new.docx'))
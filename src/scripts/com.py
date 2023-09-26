"""com study"""

import os
import win32com.client as win32

# initiate word application
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
test = word.Documents
doc = word.Documents.Open(os.path.join(os.getcwd(), "src", "data", "test.docx"))
doc = doc.Document

win32.Dispatch

# manipulate word doc


# end of life
word.ActiveDocument.Save()
word.Quit()

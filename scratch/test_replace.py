import win32com.client
import os

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(os.path.abspath("test_template.docx"), ReadOnly=False)
    
    find_zmist = doc.Content.Find
    find_zmist.Text = "{{зміст}}"
    if find_zmist.Execute():
        zmist_rng = find_zmist.Parent
        zmist_para_rng = zmist_rng.Paragraphs(1).Range
        
        # Suppose we want to insert two paragraphs.
        # We can collapse the paragraph range, but wait!
        # If we just do zmist_para_rng.Text = "", it deletes the paragraph (including \r).
        # But wait, if we delete the \r, the text below might merge.
        # Let's see:
        zmist_para_rng.Text = "Para 1\rPara 2\r"
        
    doc.SaveAs2(os.path.abspath("test_out.docx"), 16)
finally:
    doc.Close(False)
    word.Quit()

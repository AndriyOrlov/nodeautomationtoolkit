import win32com.client
import os

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
try:
    doc1 = word.Documents.Add()
    doc1.Range().Text = "Src Line 1\rSrc Line 2\rSrc Line 3\r"
    
    doc2 = word.Documents.Open(os.path.abspath("test_template.docx"), ReadOnly=False)
    
    find_zmist = doc2.Content.Find
    find_zmist.Text = "{{зміст}}"
    if find_zmist.Execute():
        zmist_rng = find_zmist.Parent
        zmist_para = zmist_rng.Paragraphs(1).Range
        
        # Select the paragraph to delete it and remember the position?
        insert_point = zmist_para.Start
        zmist_para.Delete()
        
        # Now insert_point is at the location where zmist_para used to be
        dest_rng = doc2.Range(insert_point, insert_point)
        dest_rng.FormattedText = doc1.Paragraphs(2).Range.FormattedText
        
        # Advance insert point
        insert_point = dest_rng.End
        dest_rng2 = doc2.Range(insert_point, insert_point)
        dest_rng2.FormattedText = doc1.Paragraphs(3).Range.FormattedText
        
    doc2.SaveAs2(os.path.abspath("test_out2.docx"), 16)
finally:
    doc1.Close(False)
    doc2.Close(False)
    word.Quit()

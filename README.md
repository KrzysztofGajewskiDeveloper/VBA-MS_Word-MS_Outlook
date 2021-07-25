# MassMailings (word content to email body)

Important references: Microsoft Outlook 16.0 Object Library, Microsoft Word 16.0 Object Library

The macro has been developed to provide mass mailing with standarized layouts to the receipients. 

1) Turning off screen updating and visibility of the application can make the code run faster

        .ScreenUpdating = False
        .Visible = False
    
2) Create important objects - Outlook, Word, document(word). Set the SaveInterval property to zero to turn off saving AutoRecover information.
    
    Set OlApp = New Outlook.Application
    Set wd = New Word.Application
    Set doc = wd.Documents.Open(Path)
    wd.Options.SaveInterval = 0
    
3) Loop 
    
    LastRow = wsInput.Cells(Rows.Count, 1).End(xlUp).Row
 
        For i = 2 To LastRow
              
            Recipient = wsInput.Cells(i, "A")
            Input1 = wsInput.Cells(i, "B")
            Input2 = wsInput.Cells(i, "C")
                    
                    With wd.Selection
                        .Find.Execute findText:=PlaceHolder
                        .InsertAfter Input1
                        .Find.Text = PlaceHolder
                        .Find.Replacement.Text = vbNullString
                        .Find.Execute Replace:=wdReplaceAll
                    End With
                    
                    With wd.Selection
                        .Find.Execute findText:=PlaceHolder2
                        .InsertAfter Input2
                        .Find.Text = PlaceHolder2
                        .Find.Replacement.Text = vbNullString
                        .Find.Execute Replace:=wdReplaceAll
                    End With
                        
                    Set OlEmail = OlApp.CreateItem(olMailItem)
                    With OlEmail
                        .BodyFormat = olFormatRichText
                        .Display
                        .To = Recipient
                        .Subject = Subject
                        Set Editor = .GetInspector.WordEditor
                        doc.Content.Copy
                        Editor.Content.Paste
                        .send
                    End With
            
                For Undo = 1 To 4
                    doc.Undo
                Next Undo
        
        Next i
        
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Visible = True
    End With
    
    wd.DisplayAlerts = False
    doc.Close SaveChanges:=False
    wd.Quit SaveChanges:=False
    
End Sub
© 2021 GitHub, Inc.
Terms
Privacy
Security
Status
Docs
Contact GitHub
Pricing
API
Training
Blog
About

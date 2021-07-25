# MassMailings (word content to email body)

Important references: Microsoft Outlook 16.0 Object Library, Microsoft Word 16.0 Object Library

The macro developed to provide mass mailing with standarized layouts to the recipients. 

1) Turning off screen updating and visibility of the application can make the code run faster

        .ScreenUpdating = False
        .DisplayAlerts = False
        
2) Set the SaveInterval property to zero to turn off saving AutoRecover information.
    
        .Options.SaveInterval = 0
    
3) Column A - email addressess
   Column B - first placeholder which will vary for each recipient
   Column C - second placeholder which will vary for each recipient
  
        For i = 2 To LastRow
              
            Recipient = wsInput.Cells(i, "A")
            Input1 = wsInput.Cells(i, "B")
            Input2 = wsInput.Cells(i, "C")


4) Replacing the placeholders with the input from above-mentioned columns (the same action for both placeholders)
                    
                        .Find.Execute findText:=PlaceHolder
                        .InsertAfter Input1
                        .Find.Text = PlaceHolder
                        .Find.Replacement.Text = vbNullString
                        .Find.Execute Replace:=wdReplaceAll
                        
5) Pasting the output to the body of the email and sending to each recipient with a specific output based on the standarized template

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
                    
6) Undo method to allow working on 1 word without closing and reopening the document

                For Undo = 1 To 4
                    doc.Undo
                Next Undo
        
        

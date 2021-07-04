Attribute VB_Name = "MassMailings"
Option Explicit

Sub MassMailings()

Const Path As String = "C:\Users\futur\Desktop\Input.docx"
Const Subject As String = "Expected email subject"
Const PlaceHolder As String = "<<Input1>>"
Const PlaceHolder2 As String = "<<Input2>>"

Dim OlApp As Outlook.Application, OlEmail As Outlook.MailItem, OlInsp As Outlook.Inspector, wd As Word.Application, doc As Word.document, Editor As Object
Dim LastRow As Long, i As Long, Recipient As String, Input1 As String, Input2 As String, Undo As Integer

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Visible = False
    End With
 
    Set OlApp = New Outlook.Application
    Set wd = New Word.Application
    Set doc = wd.Documents.Open(Path)
    wd.Options.SaveInterval = 0
    
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

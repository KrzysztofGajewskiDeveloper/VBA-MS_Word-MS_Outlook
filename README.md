# MassMailings (word content to email body)

Important references: Microsoft Outlook 16.0 Object Library, Microsoft Word 16.0 Object Library

The macro has been developed to provide mass mailing with standarized layouts to the receipients. 

1) Create word object and set the document in order to copy the word document content to the clipboard

    Set wd = New Word.Application
    Set doc = wd.Documents.Open(Path)


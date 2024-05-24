Attribute VB_Name = "DumpAttachments"
Sub SaveAttachmentsFromEmails()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olInbox As Object ' Store the Inbox folder
    Dim olFilteredItems As Object
    Dim olMail As Object
    Dim olAttachment As Object
    
    Dim startDate As Date
    Dim endDate As Date
    Dim senderEmail As String
    Dim saveFolderPath As String
    
    Dim i As Long
    Dim folderName As String
    Dim FSO As Object
    Dim fileNumber As Integer
    
    On Error Resume Next
    
    ' Get data from named cells in the sheet named "Dump Attachments"
    If Not IsDate(ThisWorkbook.Sheets("Dump Attachments").Range("StartDate").Value) Then
        MsgBox "Invalid start date.", vbExclamation
        Exit Sub
    End If
    startDate = ThisWorkbook.Sheets("Dump Attachments").Range("StartDate").Value
    
    If Not IsDate(ThisWorkbook.Sheets("Dump Attachments").Range("EndDate").Value) Then
        MsgBox "Invalid end date.", vbExclamation
        Exit Sub
    End If
    endDate = ThisWorkbook.Sheets("Dump Attachments").Range("EndDate").Value
    
    senderEmail = Trim(ThisWorkbook.Sheets("Dump Attachments").Range("SenderEmail").Value)
    If senderEmail = "" Or InStr(senderEmail, "@") = 0 Then
        MsgBox "Invalid sender email address.", vbExclamation
        Exit Sub
    End If
    
    saveFolderPath = Trim(ThisWorkbook.Sheets("Dump Attachments").Range("FolderPath").Value)
    If saveFolderPath = "" Then
        MsgBox "Save folder path is empty.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the save folder path exists, if not, create it
    If Dir(saveFolderPath, vbDirectory) = "" Then
        MsgBox "Save folder path does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Create a folder based on start and end dates
    folderName = Format(startDate, "yyyy-mm-dd") & " to " & Format(endDate, "yyyy-mm-dd")
    saveFolderPath = saveFolderPath & "\" & folderName
    
    ' Check if the folder already exists, if not, create it
    If Dir(saveFolderPath, vbDirectory) = "" Then
        MkDir saveFolderPath
    End If
    
    ' Initialize Outlook application
    Set olApp = CreateObject("Outlook.Application")
    If olApp Is Nothing Then
        MsgBox "Outlook is not running or not installed.", vbExclamation
        Exit Sub
    End If
    
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Attempt to get the Inbox folder
    On Error Resume Next
    Set olInbox = olNamespace.GetDefaultFolder(6) ' 6 = olFolderInbox
    On Error GoTo 0
    
    If olInbox Is Nothing Then
        MsgBox "Inbox folder not found.", vbExclamation
        Set olNamespace = Nothing
        Set olApp = Nothing
        Exit Sub
    End If
    
    ' Filter items based on the date range and sender's email address
    Set olFilteredItems = olInbox.Items.Restrict("[ReceivedTime] >= '" & Format(startDate, "ddddd h:nn AMPM") & "' And [ReceivedTime] <= '" & Format(endDate, "ddddd h:nn AMPM") & "' And [SenderEmailAddress] = '" & senderEmail & "'")
    
    ' Loop through filtered items
    For i = 1 To olFilteredItems.Count
        Set olMail = olFilteredItems.Item(i)
        
        ' Save each attachment
        For Each olAttachment In olMail.Attachments
            ' Save attachments to the folder with start and end date as its name
            Dim fileName As String
            fileName = olAttachment.fileName
            fileNumber = 1
            
            Do While Dir(saveFolderPath & "\" & fileName) <> ""
                fileName = Left(olAttachment.fileName, InStrRev(olAttachment.fileName, ".") - 1) & " (" & fileNumber & ")" & Right(olAttachment.fileName, Len(olAttachment.fileName) - InStrRev(olAttachment.fileName, ".") + 1)
                fileNumber = fileNumber + 1
            Loop
            
            olAttachment.SaveAsFile saveFolderPath & "\" & fileName
        Next olAttachment
    Next i
    
    ' Cleanup
    Set olAttachment = Nothing
    Set olMail = Nothing
    Set olFilteredItems = Nothing
    Set olInbox = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    
    MsgBox "Attachments have been saved to " & saveFolderPath, vbInformation
End Sub


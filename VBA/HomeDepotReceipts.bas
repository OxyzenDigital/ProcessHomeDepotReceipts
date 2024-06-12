Attribute VB_Name = "HomeDepotReceipts"
Option Explicit

Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As LongPtr, ByVal bInheritHandle As Long, ByVal dwProcessId As LongPtr) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Const INFINITE As Long = &HFFFF&

Sub ProcessTEXTfromPDFs()
    Dim pythonScriptPath As String
    Dim folderPath As String
    Dim pythonProcess As Double
    Dim processHandle As LongPtr
    
    ' Clear contents of the first sheet
    ThisWorkbook.Sheets(1).Cells.Clear
    'ClearAllButFirstRow
    
    ' Prompt user to select folder where text files are located
    folderPath = BrowseForFolder("") ' Start browsing from C: drive
    
    ' If user cancels folder selection, exit sub
    If folderPath = "" Then Exit Sub
    
    ' Call CSV extraction process and fill the cells.
    ExtractDataFromPDF folderPath
End Sub

Function BrowseForFolder(initialFolder As String) As String
    Dim folderDialog As FileDialog
    Dim selectedFolder As Variant
    
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.title = "Select Folder"
    folderDialog.InitialFileName = initialFolder ' Use provided initial folder
    folderDialog.AllowMultiSelect = False ' Allow only single folder selection
    
    ' Show folder picker dialog
    If folderDialog.Show = -1 Then ' If user selects a folder
        selectedFolder = folderDialog.SelectedItems(1)
        BrowseForFolder = selectedFolder
    Else
        BrowseForFolder = "" ' Return empty string if user cancels
    End If
End Function

Function GetOneDrivePaths() As Object
    Dim objShell As Object
    Dim objEnv As Object
    Dim oneDrivePaths As Object
    Dim oneDrivePath As String
    Dim oneDriveBusinessPath As String
    Dim oneDriveConsumerPath As String

    Set objShell = CreateObject("WScript.Shell")
    Set objEnv = objShell.Environment("PROCESS")
    
    Set oneDrivePaths = CreateObject("Scripting.Dictionary")
    oneDrivePath = objEnv("OneDrive")
    oneDriveBusinessPath = objEnv("OneDriveCommercial")
    oneDriveConsumerPath = objEnv("OneDriveConsumer")
    
    If oneDrivePath <> "" Then
        oneDrivePaths.Add "Personal", oneDrivePath
    End If
    If oneDriveBusinessPath <> "" Then
        oneDrivePaths.Add "Business", oneDriveBusinessPath
    End If
    If oneDriveConsumerPath <> "" Then
        oneDrivePaths.Add "Consumer", oneDriveConsumerPath
    End If
    
    Set GetOneDrivePaths = oneDrivePaths
End Function

Sub ExtractDataFromPDF(folderPath As String)
    On Error GoTo ErrorHandler
    
    Dim pythonScriptPath As String
    Dim pythonProcess As Double
    Dim oneDrivePaths As Object
    Dim oneDrivePath As Variant
    Dim fileFound As Boolean
    
    ' Get OneDrive paths
    Set oneDrivePaths = GetOneDrivePaths()
    
    ' Exit sub if no OneDrive paths are found
    If oneDrivePaths.Count = 0 Then
        MsgBox "OneDrive path not found.", vbExclamation
        Exit Sub
    End If
    
    fileFound = False
    
    ' Check each OneDrive path for the file
    For Each oneDrivePath In oneDrivePaths.Items
        pythonScriptPath = oneDrivePath & "\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\ProcessHomeDepotReceipts\HomeDepot Receipt Processing.py"
        
        If Dir(pythonScriptPath) <> "" Then
            fileFound = True
            Exit For
        End If
    Next oneDrivePath
    
    If Not fileFound Then
        MsgBox "Python script file not found in any OneDrive location: " & pythonScriptPath, vbExclamation
        Exit Sub
    End If
    
    ' Run Python script with folder path as a parameter
    pythonProcess = Shell("python """ & pythonScriptPath & """ """ & folderPath & """", vbNormalFocus)
    
    ' Wait for Python process to finish
    Do While ProcessRunning(pythonProcess)
        DoEvents
    Loop
    
    ProcessCSVFile folderPath

    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Sub OpenCSVFile()
    Dim folderPath As String
    ' Clear contents of the first sheet
    ThisWorkbook.Sheets(1).Cells.Clear
    'ClearAllButFirstRow
    
    ' Prompt user to select folder where text files are located
    folderPath = BrowseForFolder("C:\") ' Start browsing from C: drive
    
    ' If user cancels folder selection, exit sub
    If folderPath = "" Then Exit Sub
    ProcessCSVFile (folderPath)
End Sub

Sub ProcessCSVFile(folderPath As String)
    On Error GoTo ErrorHandler
    
    Dim csvWB As Workbook
    Dim csvFilePath As String
    
    ' CSV file path
    csvFilePath = folderPath & "\output.csv"
    
    ' Check if the CSV file exists
    If Dir(csvFilePath) <> "" Then
        ' Open the CSV file with UTF-8 encoding
        Set csvWB = Workbooks.Open(filename:=csvFilePath, Origin:=xlWindows, Delimiter:=",", Format:=6, Local:=True)
        
        ' Copy data from CSV workbook to current workbook
        csvWB.Sheets(1).UsedRange.Copy ThisWorkbook.Sheets(1).Range("A3")
        
        ' Close CSV workbook without saving changes
        csvWB.Close False
        
        MsgBox "Data extraction completed successfully!"
    Else
        MsgBox "CSV file not found in the folder. Please check the file path."
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Function ClearAllButFirstRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Set the worksheet you want to clear (the first sheet in this case)
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Find the last used row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Find the last used column in the first row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Clear the contents of all cells except the first row
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
End Function

Function ProcessRunning(processID As Double) As Boolean
    On Error Resume Next
    ProcessRunning = Not (GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where ProcessID = " & processID).Count = 0)
    On Error GoTo 0
End Function


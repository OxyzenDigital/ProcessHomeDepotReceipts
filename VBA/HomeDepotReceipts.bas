Attribute VB_Name = "HomeDepotReceipts"
Option Explicit

Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As LongPtr, ByVal bInheritHandle As Long, ByVal dwProcessId As LongPtr) As LongPtr
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Const INFINITE = &HFFFF&

Sub ExtractTEXTfromPDFs()
    Dim pythonScriptPath As String
    Dim folderPath As String
    Dim pythonProcess As Double
    Dim processHandle As LongPtr
    
    ' Clear contents of the first sheet
    ThisWorkbook.Sheets(1).Cells.Clear
    
    ' Prompt user to select folder where text files are located
    folderPath = BrowseForFolder("C:\") ' Start browsing from C: drive
    
    ' If user cancels folder selection, exit sub
    If folderPath = "" Then Exit Sub
    
    ' Path to Python script (assuming it's in the same folder as the Excel file)
    pythonScriptPath = "O:\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\ProcessHomeDepotReceipts\ExtractTextFromPDFs.py"
    
    ' Run Python script with folder path as a parameter
    pythonProcess = Shell("python """ & pythonScriptPath & """ """ & folderPath & """", vbHide)
    
    ' Get process handle
    processHandle = OpenProcess(0, 0, pythonProcess)
    
    ' Wait for Python script to finish
    WaitForSingleObject processHandle, INFINITE
    
    ' Close process handle
    CloseHandle processHandle
    
    ' Call CVS extraction process and fill the cells.
    ExtractDataFromPythonScript (folderPath)
End Sub


Function BrowseForFolder(initialFolder As String) As String
    Dim folderDialog As FileDialog
    Dim selectedFolder As Variant
    
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.title = "Select Folder"
    folderDialog.InitialFileName = "O:\OneDrive - Oxyzen Digital\ODI PRACTICE\_Receipts Dump" ' Set initial folder to where text files are expected
    folderDialog.AllowMultiSelect = False ' Allow only single folder selection
    
    ' Show folder picker dialog
    If folderDialog.Show = -1 Then ' If user selects a folder
        selectedFolder = folderDialog.SelectedItems(1)
        If InStr(selectedFolder, "file:///") = 1 Then
            ' If it's a URL, convert it to drive letter path
            BrowseForFolder = Replace(selectedFolder, "file:///", "")
        Else
            ' Otherwise, it's already a drive letter path
            BrowseForFolder = selectedFolder
        End If
    Else
        BrowseForFolder = "" ' Return empty string if user cancels
    End If
End Function




' Process the TXT file for CSV importing.

Sub ExtractDataFromPythonScript(folderPath As String)
    On Error GoTo ErrorHandler
    
    Dim pythonScriptPath As String
    Dim pythonProcess As Double
    Dim csvWB As Workbook
    
    ' Path to Python script
    pythonScriptPath = "O:\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\ProcessHomeDepotReceipts\HomeDepot Receipt Processing.py"
    
    ' Check if Python script file exists
    If Dir(pythonScriptPath) = "" Then
        MsgBox "Python script file not found: " & pythonScriptPath
        Exit Sub
    End If
    
    ' Run Python script with folder path as a parameter
    pythonProcess = Shell("python """ & pythonScriptPath & """ """ & folderPath & """", vbNormalFocus)
    
    ' Wait for Python process to finish
    Do While ProcessRunning(pythonProcess)
        DoEvents
    Loop
    
    ' CSV file path
    Dim csvFilePath As String
    csvFilePath = "O:\OneDrive\03_PROFESSIONAL\OXYZEN Digital\Digital\GitHub\ProcessHomeDepotReceipts\output.csv"
    
    ' Check if the CSV file exists
    If Dir(csvFilePath) <> "" Then
        ' Open the CSV file with UTF-8 encoding
        Set csvWB = Workbooks.Open(fileName:=csvFilePath, Origin:=xlWindows, Delimiter:=",", Format:=6, Local:=True)
        
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


Function ProcessRunning(processID As Double) As Boolean
    On Error Resume Next
    ' Attempt to find the process ID, if found the process is still running
    ProcessRunning = Not (GetObject("winmgmts:").ExecQuery("Select * from Win32_Process Where ProcessID = " & processID).Count = 0)
    On Error GoTo 0
End Function


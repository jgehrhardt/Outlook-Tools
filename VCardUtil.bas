Attribute VB_Name = "VCardUtil"
'---------------------------------------------------------------------------------------
' Module    : VCardUtil
' Author    : John G Ehrhardt
' Date      : 12-Sep-2018
' Purpose   : Hack to Update VCards to V3 for Apple Processing
'---------------------------------------------------------------------------------------
Option Explicit
Public Sub ConvertVCard()
'---------------------------------------------------------------------------------------
' Procedure : ConvertVCard
' Author    : John G Ehrhardt
' Date      : 12-Sep-2018
' Purpose   : Change A VCard to V3, to import to Apple iCloud
'---------------------------------------------------------------------------------------
Dim DefPath As String
DefPath = Environ("USERPROFILE") & "\Desktop"

Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim vcfFiles As Collection

'Check to process files in Default Path (desktop)
Set vcfFiles = fnGetFiles(DefPath, FSO)

'Check is files to process, if not get another folder before giving up
If vcfFiles Is Nothing Then
    'Get a New Folder to test
    DefPath = fnSelectFolder
    If DefPath <> "" Then
        'Try Again to get files
        Set vcfFiles = fnGetFiles(DefPath, FSO)
    End If
    'check to see if still nothing
    If vcfFiles Is Nothing Then
        MsgBox "Found no files to process.  Exiting.", vbInformation + vbOKOnly, "No files found"
        Set FSO = Nothing
        Set vcfFiles = Nothing
        Exit Sub
    End If
End If

'Process the files found
Dim F As Scripting.File
Dim ans As VbMsgBoxResult
ans = MsgBox("Found " & vcfFiles.Count & " files to process, continue?", vbOKCancel, "Request to Process found files")
If ans = vbCancel Then Exit Sub
For Each F In vcfFiles
    If fnProcessFile(F, FSO) = True Then
        MsgBox "Error on File: " & F.Path, vbCritical, "Error!"
        Exit Sub
    End If
Next
Set FSO = Nothing
Set vcfFiles = Nothing
MsgBox "All Files have been processed"
End Sub
Private Function fnSelectFolder() As String
'---------------------------------------------------------------------------------------
' Procedure : fnSelectFolder
' Author    : John G Ehrhardt
' Date      : 12-Sep-2018
' Purpose   : Select a Folder or Not via EXCEL, return it as a string
'---------------------------------------------------------------------------------------
Dim xlApp As Object
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Dim fd As Office.FileDialog
Set fd = xlApp.Application.FileDialog(msoFileDialogFolderPicker)
If fd.Show = -1 Then fnSelectFolder = fd.SelectedItems(1)
Set fd = Nothing
xlApp.Quit
Set xlApp = Nothing
End Function
Private Function fnGetFiles(ByRef DefPath As String, ByRef FSO As FileSystemObject) As Collection
'---------------------------------------------------------------------------------------
' Procedure : fnGetFiles
' Author    : John G Ehrhardt
' Date      : 12-Sep-2018
' Purpose   : Gather a possible collection of files to process
'---------------------------------------------------------------------------------------
'
Dim scpFile As Scripting.File
Dim scpFiles As Scripting.Files
Dim scpFolder As Scripting.Folder

Set scpFolder = FSO.GetFolder(DefPath)
Set scpFiles = scpFolder.Files

Dim vcfFiles As Collection
For Each scpFile In scpFiles
    If scpFile.Type = "vCalendar File" Then
        If vcfFiles Is Nothing Then
            Set vcfFiles = New Collection
            vcfFiles.Add scpFile
        Else
            vcfFiles.Add scpFile
        End If
    End If
Next
Set fnGetFiles = vcfFiles
End Function
Private Function fnProcessFile(ByRef scpFile As Scripting.File, ByRef FSO As FileSystemObject) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : fnProcessFile
' Author    : John G Ehrhardt
' Date      : 12-Sep-2018
' Purpose   : Process a File
'---------------------------------------------------------------------------------------
'Create a Temp File based on processed file
Dim tempName As String

'Create a Temp file in same place, for write
tempName = scpFile.Path & ".tmp"
Dim tempfile As Scripting.TextStream
Set tempfile = FSO.CreateTextFile(tempName, True)

'Now Read the file and replace the string... Do only 1 Replacement!
Dim currentLine As String
Dim newline As String
Dim FindText As String
FindText = "VERSION:2.1"
Dim NewText As String
NewText = "VERSION:3.0"

Dim OrigFile As String
OrigFile = scpFile.Path
Dim File As Scripting.TextStream
Set File = FSO.OpenTextFile(OrigFile)
Dim endIt As Boolean

Do Until File.AtEndOfStream
    currentLine = File.ReadLine
    If Not endIt Then
        newline = Replace(currentLine, FindText, NewText)
        endIt = newline <> currentLine
        tempfile.WriteLine newline
    Else
        tempfile.WriteLine currentLine
    End If
Loop

'Close the Files
File.Close
tempfile.Close

'Replace the Original Files

FSO.DeleteFile OrigFile, True
FSO.MoveFile tempName, OrigFile

End Function

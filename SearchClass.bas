Option Explicit

Dim lnglngFileInspection As LongLong
Dim lnglngFilesRecorded As LongLong
Dim objFileSystem As Object
Dim objWScriptShell As Object
Dim startTime As Date
Dim dateDeadline As Date
Dim objResultsTable As Object
Dim objLoggerClass As Object

Sub Class_Initialize()
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objWScriptShell = CreateObject("Wscript.Shell")
    Set objResultsTable = ActiveWorkbook.Sheets("Results").ListObjects("ResultsTable")
    Set objLoggerClass = New LoggerClass
    
    startTime = Now
    dateDeadline = ActiveWorkbook.Sheets("Instructions").Range("deadline_value").Value
End Sub

Sub Class_Terminate()
    objResultsTable.Parent.Activate
    objLoggerClass.MakeEntry strAction:="Search", strDetails:="files=" & lnglngFileInspection & " results=" & lnglngFilesRecorded & " time=" & Format(Now - startTime, "hh:mm:ss")

    Set objFileSystem = Nothing
    Set objWScriptShell = Nothing
    Set objResultsTable = Nothing
    Set objLoggerClass = Nothing
End Sub

Public Sub MainSub()
    Dim strStartFolder As String: strStartFolder = "X:\Retic_Transfer\__Supplementary Files"
    
    Search objFileSystem.GetFolder(strStartFolder)
End Sub

'************** Search Sub
Sub Search(Folder)
    Dim SubFolder As Object
    Dim bolRecordResult As Boolean
    
    For Each SubFolder In Folder.SubFolders
        Search SubFolder
    Next
    
    Dim File As Object
    For Each File In Folder.Files
        lnglngFileInspection = lnglngFileInspection + 1
        If lnglngFileInspection Mod 5000 = 0 Then
            objLoggerClass.MakeEntry strAction:="Search", strDetails:="files=" & lnglngFileInspection & " results=" & lnglngFilesRecorded & " time=" & Format(Now - startTime, "hh:mm:ss")
        End If
        If Now > dateDeadline Then
            Application.StatusBar = "Deadline Exceeded"
            Exit Sub
        End If
        If getFileExt(objFileSystem.GetAbsolutePathName(File)) = "vob" Or getPerceivedType(getFileExt(objFileSystem.GetAbsolutePathName(File))) = "video" Then
            objFileSystem.GetAbsolutePathName (File)
            bolRecordResult = recordResultInTable(objFileSystem.GetAbsolutePathName(File))
            If Not (bolRecordResult) Then
                Stop
            End If
        End If
    Next
End Sub

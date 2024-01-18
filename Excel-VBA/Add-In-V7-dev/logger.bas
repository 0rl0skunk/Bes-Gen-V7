Attribute VB_Name = "logger"
Attribute VB_Description = "Logging Module."

'@Folder "Debug Logger"
'@IgnoreModule EmptyStringLiteral
'@ModuleDescription "Logging Module."
'@Version "Release V1.0.0"

Option Explicit

Private pLogFile          As String
Private pLogFolder As String

Public Enum ErrorLevel
    LogError = 0
    LogWarning = 1
    LogInfo = 2
    LogTrace = 3
End Enum

Const LogDepth = 3
' 3 = Trace
' 2 = Info
' 1 = Warnings
' 0 = Errors

Public Property Get LogFile() As String
LogFile = pLogFile
End Property

Public Sub writelog(ByVal Typ As ErrorLevel, ByVal a_stringLogThis As String)
    
    pLogFolder = Environ("localappdata") & "\Bes-Gen-V7"
    pLogFile = pLogFolder & "\Bes-Gen_V7.log"
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(pLogFolder) Then
    MkDir pLogFolder
    End If
    ' prepare date
    Dim l_StringDateTimeNow  As String
    Dim l_StringToday        As String
    Dim l_StringLogStatement As String
    Dim l_StringSource       As String
    Dim Typstr               As String

    l_StringSource = "  ADD-IN  "

    Select Case Typ
    Case 0
        If LogDepth >= 0 Then
            Typstr = "ERROR  "
        Else
            Exit Sub
        End If
    Case 1
        If LogDepth >= 1 Then
            Typstr = "WARNING"
        Else
            Exit Sub
        End If
    Case 2
        If LogDepth >= 2 Then
            Typstr = "INFO   "
        Else
            Exit Sub
        End If
    Case 3
        If LogDepth >= 3 Then
            Typstr = "TRACE  "
        Else
            Exit Sub
        End If
    End Select
    l_StringDateTimeNow = Now
    l_StringToday = Format$(l_StringDateTimeNow, "YYYY-MM-DD hh:mm:ss")
    ' concatenate date and what the user wants logged
    l_StringLogStatement = Join(Array(l_StringToday, l_StringSource, Typstr, a_stringLogThis), " | ")
    ' send to TTY
    Debug.Print (l_StringLogStatement)
    ' append (not write) to disk
    Open LogFile For Append As #1
    Print #1, l_StringLogStatement
    Close #1
End Sub

Public Sub LogClear()
    Debug.Print ("Erasing the previous logs.")
    Open LogFile For Output As #1
    Print #1, ""
    Close #1
End Sub



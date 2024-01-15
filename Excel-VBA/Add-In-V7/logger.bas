Attribute VB_Name = "logger"
Attribute VB_Description = "Logging Module."

'@Folder "Debug Logger"
'@IgnoreModule EmptyStringLiteral
'@ModuleDescription "Logging Module."
'@Version "Release V1.0.0"

Option Explicit

Public Const LogFile         As String = "C:\Users\Public\Documents\TinLine\Bes-Gen_V7.log"
Public Enum ErrorLevel
    LogError = 0
    Logwarning = 1
    LogInfo = 2
    LogTrace = 3
End Enum

Const LogDepth = 3
' 3 = Trace
' 2 = Info
' 1 = Warnings
' 0 = Errors

Public Sub writelog(ByVal Typ As ErrorLevel, ByVal a_stringLogThis As String)
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
    l_StringLogStatement = "| " & Join(Array(l_StringToday, l_StringSource, Typstr, a_stringLogThis), " | ") & " |"
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

Private Sub samples()
    'for error Logging:
    writelog LogError, "Where did the error occure?" & vbNewLine & _
                      err.Number & vbNewLine & err.Description & vbNewLine & err.source
End Sub



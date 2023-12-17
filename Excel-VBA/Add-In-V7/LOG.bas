Attribute VB_Name = "LOG"
Attribute VB_Description = "Logging Module."
Option Explicit
'@IgnoreModule EmptyStringLiteral
'@Folder "Debug Logger"
'@ModuleDescription "Logging Module."

Public Const LogFile         As String = "C:\Users\Public\Documents\TinLine\Bes-Gen_V7.log"
Public Enum ErrorLevel
    LogError = 0
    logwarning = 1
    LogInfo = 2
    LogTrace = 3
End Enum

Public Sub writelog(ByVal Typ As ErrorLevel, ByVal a_stringLogThis As String)
    ' prepare date
    Dim l_StringDateTimeNow  As String
    Dim l_StringToday        As String
    Dim l_StringLogStatement As String

    Dim Typstr               As String
    Select Case Typ
        Case 0
            If Globals.LogDepth >= 1 Then
                Typstr = ">> ERROR   "
            Else
                Exit Sub
            End If
        Case 1
            If Globals.LogDepth >= 2 Then
                Typstr = ">> WARNING "
            Else
                Exit Sub
            End If
        Case 2
            If Globals.LogDepth >= 3 Then
                Typstr = ">> INFO    "
            Else
                Exit Sub
            End If
        Case 3
            If Globals.LogDepth >= 3 Then
                Typstr = ">> TRACE   "
            Else
                Exit Sub
            End If
    End Select
    l_StringDateTimeNow = Now
    l_StringToday = Format$(l_StringDateTimeNow, "YYYY-MM-DD hh:mm:ss")
    ' concatenate date and what the user wants logged
    l_StringLogStatement = l_StringToday & " " & Typstr & a_stringLogThis
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
                     ERR.Number & vbNewLine & ERR.description & vbNewLine & ERR.source
End Sub



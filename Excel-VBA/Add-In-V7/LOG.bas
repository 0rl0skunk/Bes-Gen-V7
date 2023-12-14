Attribute VB_Name = "LOG"
Option Explicit
'@IgnoreModule EmptyStringLiteral
'@Folder "Debug Logger"
'@ModuleDescription "Logging Module."

Public Const LOGFile         As String = "C:\Users\Public\Documents\TinLine\Bes-Gen_V7.log"

Public Sub writelog(ByVal Typ As String, ByVal a_stringLogThis As String)
    ' prepare date
    Dim l_StringDateTimeNow  As String, _
    l_StringToday            As String, _
    l_StringLogStatement     As String
    Dim Typstr               As String
    Select Case Typ
        Case "Error"
            If LogDepth >= 1 Then
                Typstr = ">> ERROR   "
            Else
                Exit Sub
            End If
        Case "Warning"
            If LogDepth >= 1 Then
                Typstr = ">> WARNING "
            Else
                Exit Sub
            End If
        Case "Info"
            If LogDepth >= 3 Then
                Typstr = ">> INFO    "
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
    Open LOGFile For Append As #1
    Print #1, l_StringLogStatement
    Close #1
End Sub

Public Sub LogClear()
Debug.Print ("Erasing the previous logs.")
    Open LOGFile For Output As #1
    Print #1, ""
    Close #1
End Sub

Private Function samples() As String
    'for error Logging:
    writelog "Error", "Where did the error occure?" & vbNewLine & _
                     ERR.Number & vbNewLine & ERR.description & vbNewLine & ERR.source
End Function



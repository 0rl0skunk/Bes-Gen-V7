Attribute VB_Name = "LOG"
Option Explicit
'@IgnoreModule EmptyStringLiteral
'@Folder "Debug Logger"
Public Const LOGFile         As String = "C:\Users\Public\Documents\TinLine\Bes-Gen_V7.log"

Public Sub Log(ByVal a_stringLogThis As String)
    ' prepare date
    Dim l_StringDateTimeNow  As String, _
    l_StringToday            As String, _
    l_StringLogStatement     As String

    l_StringDateTimeNow = Now
    l_StringToday = Format$(l_StringDateTimeNow, "YYYY-MM-DD hh:mm:ss")
    ' concatenate date and what the user wants logged
    l_StringLogStatement = l_StringToday & " " & a_stringLogThis
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



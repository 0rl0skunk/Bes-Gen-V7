Attribute VB_Name = "INIFile"
'@IgnoreModule
'@Folder "TinLine"
Option Explicit
'declarations for working with Ini files
#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias _
                                     "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, _
                                                                  ByVal nSize As Long, ByVal lpFileName As String) As Long

    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias _
                                     "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                                                 ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
                                                                 ByVal lpFileName As String) As Long

    Private Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" Alias _
                                     "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, _
                                                                    ByVal lpFileName As String) As Long

    Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias _
                                     "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                                                                   ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

'// INI CONTROLLING PROCEDURES
'reads an Ini string
Public Function ReadIni(filename As String, Section As String, Key As String) As String
    Dim RetVal               As String * 1024, v As Long
    v = GetPrivateProfileString(Section, Key, vbNullString, RetVal, 1024, filename)
    ReadIni = Left(RetVal, v)
End Function

'reads an Ini section
Public Function ReadIniSection(filename As String, Section As String) As String
    Dim RetVal               As String * 1024, v As Long
    v = GetPrivateProfileSection(Section, RetVal, 1024, filename)
    ReadIniSection = Left(RetVal, v)
End Function

'writes an Ini string
Public Sub WriteIni(filename As String, Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, filename
End Sub

'writes an Ini section
Public Sub WriteIniSection(filename As String, Section As String, Value As String)
    WritePrivateProfileSection Section, Value, filename
End Sub



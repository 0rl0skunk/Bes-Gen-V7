Attribute VB_Name = "ClipboardWindowsAPI"
Attribute VB_Description = "Handle 64-bit and 32-bit Office"

'@Folder "Excel-Items"
'@Version "Release V1.0.0"
'@ModuleDescription "Handle 64-bit and 32-bit Office"

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, _
                                                             ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                         ByVal lpString2 As Any) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, _
                                                                ByVal hMem As LongPtr) As LongPtr
#Else
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                     ByVal dwBytes As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                 ByVal lpString2 As Any) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat _
                                                        As Long, ByVal hMem As Long) As Long
#End If

Const GHND = &H42
Const CF_TEXT = 1
Const MAXSIZE = 4096

Function CopyToClipBoard(MyString As String)
    'PURPOSE: API function to copy text to clipboard
    'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

#If VBA7 Then
    Dim hGlobalMemory        As LongPtr
    Dim lpGlobalMemory       As LongPtr

    Dim hClipMemory          As LongPtr
    Dim X                    As LongPtr

#Else
    Dim hGlobalMemory        As Long, lpGlobalMemory As Long
    Dim hClipMemory          As Long, X As Long
#End If

    'Allocate moveable global memory
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

    'Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    'Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    'Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    'Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    'Clear the Clipboard.
    X = EmptyClipboard()

    'Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If

End Function



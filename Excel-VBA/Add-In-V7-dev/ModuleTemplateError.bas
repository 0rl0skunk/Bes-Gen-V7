Attribute VB_Name = "ModuleTemplateError"
Attribute VB_Description = "Vorlage für Error-Handling."

'@Folder "Templates"
'@ModuleDescription "Vorlage für Error-Handling."
'@Version "Release V1.0.0"

Option Explicit

Public Const Dev             As Boolean = False

Public Sub TemplateErrorHandler()

    ' vv define a usefull error source vv
    Dim ErrSource            As String: ErrSource = "Module: ModuleTemplateError" & vbNewLine & _
        "Sub:    TemplateErrorHandler" & vbNewLine & _
        "Code:   "

    ' vv something could go wrong here vv
    If Not Dev Then On Error GoTo Err1           ' show the fancy error messages for the Users and the functional one for the developers
    writelog LogError, "trying to divide 9/0"
    Debug.Print 9 / 0

    GoTo noerr1

Err1:
    ' the "SOMETHING" happened
    On Error GoTo -1                             ' reset error code from excel to create a own one
    On Error GoTo ErrHandler                     ' goto the fancy error messages
    err.Raise 83, , "A good Description of what happened" & vbNewLine & _
                   "maybe even on two seperate Lines"

noerr1:
    ' it worked as intended

    Exit Sub

    ' ------ ERROR HANDLER ------
ErrHandler:

    Dim errFrm               As New UserFormMessage

    Select Case err.Number
    Case 81
        ' if it is solvable then do so and
        GoTo errSolved
    Case 82
        ' if it is NOT solvable then display the error message
        errFrm.Typ TypWarning, ErrSource & err.Number & vbNewLine & "Decsription:" & vbNewLine & err.Description
        errFrm.Show 1
    Case Else                                    'a "unhandled" error occured
        errFrm.Typ typError, ErrSource & err.Number & vbNewLine & "Decsription:" & vbNewLine & err.Description, , True
        errFrm.Show 1
    End Select

    ' error has been solved
errSolved:
    Exit Sub

End Sub



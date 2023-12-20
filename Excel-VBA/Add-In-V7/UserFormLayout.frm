VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormLayout 
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5880
   OleObjectBlob   =   "UserFormLayout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "Plankopf"
Option Explicit
Private pMasstab             As Integer
Private MultiPageType        As Integer
Private icons                As UserFormIconLibrary

Public Sub load(ByVal Format As String, ByVal Masstab As String, ByVal mpType As Integer)

    Me.TextBoxFormatH.value = Split(Format, "H")(0)
    Me.TextBoxFormatB.value = Split(Split(Format, "H")(1), "B")(0)
    MultiPageType = mpType

    pMasstab = CInt(Split(Masstab, ":")(1))

    Me.TextBoxMasstab.value = pMasstab

    ChangeFormat

End Sub

Private Sub ChangeFormat()

    ' --- change displayed Layout based on new inputs
    Dim paper                As control
    Dim border               As control
    Dim Plankopf             As control
    Dim legende              As control
    Dim maxHeight            As Double
    Dim maxWidth             As Double
    Dim height               As Double
    Dim width                As Double
    Dim tHeight              As Double
    Dim tWidth               As Double
    Dim H                    As Integer
    Dim B                    As Integer
    Dim ratio                As Double
    Dim mHeight              As Double
    Dim mWidth               As Double


    Set paper = Me.FramePaperSize
    Set border = Me.FramePaperBorder
    Set Plankopf = Me.FramePlankopf
    Set legende = Me.FrameLegende

    maxHeight = border.height - 12
    maxWidth = border.width - 12

    H = CInt(Me.TextBoxFormatH.value)
    B = CInt(Me.TextBoxFormatB.value)

    ' --- get height
    height = H * 29.7
    tHeight = height
    ' --- get width
    width = B * 21
    tWidth = width

    ' --- side Ratio H/W

    ratio = height / width

    If ratio > 1 Then
        ' --- Vertikal
        tWidth = 1 / ratio * maxHeight
        tHeight = maxHeight

        If tWidth > maxWidth Then
            tWidth = maxWidth
            tHeight = ratio * maxWidth
        End If
    Else
        ' --- Horizontal
        tWidth = maxWidth
        tHeight = ratio * maxWidth

        If tHeight > maxHeight Then
            tWidth = 1 / ratio * maxHeight
            tHeight = maxHeight
        End If
    End If

    paper.Top = (maxHeight - tHeight + 12) / 2
    paper.Left = (maxWidth - tWidth + 12) / 2

    paper.width = tWidth
    paper.height = tHeight
    Me.TextBoxFormatB.value = B
    Me.TextBoxFormatH.value = H

    Plankopf.width = (tWidth / B)
    Plankopf.Left = (tWidth / B) * (B - 1)

    legende.Top = 0

    Select Case H
        Case 1
            Select Case B
                Case 1
                    legende.Visible = False
                    Plankopf.height = (tHeight / H) / 3
                    Plankopf.Top = (tHeight / H) * (H - 1) + 2 * (tHeight / H) / 3

                    mHeight = ((height - 5 - 2) * pMasstab) / 100
                    mWidth = ((width - 2) * pMasstab) / 100
                Case 2
                    legende.Visible = False
                    Plankopf.height = (tHeight / H) / 3
                    Plankopf.Top = (tHeight / H) * (H - 1) + 2 * (tHeight / H) / 3
                    mHeight = ((height - 5 - 2) * pMasstab) / 100
                    mWidth = ((width - 2) * pMasstab) / 100
                Case Is >= 3
                    legende.Visible = True
                    legende.width = (tWidth / B)
                    legende.height = (tHeight / H)
                    legende.Left = (tWidth / B) * (B - 2)
                    Plankopf.height = tHeight / H
                    Plankopf.Top = (tHeight / H) * (H - 1) + (H - 1) * (tHeight / H)
                    mHeight = ((height - 2) * pMasstab) / 100
                    mWidth = ((width - 4) * pMasstab) / 100
            End Select
        Case 2
            Select Case B
                Case 1
                    legende.Visible = False
                    Plankopf.height = (tHeight / H) / 3
                    Plankopf.Top = (tHeight / H) * (H - 1) + 2 * (tHeight / H) / 3
                    mHeight = ((height - 5 - 2) * pMasstab) / 100
                    mWidth = ((width - 2) * pMasstab) / 100
                Case 2
                    legende.Visible = True
                    legende.width = (tWidth / B)
                    legende.height = (tHeight / H) * (H - 1)
                    legende.Left = (tWidth / B) * (B - 1)
                    Plankopf.height = tHeight / H
                    Plankopf.Top = (tHeight / H) + (H - 1)
                    mHeight = ((height - 2) * pMasstab) / 100
                    mWidth = ((width - 23) * pMasstab) / 100
                Case Is >= 3
                    legende.Visible = True
                    legende.width = (tWidth / B)
                    legende.height = (tHeight / H) * (H - 1)
                    legende.Left = (tWidth / B) * (B - 1)
                    Plankopf.height = tHeight / H
                    Plankopf.Top = (tHeight / H) * (H - 1)
                    mHeight = ((height - 2) * pMasstab) / 100
                    mWidth = ((width - 23) * pMasstab) / 100
            End Select
        Case 3
            Select Case B
                Case 1
                    legende.Visible = False
                    Plankopf.height = (tHeight / H) / 3
                    Plankopf.Top = (tHeight / H) * (H - 1) + 2 * (tHeight / H) / 3
                    mHeight = ((height - 5 - 2) * pMasstab) / 100
                    mWidth = ((width - 2) * pMasstab) / 100
                Case 2
                    legende.Visible = True
                    legende.width = (tWidth / B)
                    legende.height = (tHeight / H) * (H - 1)
                    legende.Left = (tWidth / B) * (B - 1)
                    Plankopf.height = tHeight / H
                    Plankopf.Top = (tHeight / H) * (H - 1)
                    mHeight = ((height - 2) * pMasstab) / 100
                    mWidth = ((width - 23) * pMasstab) / 100
                Case Is >= 3
                    legende.Visible = True
                    legende.width = (tWidth / B)
                    legende.height = (tHeight / H) * (H - 1)
                    legende.Left = (tWidth / B) * (B - 1)
                    Plankopf.height = tHeight / H
                    Plankopf.Top = (tHeight / H) * (H - 1)
                    mHeight = ((height - 2) * pMasstab) / 100
                    mWidth = ((width - 23) * pMasstab) / 100
            End Select
    End Select

    Me.TextBoxLayout.value = "Höhe:" & H & "H" & vbLf & _
                             "Beite:" & B & "B" & vbLf & _
                             height & "x" & width & "cm"

    Select Case MultiPageType
        Case 0                                   'Plan
            Me.TextBoxModell.value = "Modellbereich: " & vbLf & _
                                     "Höhe: " & mHeight & "m" & vbLf & _
                                     "Beite: " & mWidth & "m"
        Case 1                                   'Schema
            Me.TextBoxModell.value = "Modellbereich: " & vbLf & _
                                     "Höhe: " & mHeight & "m" & vbLf & _
                                     "Beite: " & mWidth & "m"
        Case 2                                   'Prinzip
            Me.TextBoxModell.value = "Modellbereich: " & vbLf & _
                                     "Höhe: " & Application.WorksheetFunction.RoundDown(mHeight / 3, 0) & " Geschosse" & vbLf & _
                                     "Beite: " & mWidth & "m"

        Case 3                                   'Detail
            Me.TextBoxModell.value = "Modellbereich: " & vbLf & _
                                     "Höhe: " & mHeight & "m" & vbLf & _
                                     "Beite: " & mWidth & "m"
    End Select

End Sub

Private Sub CommandButton2_Click()

    Me.CheckBoxLoad.value = True
    Unload Me

End Sub

Private Sub CommandButton3_Click()
    pMasstab = CInt(Me.TextBoxMasstab.value)

    ChangeFormat
End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub SpinButtonFormatB_SpinDown()
    If CInt(Me.TextBoxFormatB.value) - 1 <= 0 Then Exit Sub
    Me.TextBoxFormatB.value = Me.TextBoxFormatB.value - 1
    ChangeFormat
End Sub

Private Sub SpinButtonFormatB_SpinUp()
    If CInt(Me.TextBoxFormatB.value) + 1 > 20 Then Exit Sub
    Me.TextBoxFormatB.value = Me.TextBoxFormatB.value + 1
    ChangeFormat
End Sub

Private Sub SpinButtonFormatH_SpinDown()
    If CInt(Me.TextBoxFormatH.value) - 1 <= 0 Then Exit Sub
    Me.TextBoxFormatH.value = Me.TextBoxFormatH.value - 1
    ChangeFormat
End Sub

Private Sub SpinButtonFormatH_SpinUp()
    If CInt(Me.TextBoxFormatH.value) + 1 > 3 Then Exit Sub
    Me.TextBoxFormatH.value = Me.TextBoxFormatH.value + 1
    ChangeFormat
End Sub

Private Sub UserForm_Initialize()

    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconLayout.Picture

    Me.Caption = vbNullString
    Me.TitleLabel = "Layout Voransicht"
    Me.LabelInstructions.Caption = "Die Layoutgrösse im Modell und die Plankopfposition sowie die Standardlegenden können hier abgelesen werden."
End Sub



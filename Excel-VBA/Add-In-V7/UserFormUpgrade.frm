VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormUpgrade 
   Caption         =   "UserFormUpgrade"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "UserFormUpgrade.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Upgrade")

Option Explicit

Private icons As UserFormIconLibrary
Private WSOldVersion As Workbook

Private Sub CommandButtonClose_Click()
Unload Me
End Sub

Private Sub CommandButtonLoadOldVersion_Click()
Dim fDialog As FileDialog, result As Integer
Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
'Optional: FileDialog properties
fDialog.AllowMultiSelect = False
fDialog.Title = "Alte Version vom Beschriftungsgenerator auswählen"
fDialog.InitialFileName = "C:\"
'Optional: Add filters
fDialog.Filters.Clear
fDialog.Filters.Add "Excel files", "*.xlsx"
fDialog.Filters.Add "Excel files", "*.xlsm"
fDialog.Filters.Add "All files", "*.*"
 
'Show the dialog. -1 means success!
If fDialog.Show = -1 Then
   writelog LogInfo, fDialog.SelectedItems(1)
   Me.TextBox1.value = fDialog.SelectedItems(1)
   Set WSOldVersion = Application.Workbooks.Open(FileName:=Me.TextBox1.value, ReadOnly:=True)
End If

If Not WSOldVersion Is Nothing Then
' something was loaded
    ' try to automatically get the version
    Select Case WSOldVersion.Sheets("Projektdaten").range("B3").value
    Case 1
    Case ""
    Case 2
    Case 3
    Case Else
    End Select
End If

End Sub

Private Sub CommandButtonUpgrade_Click()
Upgrade Me.textboxversion.value
End Sub

Private Sub UserForm_Initialize()
    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconWindowsupdate.Picture
    Me.TitleLabel.Caption = "Beschriftungsgenerator Upgraden"
    Me.LabelInstructions.Caption = "vorherige Versionen auf V7 upgraden"
End Sub

Private Sub Upgrade(ByVal FromVersion As String)

Dim ws                   As Worksheet: Set ws = Globals.shStoreData

Select Case FromVersion
Case 1
' for each row in shStoreData transpose it to the new order
    With ws
        .Cells(row, 1).value = Plankopf.ID
        .Cells(row, 2).value = Plankopf.IDTinLine
        .Cells(row, 3).value = Plankopf.Gewerk
        .Cells(row, 4).value = Plankopf.UnterGewerk
        .Cells(row, 5).value = Plankopf.Planart
        .Cells(row, 6).value = Plankopf.Plantyp
        .Cells(row, 7).value = Plankopf.Gebäude
        .Cells(row, 8).value = Plankopf.Gebäudeteil
        .Cells(row, 9).value = Plankopf.Geschoss
        .Cells(row, 10).value = Plankopf.CustomPlanüberschrift
        .Cells(row, 11).value = Plankopf.dwgFile
        .Cells(row, 13).value = Plankopf.Planüberschrift
        .Cells(row, 14).value = Plankopf.Plannummer
        .Cells(row, 15).value = Plankopf.LayoutGrösse
        .Cells(row, 16).value = Plankopf.LayoutMasstab
        .Cells(row, 17).value = Plankopf.LayoutPlanstand
        .Cells(row, 18).value = Plankopf.GezeichnetPerson
        .Cells(row, 19).value = Replace(Plankopf.GezeichnetDatum, ".", "/")
        .Cells(row, 20).value = Plankopf.GeprüftPerson
        .Cells(row, 21).value = Replace(Plankopf.GeprüftDatum, ".", "/")
        .Cells(row, 12).value = Plankopf.CurrentIndex.Index
        .Cells(row, 22).value = Plankopf.TinLinePKNr
        .Cells(row, 23).value = Plankopf.AnlageTyp
        .Cells(row, 24).value = Plankopf.AnlageNummer
    End With
' for each row in shIndex transpose it to the new order
    With Globals.shIndex
        .Cells(row, 1).value = Index.PlanID
        .Cells(row, 2).value = Index.Index
        .Cells(row, 3).value = Split(Gezeichnet, ";")(0)
        .Cells(row, 4).value = Split(Gezeichnet, ";")(1)
        .Cells(row, 5).value = Split(Geprüft, ";")(0)
        .Cells(row, 6).value = Split(Geprüft, ";")(1)
        .Cells(row, 7).value = Index.Klartext
        .Cells(row, 8).value = Index.IndexID
    End With
' Transfer Projektdaten
' Transfer Gebäudedaten -> evtl. müssen diese von Hand noch angepasst / ausgefüllt werden.
Case 2
' for each row in shStoreData transpose it to the new order
' for each row in shIndex transpose it to the new order
' Transfer Projektdaten
' Transfer Gebäudedaten -> evtl. müssen diese von Hand noch angepasst / ausgefüllt werden.
Case 3
' for each row in shStoreData transpose it to the new order
' for each row in shIndex transpose it to the new order
' Transfer Projektdaten
' Transfer Gebäudedaten -> evtl. müssen diese von Hand noch angepasst / ausgefüllt werden.
Case 4
' for each row in shStoreData transpose it to the new order
' for each row in shIndex transpose it to the new order
' Transfer Projektdaten
' Transfer Gebäudedaten -> evtl. müssen diese von Hand noch angepasst / ausgefüllt werden.
End Select

End Sub

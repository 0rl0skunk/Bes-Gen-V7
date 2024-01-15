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
'@Version "Release V1.0.0"

Option Explicit

Private icons                As UserFormIconLibrary
Private WBOldVersion         As Workbook

Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

Private Sub CommandButtonLoadOldVersion_Click()
    Dim fDialog              As FileDialog
    Dim result               As Long

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    'Optional: FileDialog properties
    fDialog.AllowMultiSelect = False
    fDialog.Title = "Alte Version vom Beschriftungsgenerator ausw�hlen"
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
        Set WBOldVersion = Application.Workbooks.Open(FileName:=Me.TextBox1.value, ReadOnly:=True)
    End If

    If Not WBOldVersion Is Nothing Then
        ' something was loaded
        ' try to automatically get the version
        Select Case Left$(WBOldVersion.Sheets("Projektdaten").range("B4").value, 1)
        Case 1
            Me.OptionButton1.value = True
        Case vbNullString
            Me.OptionButton1.value = True
        Case 2
            Me.OptionButton2.value = True
        Case 3
            Me.OptionButton3.value = True
        Case 4
            Me.OptionButton4.value = True
        Case Else
        End Select
    End If

End Sub

Private Sub CommandButtonUpgrade_Click()

    Upgrade

End Sub

Private Sub UserForm_Initialize()
    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconWindowsupdate.Picture
    Me.TitleLabel.Caption = "Beschriftungsgenerator Upgraden"
    Me.LabelInstructions.Caption = "vorherige Versionen auf V7 upgraden"
End Sub

Private Sub Upgrade()

    Dim FromVersion          As String
    If Me.OptionButton1.value Then FromVersion = 1
    If Me.OptionButton2.value Then FromVersion = 2
    If Me.OptionButton3.value Then FromVersion = 3
    If Me.OptionButton4.value Then FromVersion = 4

    Dim PLANTYP              As String
    Dim row                  As Long
    Dim lastrow              As Long

    ' --- oldWorksheets ------------------------------------------------------------------------
    On Error Resume Next
    Dim shPDataOld           As Worksheet: Set shPDataOld = WBOldVersion.Sheets("Projektdaten")
    Dim shIndexOld           As Worksheet: Set shIndexOld = WBOldVersion.Sheets("Index")
    Dim shStoreDataOld       As Worksheet: Set shStoreDataOld = WBOldVersion.Sheets("Datenbank")
    Dim shGeb�udeOld         As Worksheet: Set shGeb�udeOld = WBOldVersion.Sheets("Geb�ude")
    Dim shAdresseOld         As Worksheet: Set shAdresseOld = WBOldVersion.Sheets("Adressverzeichnis")
    On Error GoTo 0
    ' --- check old Worksheets -----------------------------------------------------------------

    Select Case FromVersion
    Case 3
        ' for each row in shStoreData transpose it to the new order
        With Globals.shStoreData
            lastrow = shStoreDataOld.range("A2").CurrentRegion.rows.Count
            For row = 3 To lastrow
                ' f�r jede zeile welche verwendet wird in den Neuen Bes-Gen �bertragen.
                Select Case shStoreDataOld.Cells(row, 2).value
                Case 0
                    PLANTYP = "PLA"
                Case 1
                    PLANTYP = "SCH"
                Case 2
                    PLANTYP = "PRI"
                End Select
                .Cells(row, 1).value = shStoreDataOld.Cells(row, 1).value
                .Cells(row, 2).value = shStoreDataOld.Cells(row, 21).value
                .Cells(row, 3).value = shStoreDataOld.Cells(row, 6).value
                .Cells(row, 4).value = shStoreDataOld.Cells(row, 7).value
                .Cells(row, 5).value = shStoreDataOld.Cells(row, 9).value
                .Cells(row, 6).value = PLANTYP                            ' Muss etwas komplizierter generiert werden siehe oben
                .Cells(row, 7).value = shStoreDataOld.Cells(row, 4).value
                .Cells(row, 8).value = shStoreDataOld.Cells(row, 3).value
                .Cells(row, 9).value = shStoreDataOld.Cells(row, 30).value
                .Cells(row, 10).value = False
                .Cells(row, 11).value = vbNullString                      ' wird beim updaten vom Plankopf geschrieben
                .Cells(row, 13).value = shStoreDataOld.Cells(row, 29).value
                .Cells(row, 14).value = shStoreDataOld.Cells(row, 2).value
                .Cells(row, 15).value = shStoreDataOld.Cells(row, 20).value
                .Cells(row, 16).value = shStoreDataOld.Cells(row, 23).value
                .Cells(row, 17).value = shStoreDataOld.Cells(row, 8).value
                .Cells(row, 18).value = shStoreDataOld.Cells(row, 25).value
                .Cells(row, 19).value = Replace(shStoreDataOld.Cells(row, 26).value, ".", "/")
                .Cells(row, 20).value = shStoreDataOld.Cells(row, 27).value
                .Cells(row, 21).value = Replace(shStoreDataOld.Cells(row, 28).value, ".", "/")
                .Cells(row, 12).value = shStoreDataOld.Cells(row, 24).value
                '.Cells(row, 22).value = Plankopf.TinLinePKNr ' muss Manuell eingef�gt werden oder das xml muss ge�ffnet und durchsucht werden.
                .Cells(row, 23).value = shStoreDataOld.Cells(row, 14).value
                .Cells(row, 24).value = shStoreDataOld.Cells(row, 10).value
            Next row
        End With
        ' for each row in shIndex transpose it to the new order
        With Globals.shIndex
            lastrow = shIndexOld.range("A3").CurrentRegion.rows.Count
            For row = 3 To lastrow
                .Cells(row - 1, 1).value = shIndexOld.Cells(row, 2).value
                .Cells(row - 1, 2).value = shIndexOld.Cells(row, 3).value
                .Cells(row - 1, 3).value = shIndexOld.Cells(row, 5).value
                .Cells(row - 1, 4).value = shIndexOld.Cells(row, 4).value
                .Cells(row - 1, 5).value = vbNullString
                .Cells(row - 1, 6).value = vbNullString
                .Cells(row - 1, 7).value = shIndexOld.Cells(row, 6).value
                .Cells(row - 1, 8).value = shIndexOld.Cells(row, 1).value
            Next row
        End With
        ' Transfer Projektdaten
        With Globals.shPData
            .range("ADM_Projektnummer").value = shPDataOld.range("C5").value
            .range("ADM_ADR_Strasse").value = shPDataOld.range("F5").value
            .range("ADM_ADR_PLZ").value = shPDataOld.range("F6").value
            .range("ADM_ADR_Ort").value = shPDataOld.range("F7").value
            .range("ADM_Projektbezeichnung").value = shPDataOld.range("C6").value
            .range("ADM_Projektphase").value = shPDataOld.range("C7").value
            .range("ADM_ProjektpfadSharePoint").value = "SherePoint Link ausf�llen"
            .range("ADM_ProjektPfadCAD").value = shPDataOld.range("C8").value
            ' UnterProjekte
            .range("PRO_Unterprojekte") = shPDataOld.range("Unterprojekte")
        End With
        ' Transfer Geb�udedaten -> evtl. m�ssen diese von Hand noch angepasst / ausgef�llt werden.
        Globals.shPData.range("A14:A50") = shPDataOld.range("A13:A49")    ' Geb�udeteil
        Globals.shGeb�ude.range("D1:AQ95") = shGeb�udeOld.range("D1:AQ95") ' Geb�ude
        Globals.shGeb�ude.range("B6:AQ95") = shGeb�udeOld.range("B6:AQ95") ' Geschosse
        ' Transfer Adressen
        With Globals.shAdress
            lastrow = shAdresseOld.range("A3").CurrentRegion.rows.Count
            For row = 6 To lastrow
                .Cells(row, 1).value = shAdresseOld.Cells(row, 1).value
                .Cells(row, 2).value = shAdresseOld.Cells(row, 2).value
                .Cells(row, 3).value = shAdresseOld.Cells(row, 3).value
                .Cells(row, 4).value = shAdresseOld.Cells(row, 4).value
                .Cells(row, 5).value = shAdresseOld.Cells(row, 5).value
                .Cells(row, 6).value = shAdresseOld.Cells(row, 6).value
                .Cells(row, 7).value = shAdresseOld.Cells(row, 7).value
            Next row
        End With
    Case 1
        ' for each row in shStoreData transpose it to the new order
        With Globals.shStoreData
            lastrow = shStoreDataOld.range("A2").CurrentRegion.rows.Count
            For row = 3 To lastrow
                ' f�r jede zeile welche verwendet wird in den Neuen Bes-Gen �bertragen.
                Select Case shStoreDataOld.Cells(row, 15).value
                Case 0
                    PLANTYP = "PLA"
                Case 1
                    PLANTYP = "SCH"
                Case 2
                    PLANTYP = "PRI"
                End Select
                .Cells(row, 1).value = shStoreDataOld.Cells(row, 1).value
                .Cells(row, 2).value = shStoreDataOld.Cells(row, 21).value
                .Cells(row, 3).value = shStoreDataOld.Cells(row, 6).value
                .Cells(row, 4).value = shStoreDataOld.Cells(row, 7).value
                .Cells(row, 5).value = shStoreDataOld.Cells(row, 9).value
                .Cells(row, 6).value = PLANTYP                            ' Muss etwas komplizierter generiert werden siehe oben
                .Cells(row, 7).value = shStoreDataOld.Cells(row, 3).value
                .Cells(row, 8).value = shStoreDataOld.Cells(row, 3).value
                .Cells(row, 9).value = shStoreDataOld.Cells(row, 4).value
                .Cells(row, 10).value = False
                .Cells(row, 11).value = vbNullString                      ' wird beim updaten vom Plankopf geschrieben
                .Cells(row, 13).value = shStoreDataOld.Cells(row, 29).value
                .Cells(row, 14).value = shStoreDataOld.Cells(row, 2).value
                .Cells(row, 15).value = shStoreDataOld.Cells(row, 20).value
                .Cells(row, 16).value = shStoreDataOld.Cells(row, 23).value
                .Cells(row, 17).value = shStoreDataOld.Cells(row, 8).value
                .Cells(row, 18).value = shStoreDataOld.Cells(row, 25).value
                .Cells(row, 19).value = Replace(shStoreDataOld.Cells(row, 26).value, ".", "/")
                .Cells(row, 20).value = shStoreDataOld.Cells(row, 27).value
                .Cells(row, 21).value = Replace(shStoreDataOld.Cells(row, 28).value, ".", "/")
                .Cells(row, 12).value = shStoreDataOld.Cells(row, 24).value
                '.Cells(row, 22).value = Plankopf.TinLinePKNr ' muss Manuell eingef�gt werden oder das xml muss ge�ffnet und durchsucht werden.
                .Cells(row, 23).value = shStoreDataOld.Cells(row, 14).value
                .Cells(row, 24).value = shStoreDataOld.Cells(row, 10).value
            Next row
        End With
        ' for each row in shIndex transpose it to the new order
        With Globals.shIndex
            lastrow = shIndexOld.range("A3").CurrentRegion.rows.Count
            For row = 3 To lastrow
                .Cells(row - 1, 1).value = shIndexOld.Cells(row, 2).value
                .Cells(row - 1, 2).value = shIndexOld.Cells(row, 3).value
                .Cells(row - 1, 3).value = shIndexOld.Cells(row, 5).value
                .Cells(row - 1, 4).value = shIndexOld.Cells(row, 4).value
                .Cells(row - 1, 5).value = vbNullString
                .Cells(row - 1, 6).value = vbNullString
                .Cells(row - 1, 7).value = shIndexOld.Cells(row, 6).value
                .Cells(row - 1, 8).value = shIndexOld.Cells(row, 1).value
            Next row
        End With
        ' Transfer Projektdaten
        With Globals.shPData
            .range("ADM_Projektnummer").value = shPDataOld.range("C5").value
            .range("ADM_ADR_Strasse").value = shPDataOld.range("F5").value
            .range("ADM_ADR_PLZ").value = shPDataOld.range("F6").value
            .range("ADM_ADR_Ort").value = shPDataOld.range("F7").value
            .range("ADM_Projektbezeichnung").value = shPDataOld.range("C6").value
            .range("ADM_Projektphase").value = shPDataOld.range("C7").value
            .range("ADM_ProjektpfadSharePoint").value = "SherePoint Link ausf�llen"
            .range("ADM_ProjektPfadCAD").value = shPDataOld.range("C8").value
            ' UnterProjekte
            .range("PRO_Unterprojekte") = shPDataOld.range("Unterprojekte")
        End With
        ' Transfer Geb�udedaten -> evtl. m�ssen diese von Hand noch angepasst / ausgef�llt werden.
        Globals.shPData.range("A14:A50") = shPDataOld.range("A13:A49")    ' Geb�udeteil
        Globals.shGeb�ude.range("D1:AQ95") = shGeb�udeOld.range("D1:AQ95") ' Geb�ude
        Globals.shGeb�ude.range("B6:AQ95") = shGeb�udeOld.range("B6:AQ95") ' Geschosse
        ' Transfer Adressen
        With Globals.shAdress
            lastrow = shAdresseOld.range("A3").CurrentRegion.rows.Count
            For row = 6 To lastrow
                .Cells(row, 1).value = shAdresseOld.Cells(row, 1).value
                .Cells(row, 2).value = shAdresseOld.Cells(row, 2).value
                .Cells(row, 3).value = shAdresseOld.Cells(row, 3).value
                .Cells(row, 4).value = shAdresseOld.Cells(row, 4).value
                .Cells(row, 5).value = shAdresseOld.Cells(row, 5).value
                .Cells(row, 6).value = shAdresseOld.Cells(row, 6).value
                .Cells(row, 7).value = shAdresseOld.Cells(row, 7).value
            Next row
        End With
    Case 2
        ' for each row in shStoreData transpose it to the new order
        With Globals.shStoreData
            lastrow = shStoreDataOld.range("A2").CurrentRegion.rows.Count
            For row = 3 To lastrow
                ' f�r jede zeile welche verwendet wird in den Neuen Bes-Gen �bertragen.
                Select Case shStoreDataOld.Cells(row, 15).value
                Case 0
                    PLANTYP = "PLA"
                Case 1
                    PLANTYP = "SCH"
                Case 2
                    PLANTYP = "PRI"
                End Select
                .Cells(row, 1).value = shStoreDataOld.Cells(row, 1).value
                .Cells(row, 2).value = shStoreDataOld.Cells(row, 21).value
                .Cells(row, 3).value = shStoreDataOld.Cells(row, 6).value
                .Cells(row, 4).value = shStoreDataOld.Cells(row, 7).value
                .Cells(row, 5).value = shStoreDataOld.Cells(row, 9).value
                .Cells(row, 6).value = PLANTYP                            ' Muss etwas komplizierter generiert werden siehe oben
                .Cells(row, 7).value = shStoreDataOld.Cells(row, 3).value
                .Cells(row, 8).value = shStoreDataOld.Cells(row, 3).value
                .Cells(row, 9).value = shStoreDataOld.Cells(row, 4).value
                .Cells(row, 10).value = False
                .Cells(row, 11).value = vbNullString                      ' wird beim updaten vom Plankopf geschrieben
                .Cells(row, 13).value = shStoreDataOld.Cells(row, 29).value
                .Cells(row, 14).value = shStoreDataOld.Cells(row, 2).value
                .Cells(row, 15).value = shStoreDataOld.Cells(row, 20).value
                .Cells(row, 16).value = shStoreDataOld.Cells(row, 23).value
                .Cells(row, 17).value = shStoreDataOld.Cells(row, 8).value
                .Cells(row, 18).value = shStoreDataOld.Cells(row, 25).value
                .Cells(row, 19).value = Replace(shStoreDataOld.Cells(row, 26).value, ".", "/")
                .Cells(row, 20).value = shStoreDataOld.Cells(row, 27).value
                .Cells(row, 21).value = Replace(shStoreDataOld.Cells(row, 28).value, ".", "/")
                .Cells(row, 12).value = shStoreDataOld.Cells(row, 24).value
                '.Cells(row, 22).value = Plankopf.TinLinePKNr ' muss Manuell eingef�gt werden oder das xml muss ge�ffnet und durchsucht werden.
                .Cells(row, 23).value = shStoreDataOld.Cells(row, 14).value
                .Cells(row, 24).value = shStoreDataOld.Cells(row, 10).value
            Next row
        End With
        ' for each row in shIndex transpose it to the new order
        With Globals.shIndex
            lastrow = shIndexOld.range("A3").CurrentRegion.rows.Count
            For row = 3 To lastrow
                .Cells(row - 1, 1).value = shIndexOld.Cells(row, 2).value
                .Cells(row - 1, 2).value = shIndexOld.Cells(row, 3).value
                .Cells(row - 1, 3).value = shIndexOld.Cells(row, 5).value
                .Cells(row - 1, 4).value = shIndexOld.Cells(row, 4).value
                .Cells(row - 1, 5).value = vbNullString
                .Cells(row - 1, 6).value = vbNullString
                .Cells(row - 1, 7).value = shIndexOld.Cells(row, 6).value
                .Cells(row - 1, 8).value = shIndexOld.Cells(row, 1).value
            Next row
        End With
        ' Transfer Projektdaten
        With Globals.shPData
            .range("ADM_Projektnummer").value = shPDataOld.range("C5").value
            .range("ADM_ADR_Strasse").value = shPDataOld.range("F5").value
            .range("ADM_ADR_PLZ").value = shPDataOld.range("F6").value
            .range("ADM_ADR_Ort").value = shPDataOld.range("F7").value
            .range("ADM_Projektbezeichnung").value = shPDataOld.range("C6").value
            .range("ADM_Projektphase").value = shPDataOld.range("C7").value
            .range("ADM_ProjektpfadSharePoint").value = "SherePoint Link ausf�llen"
            .range("ADM_ProjektPfadCAD").value = shPDataOld.range("C8").value
            ' UnterProjekte
            .range("PRO_Unterprojekte") = shPDataOld.range("Unterprojekte")
        End With
        ' Transfer Geb�udedaten -> evtl. m�ssen diese von Hand noch angepasst / ausgef�llt werden.
        Globals.shPData.range("A14:A50") = shPDataOld.range("A13:A49")    ' Geb�udeteil
        Globals.shGeb�ude.range("D1:AQ95") = shGeb�udeOld.range("D1:AQ95") ' Geb�ude
        Globals.shGeb�ude.range("B6:AQ95") = shGeb�udeOld.range("B6:AQ95") ' Geschosse
        ' Transfer Adressen
        With Globals.shAdress
            lastrow = shAdresseOld.range("A3").CurrentRegion.rows.Count
            For row = 6 To lastrow
                .Cells(row, 1).value = shAdresseOld.Cells(row, 1).value
                .Cells(row, 2).value = shAdresseOld.Cells(row, 2).value
                .Cells(row, 3).value = shAdresseOld.Cells(row, 3).value
                .Cells(row, 4).value = shAdresseOld.Cells(row, 4).value
                .Cells(row, 5).value = shAdresseOld.Cells(row, 5).value
                .Cells(row, 6).value = shAdresseOld.Cells(row, 6).value
                .Cells(row, 7).value = shAdresseOld.Cells(row, 7).value
            Next row
        End With
    Case 4
        ' for each row in shStoreData transpose it to the new order
        With Globals.shStoreData
            lastrow = shStoreDataOld.range("A2").CurrentRegion.rows.Count
            For row = 3 To lastrow
                ' f�r jede zeile welche verwendet wird in den Neuen Bes-Gen �bertragen.
                Select Case shStoreDataOld.Cells(row, 2).value
                Case 0
                    PLANTYP = "PLA"
                Case 1
                    PLANTYP = "SCH"
                Case 2
                    PLANTYP = "PRI"
                End Select
                .Cells(row, 1).value = shStoreDataOld.Cells(row, 1).value
                .Cells(row, 2).value = shStoreDataOld.Cells(row, 21).value
                .Cells(row, 3).value = shStoreDataOld.Cells(row, 6).value
                .Cells(row, 4).value = shStoreDataOld.Cells(row, 7).value
                .Cells(row, 5).value = shStoreDataOld.Cells(row, 9).value
                .Cells(row, 6).value = PLANTYP                            ' Muss etwas komplizierter generiert werden siehe oben
                .Cells(row, 7).value = shStoreDataOld.Cells(row, 4).value
                .Cells(row, 8).value = shStoreDataOld.Cells(row, 3).value
                .Cells(row, 9).value = shStoreDataOld.Cells(row, 30).value
                .Cells(row, 10).value = False
                .Cells(row, 11).value = vbNullString                      ' wird beim updaten vom Plankopf geschrieben
                .Cells(row, 13).value = shStoreDataOld.Cells(row, 29).value
                .Cells(row, 14).value = shStoreDataOld.Cells(row, 2).value
                .Cells(row, 15).value = shStoreDataOld.Cells(row, 20).value
                .Cells(row, 16).value = shStoreDataOld.Cells(row, 23).value
                .Cells(row, 17).value = shStoreDataOld.Cells(row, 8).value
                .Cells(row, 18).value = shStoreDataOld.Cells(row, 25).value
                .Cells(row, 19).value = Replace(shStoreDataOld.Cells(row, 26).value, ".", "/")
                .Cells(row, 20).value = shStoreDataOld.Cells(row, 27).value
                .Cells(row, 21).value = Replace(shStoreDataOld.Cells(row, 28).value, ".", "/")
                .Cells(row, 12).value = shStoreDataOld.Cells(row, 24).value
                '.Cells(row, 22).value = Plankopf.TinLinePKNr ' muss Manuell eingef�gt werden oder das xml muss ge�ffnet und durchsucht werden.
                .Cells(row, 23).value = shStoreDataOld.Cells(row, 14).value
                .Cells(row, 24).value = shStoreDataOld.Cells(row, 10).value
            Next row
        End With
        ' for each row in shIndex transpose it to the new order
        With Globals.shIndex
            lastrow = shIndexOld.range("A3").CurrentRegion.rows.Count
            For row = 3 To lastrow
                .Cells(row - 1, 1).value = shIndexOld.Cells(row, 2).value
                .Cells(row - 1, 2).value = shIndexOld.Cells(row, 3).value
                .Cells(row - 1, 3).value = shIndexOld.Cells(row, 5).value
                .Cells(row - 1, 4).value = shIndexOld.Cells(row, 4).value
                .Cells(row - 1, 5).value = vbNullString
                .Cells(row - 1, 6).value = vbNullString
                .Cells(row - 1, 7).value = shIndexOld.Cells(row, 6).value
                .Cells(row - 1, 8).value = shIndexOld.Cells(row, 1).value
            Next row
        End With
        ' Transfer Projektdaten
        With Globals.shPData
            .range("ADM_Projektnummer").value = shPDataOld.range("C5").value
            .range("ADM_ADR_Strasse").value = shPDataOld.range("F5").value
            .range("ADM_ADR_PLZ").value = shPDataOld.range("F6").value
            .range("ADM_ADR_Ort").value = shPDataOld.range("F7").value
            .range("ADM_Projektbezeichnung").value = shPDataOld.range("C6").value
            .range("ADM_Projektphase").value = shPDataOld.range("C7").value
            .range("ADM_ProjektpfadSharePoint").value = "SherePoint Link ausf�llen"
            .range("ADM_ProjektPfadCAD").value = shPDataOld.range("C8").value
            ' UnterProjekte
            .range("PRO_Unterprojekte") = shPDataOld.range("Unterprojekte")
        End With
        ' Transfer Geb�udedaten -> evtl. m�ssen diese von Hand noch angepasst / ausgef�llt werden.
        Globals.shPData.range("A14:A50") = shPDataOld.range("A13:A49")    ' Geb�udeteil
        Globals.shGeb�ude.range("D1:AQ95") = shGeb�udeOld.range("D1:AQ95") ' Geb�ude
        Globals.shGeb�ude.range("B6:AQ95") = shGeb�udeOld.range("B6:AQ95") ' Geschosse
        ' Transfer Adressen
        With Globals.shAdress
            lastrow = shAdresseOld.range("A3").CurrentRegion.rows.Count
            For row = 6 To lastrow
                .Cells(row, 1).value = shAdresseOld.Cells(row, 1).value
                .Cells(row, 2).value = shAdresseOld.Cells(row, 2).value
                .Cells(row, 3).value = shAdresseOld.Cells(row, 3).value
                .Cells(row, 4).value = shAdresseOld.Cells(row, 4).value
                .Cells(row, 5).value = shAdresseOld.Cells(row, 5).value
                .Cells(row, 6).value = shAdresseOld.Cells(row, 6).value
                .Cells(row, 7).value = shAdresseOld.Cells(row, 7).value
            Next row
        End With
    End Select

End Sub


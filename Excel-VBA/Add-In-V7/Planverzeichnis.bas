Attribute VB_Name = "Planverzeichnis"

'@Folder("Planverzeichnis")
'@Version "Release V1.0.0"

Option Explicit

Private Enum Sorting
    ELEPLA = 0
    ELESCH = 1
    ELEPRI = 2
    GWKPLA = 3
    GWKSCH = 4
    GWKPRI = 5
    KOOPLA = 6
    KOOSCH = 7
    KOOPRI = 8
    HKAPLA = 9
    HKASCH = 10
    HKAPRI = 11
    KAEPLA = 12
    KAESCH = 13
    KAEPRI = 14
    LUEPLA = 15
    LUESCH = 16
    LUEPRI = 17
    GAMPLA = 18
    GAMSCH = 19
    GAMPRI = 20
    SANPLA = 21
    SANSCH = 22
    SANPRI = 23
    SPRPLA = 24
    SPRSCH = 25
    SPRPRI = 26
    XXXPLA = 27
    XXXSCH = 28
    XXXPRI = 29
    TUEPLA = 30
    TUESCH = 31
    TUEPRI = 32
    BRAPLA = 33
    BRASCH = 34
    BRAPRI = 35
End Enum

Public Sub Create()
    Dim Plankopf             As IPlankopf
    Dim TempWS               As Worksheet
    Set TempWS = Application.ActiveWorkbook.Worksheets.Add
    Dim row                  As Long
    row = 1
    For Each Plankopf In Globals.Planköpfe
        With TempWS
            Select Case Plankopf.PLANTYP
            Case "PLA"
                .Cells(row, 1).value = Plankopf.Plannummer
                .Cells(row, 2).value = Plankopf.Planart
                .Cells(row, 3).value = Plankopf.UnterGewerk
                .Cells(row, 4).value = Plankopf.Geschoss
                .Cells(row, 5).value = Plankopf.LayoutMasstab
            Case "SCH"
                .Cells(row, 1).value = Plankopf.Plannummer
                .Cells(row, 2).value = Plankopf.UnterGewerk
            Case "PRI"
                .Cells(row, 1).value = Plankopf.Plannummer
                .Cells(row, 2).value = Plankopf.UnterGewerk
            End Select
            row = row + 1
            .Cells(row, 1).value = Plankopf.Plannummer
        End With
    Next
    TempWS.Delete
End Sub


Attribute VB_Name = "Helper"
Attribute VB_Description = "Beinhaltet n�tzliche Funktionen welche nicht einem Modul zugeordnet werden k�nnen."
'@IgnoreModule VariableNotUsed
Option Explicit
Public Enum IDType
    IDPlankopf = 0
    IDIndex = 1
    IDTask = 2
    IDPerson = 3
End Enum

'@ModuleDescription "Beinhaltet n�tzliche Funktionen welche nicht einem Modul zugeordnet werden k�nnen."
Public Function GetPlanartNamedRange(Planart As String, Hauptgewerk As String) As String
    ' Gibt die Range der verschiedenen Planarten des aktuellen Hauptgewerkes zur�ck
    Dim result               As String

    Select Case Hauptgewerk
        Case "Elektro"
            result = "ELE_Planart"
        Case "Gewerbliche K�lte"
            result = "GWK_Planart"
        Case "Koordination"
            result = "KOO_Planart"
        Case "Heizung K�lte"
            result = "HKA_Planart"
        Case "K�lte"
            result = "KAE_Planart"
        Case "L�ftung"
            result = "LUE_Planart"
        Case "Geb�udeautomation"
            result = "GAM_Planart"
        Case "Sanit�r"
            result = "SAN_Planart"
        Case "Sprinkler"
            result = "SPR_Planart"
        Case "HLKS/GA Allgemein"
            result = "XXX_Planart"
        Case "T�rfachplanung"
            result = "TUE_Planart"
        Case "Brandschutzplanung"
            result = "BRA_Planart"
    End Select

    GetPlanartNamedRange = result

End Function

Public Function GetUnterGewerkKF(UnterGewerk As String, Hauptgewerk As String, Planart As String) As String
    ' Gibt die Kurzform des Untergewerke zur�ck
    Dim result               As String
    Select Case Hauptgewerk
        Case "Elektro"
            Select Case Planart
                Case "Plan"
                    result = "ELE" & "_PLA"
                Case "Schema"
                    result = "ELE" & "_SCH"
                Case "Prinzip"
                    result = "ELE" & "_PRI"
            End Select
        Case "Gewerbliche K�lte"
            Select Case Planart
                Case "Plan"
                    result = "GWK" & "_PLA"
                Case "Schema"
                    result = "GWK" & "_SCH"
                Case "Prinzip"
                    result = "GWK" & "_PRI"
            End Select
        Case "Koordination"
            Select Case Planart
                Case "Plan"
                    result = "KOO" & "_PLA"
                Case "Schema"
                    result = "KOO" & "_SCH"
                Case "Prinzip"
                    result = "KOO" & "_PRI"
            End Select
        Case "Heizung K�lte"
            Select Case Planart
                Case "Plan"
                    result = "HKA" & "_PLA"
                Case "Schema"
                    result = "HKA" & "_SCH"
                Case "Prinzip"
                    result = "HKA" & "_PRI"
            End Select
        Case "K�lte"
            Select Case Planart
                Case "Plan"
                    result = "KAE" & "_PLA"
                Case "Schema"
                    result = "KAE" & "_SCH"
                Case "Prinzip"
                    result = "KAE" & "_PRI"
            End Select
        Case "L�ftung"
            Select Case Planart
                Case "Plan"
                    result = "LUE" & "_PLA"
                Case "Schema"
                    result = "LUE" & "_SCH"
                Case "Prinzip"
                    result = "LUE" & "_PRI"
            End Select
        Case "Geb�udeautomation"
            Select Case Planart
                Case "Plan"
                    result = "GAM" & "_PLA"
                Case "Schema"
                    result = "GAM" & "_SCH"
                Case "Prinzip"
                    result = "GAM" & "_PRI"
            End Select
        Case "Sanit�r"
            Select Case Planart
                Case "Plan"
                    result = "SAN" & "_PLA"
                Case "Schema"
                    result = "SAN" & "_SCH"
                Case "Prinzip"
                    result = "SAN" & "_PRI"
            End Select
        Case "Sprinkler"
            Select Case Planart
                Case "Plan"
                    result = "SPR" & "_PLA"
                Case "Schema"
                    result = "SPR" & "_SCH"
                Case "Prinzip"
                    result = "SPR" & "_PRI"
            End Select
        Case "HLKS/GA Allgemein"
            Select Case Planart
                Case "Plan"
                    result = "XXX" & "_PLA"
                Case "Schema"
                    result = "XXX" & "_SCH"
                Case "Prinzip"
                    result = "XXX" & "_PRI"
            End Select
        Case "T�rfachplanung"
            Select Case Planart
                Case "Plan"
                    result = "TUE" & "_PLA"
                Case "Schema"
                    result = "TUE" & "_SCH"
                Case "Prinzip"
                    result = "TUE" & "_PRI"
            End Select
        Case "Brandschutzplanung"
            Select Case Planart
                Case "Plan"
                    result = "BRA" & "_PLA"
                Case "Schema"
                    result = "BRA" & "_SCH"
                Case "Prinzip"
                    result = "BRA" & "_PRI"
            End Select
    End Select

    GetUnterGewerkKF = WLookup(UnterGewerk, shPData.range(result), 2)

End Function

Public Function WLookup(Lookup As Variant, range As range, Index As Integer, Optional onError As String = "-") As String
    ' VLookup mit 'onError' wert welcher selbst zugeordnet werden kann.
    On Error GoTo ERR

    Lookup = CStr(Lookup)

    If IsError(Application.VLookup(Lookup, range, Index, False)) Then
        WLookup = onError
    Else
        WLookup = Application.VLookup(Lookup, range, Index, False)
    End If

    Exit Function
    writelog LogInfo, "Wlookup Value Found " & WLookup

ERR:

    WLookup = onError
    writelog LogError, "Wlookup Value for " & Lookup & " Not Found"

End Function

Public Sub deleteIndexesXml()

    Dim FSO                  As Object
    Dim lastrow              As Integer
    Dim row                  As Integer
    Dim col                  As Integer
    Dim lastcol              As Integer


    Set FSO = CreateObject("scripting.FileSystemObject")

    lastrow = shGeb�ude.Cells(shGeb�ude.rows.Count, 2).End(xlUp).row
    lastcol = shGeb�ude.Cells(1, shGeb�ude.Columns.Count).End(xlToLeft).Column

    For col = 2 To lastcol Step 2
        For row = 6 To lastrow
            writelog LogInfo, "> empty cell " & row & " " & col & " " & IsEmpty(shGeb�ude.Cells(row, col)) & " " & shGeb�ude.Cells(row, col).Address
            If Not IsEmpty(shGeb�ude.Cells(row, col)) Then
                writelog LogInfo, "> " & shGeb�ude.Cells(row, col).Value
                ' only go if there is something in the cell
                deleteIndexXml row, col
            End If
        Next row
    Next col

    ' ----------------------------- PRINZIPSCHEMAS
    Dim j                    As Integer
    Dim i                    As Integer
    Dim filename             As String
    Dim TinLine              As String
    Dim Projektname          As String
    Dim Projektpfad          As String
    Dim EP                   As String
    Dim Gewerk               As String
    Dim GewerkNr             As String
    Dim rng                  As range
    Dim arr()                As Variant
    Dim tmparr()             As Variant


    Set rng = shPData.range("ELE_PRI")
    arr() = rng.Resize(rng.rows.Count, 1)
    tmparr() = RemoveBlanksFromStringArray(arr())

    Projektname = shPData.range("ADM_Projektnummer") & "_" & shPData.range("ADM_Projektbezeichnung")

    For i = LBound(tmparr) To UBound(tmparr)     ' for every Prinzipschema
        Gewerk = rng.Find(tmparr(i)).Offset(0, 1).Value
        GewerkNr = rng.Find(tmparr(i)).Offset(0, 2).Value
        If Len(GewerkNr) < 2 Then
            GewerkNr = "0" & GewerkNr
        End If
        deleteIndexXml 0, 0, shPData.range("ADM_ProjektpfadCAD").Value & "\03_PR\" & GewerkNr & "_" & Gewerk & "\TinPlan_PR_" & Gewerk & ".xml"
    Next i

End Sub

Public Sub deleteIndexXml(row As Integer, col As Integer, Optional i_xmlfile As String = vbNullString)

    Dim XMLFile              As String
    If i_xmlfile <> vbNullString Then
        XMLFile = i_xmlfile
    Else
        XMLFile = i_xmlfile
    End If

    Dim oXml                 As New MSXML2.DOMDocument60
    oXml.load XMLFile
    Dim nodes                As IXMLDOMNodeList
    Dim node                 As IXMLDOMNode
    Dim root                 As IXMLDOMNode

    Set root = oXml.SelectSingleNode("//tinPlan1")
    Set nodes = oXml.SelectNodes("//tinPlan1/*[contains(local-name(), 'IN')]")
    For Each node In nodes
        root.RemoveChild node
    Next
    On Error Resume Next
    oXml.Save XMLFile
    Set oXml = Nothing
    On Error GoTo 0

End Sub

Public Function getXML(PCol As Collection) As String
    ' get the xml file path for the genearted PK

    Dim Projektpath          As String
    Dim result               As String

    On Error GoTo ErrHandler
    Projektpath = shPData.range("ADM_ProjektpfadCAD")
    'On Error Resume Next
    Dim buildings            As Boolean
    buildings = Not (shGeb�ude.range("D1").Value = vbNullString)
    If PCol(1) = 0 Then
        ' Plan
        If PCol.Count > 7 Then
            If PCol(7)(1) = "DE" Then
                ' Detail
                result = Projektpath & "\04_DE\TinPlan_DE_" & PCol(15) & ".xml"
            Else
                GoTo Plan
            End If
        Else
Plan:
            If PCol(6)(1) = "TUE" Then
                ' --- TF
                If buildings Then
                    result = Projektpath & "\05_TF\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_TF_" & PCol(3)(1)(1) & ".xml"
                Else
                    result = Projektpath & "\05_TF\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_TF_" & PCol(3)(1)(1) & ".xml"
                End If
            Else
                ' --- EP
                If buildings Then
                    result = Projektpath & "\01_EP\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
                Else
                    result = Projektpath & "\01_EP\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
                End If
            End If
        End If
    Else
        ' Prinzip
        result = Projektpath & "\03_PR\" & PCol(6)(2) & "_" & PCol(6)(1) & "\TinPlan_PR_" & PCol(6)(1) & ".xml"
    End If

    getXML = result

    Exit Function

ErrHandler:
    If PCol.Count < 6 Then
        If buildings Then
            result = Projektpath & "\01_EP\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
        Else
            result = Projektpath & "\01_EP\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\TinPlan_EP_" & PCol(3)(1)(1) & ".xml"
        End If
    End If
    getXML = result


End Function

Public Function getDWG(PCol As Collection) As String
    ' get the dwg file path for the genearted PK
    Dim Projektpath          As String
    Dim result               As String


    Projektpath = shPData.range("ADM_ProjektpfadCAD")

    Dim buildings            As Boolean
    buildings = Not (shGeb�ude.range("D1").Value = vbNullString)
    If PCol(1) = 0 Then
        ' Plan
        If PCol(7)(1) = "DE" Then
            ' Detail
            result = Projektpath & "\04_DE\DE_" & PCol(15) & ".dwg"
        Else
            If buildings Then
                result = Projektpath & "\01_EP\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\EP_" & PCol(3)(1)(1) & ".dwg"
            Else
                result = Projektpath & "\01_EP\" & Right(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\EP_" & PCol(3)(1)(1) & ".dwg"
            End If
        End If
    Else
        ' Prinzip
        result = Projektpath & "\03_PR\" & PCol(6)(2) & "_" & PCol(6)(1) & "\PR_" & PCol(6)(1) & ".dwg"
    End If

    getDWG = result

End Function

Public Sub DeleteRow(ByVal ID As String)
    ' l�scht die gegebene zeile(row) im Worksheet (DATA [shstoredata])
    Dim row                  As Double
    row = Application.WorksheetFunction.Match(ID, shStoreData.range("A:A"), False)
    shStoreData.rows(row).EntireRow.Delete

End Sub

Public Function getNewRow() As Long
    ' get the next free row in the store sheet
    getNewRow = shStoreData.range("A1").CurrentRegion.rows.Count + 1

End Function

Public Function getRow(PCol As Collection) As Integer
    ' get the corresponding row from the stored data
    getRow = shStoreData.range("A:A").Find(PCol(11), LookIn:=xlValues).row

End Function

Public Function GetArrLength(a As Variant) As Long
    ' get the length of an 1D array
    If IsEmpty(a) Then
        GetArrLength = 0
    Else
        GetArrLength = UBound(a) - LBound(a) + 1
    End If

End Function

Public Function getNewID(ByVal Typ As IDType) As String
    ' get a new unique ID for a PK
    Dim length               As Integer, ws As Worksheet, Region As range, IDcol As Integer

    Select Case Typ
        Case 0                                   ' Plan
            length = 6
            Set ws = Globals.shStoreData
            Set Region = shStoreData.range("A1").CurrentRegion
            IDcol = 1
        Case 1                                   ' Index
            length = 4
            Set ws = Globals.shIndex
            Set Region = Globals.shIndex.range("A1").CurrentRegion
            IDcol = 1
        Case 2                                   ' Task
        Case 3                                   ' Person
            length = 6
            Set ws = Globals.shAdress
            Set Region = Globals.shAdress.range("ADR_Adressen")
            IDcol = 9
    End Select
    Dim i                    As Integer

    Dim rg                   As range
    Set rg = getRange(Region)
    Dim rows                 As Integer
    Dim r                    As Integer

    rows = rg.rows.Count

    i = 4 * length
    Randomize
newID:
    getNewID = Hex(Int(2 ^ i * Rnd(Rnd)))

    For r = 2 To rows + 1
        ' check if the ID already exists
        If getNewID = ws.Cells(r, IDcol).Value Then GoTo newID
    Next r

End Function

Public Function getList(RangeName As String) As Variant()

    Dim arr()                As Variant
    Dim tmparr()             As Variant

    Dim tmprng               As range
    Globals.SetWBs
    Select Case RangeName
        Case "PRO_Geb�ude"
            Set tmprng = Globals.shGeb�ude.range(RangeName)
            arr() = tmprng.Resize(1, tmprng.Columns.Count)
            tmparr() = RemoveBlanksFromStringArray(arr(), True)
        Case "ADM_Firmen"
            Set tmprng = Globals.shAdress.range(RangeName)
            arr() = tmprng.Resize(tmprng.rows.Count, 1)
            tmparr() = RemoveBlanksFromStringArray(arr())
        Case Else
            Set tmprng = Globals.shPData.range(RangeName)
            arr() = tmprng.Resize(tmprng.rows.Count, 1)
            tmparr() = RemoveBlanksFromStringArray(arr())
    End Select

    getList = tmparr()

End Function

Public Function getRange(Region As range, Optional Off As Integer = 1) As range
    ' Auswahl der aktuell gespeicherten Daten im Worksheet (DATA [shData]) ohne �berschriften
    On Error GoTo ERR

    Dim rng                  As range

    If Region.CurrentRegion.rows.Count = Off Then
        ' move one down outside of headers
        Set rng = Region.CurrentRegion.Offset(Off)
    Else
        ' remove the headers
        Set rng = Region.CurrentRegion.Offset(Off).Resize(Region.CurrentRegion.rows.Count - Off)
    End If

    Set getRange = rng

    Exit Function

ERR:
    Set rng = Nothing

End Function

Public Function getUserName() As String

    Dim arrUsername()        As String
    Dim UserName             As String

    UserName = Application.UserName

    'arrUsername = Split(UserName, " ")

    'getUserName = Left(arrUsername(1), 2) & Left(arrUsername(0), 2)

    getUserName = UserName

End Function

Function IsInArray(stringToBeFound As String, arr() As String) As Boolean
    Dim i                    As Variant
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If arr(i) = stringToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next i
    End If
    IsInArray = False
End Function

Public Function RemoveBlanksFromStringArray(ByRef inputArray() As Variant, Optional cols As Boolean = False) As Variant()

    Dim base                 As Long
    base = LBound(inputArray)

    Dim result()             As Variant


    Dim countOfNonBlanks     As Long
    Dim i                    As Long
    Dim myElement            As Variant

    If cols Then
        ReDim result(base To UBound(inputArray, 2))
        For i = base To UBound(inputArray, 2)
            myElement = inputArray(1, i)
            If Not (myElement = vbNullString Or myElement = "-") Then
                result(base + countOfNonBlanks) = myElement
                countOfNonBlanks = countOfNonBlanks + 1
            End If
        Next i
        If countOfNonBlanks = 0 Then
            ReDim result(base To base)
        Else
            ReDim Preserve result(base To base + countOfNonBlanks - 1)
        End If
    Else
        ReDim result(base To UBound(inputArray))
        For i = base To UBound(inputArray)
            myElement = inputArray(i, 1)
            If Not (myElement = vbNullString Or myElement = "-") Then
                result(base + countOfNonBlanks) = myElement
                countOfNonBlanks = countOfNonBlanks + 1
            End If
        Next i
        If countOfNonBlanks = 0 Then
            ReDim result(base To base)
        Else
            ReDim Preserve result(base To base + countOfNonBlanks - 1)
        End If
    End If

    RemoveBlanksFromStringArray = result

End Function


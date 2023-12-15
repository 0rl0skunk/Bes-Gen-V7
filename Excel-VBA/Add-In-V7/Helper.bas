Attribute VB_Name = "Helper"
Option Explicit

'@ModuleDescription "Beinhaltet nützliche Funktionen welche nicht einem Modul zugeordnet werden können."
Public Function GetPlanartNamedRange(Planart As String, Hauptgewerk As String) As String
    ' Gibt die Range der verschiedenen Planarten des aktuellen Hauptgewerkes zurück
    Dim result               As String

    Select Case Hauptgewerk
        Case "Elektro"
            result = "ELE_Planart"
        Case "Gewerbliche Kälte"
            result = "GWK_Planart"
        Case "Koordination"
            result = "KOO_Planart"
        Case "Heizung Kälte"
            result = "HKA_Planart"
        Case "Kälte"
            result = "KAE_Planart"
        Case "Lüftung"
            result = "LUE_Planart"
        Case "Gebäudeautomation"
            result = "GAM_Planart"
        Case "Sanitär"
            result = "SAN_Planart"
        Case "Sprinkler"
            result = "SPR_Planart"
        Case "HLKS/GA Allgemein"
            result = "XXX_Planart"
        Case "Türfachplanung"
            result = "TUE_Planart"
        Case "Brandschutzplanung"
            result = "BRA_Planart"
    End Select

    GetPlanartNamedRange = result

End Function

Public Function GetUnterGewerkKF(UnterGewerk As String, Hauptgewerk As String, Planart As String) As String
    ' Gibt die Kurzform des Untergewerke zurück
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
        Case "Gewerbliche Kälte"
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
        Case "Heizung Kälte"
            Select Case Planart
                Case "Plan"
                    result = "HKA" & "_PLA"
                Case "Schema"
                    result = "HKA" & "_SCH"
                Case "Prinzip"
                    result = "HKA" & "_PRI"
            End Select
        Case "Kälte"
            Select Case Planart
                Case "Plan"
                    result = "KAE" & "_PLA"
                Case "Schema"
                    result = "KAE" & "_SCH"
                Case "Prinzip"
                    result = "KAE" & "_PRI"
            End Select
        Case "Lüftung"
            Select Case Planart
                Case "Plan"
                    result = "LUE" & "_PLA"
                Case "Schema"
                    result = "LUE" & "_SCH"
                Case "Prinzip"
                    result = "LUE" & "_PRI"
            End Select
        Case "Gebäudeautomation"
            Select Case Planart
                Case "Plan"
                    result = "GAM" & "_PLA"
                Case "Schema"
                    result = "GAM" & "_SCH"
                Case "Prinzip"
                    result = "GAM" & "_PRI"
            End Select
        Case "Sanitär"
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
        Case "Türfachplanung"
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

Public Function WLookup(Lookup, range As range, Index As Integer, Optional onError As String = "-") As String
    ' VLookup mit 'onError' wert welcher selbst zugeordnet werden kann.
    On Error GoTo ERR

    Lookup = CStr(Lookup)

    If IsError(Application.VLookup(Lookup, range, Index, False)) Then
        WLookup = onError
    Else
        WLookup = Application.VLookup(Lookup, range, Index, False)
    End If

    Exit Function
    writelog "Info", "Wlookup Value Found " & WLookup

ERR:

    WLookup = onError
    writelog "Error", "Wlookup Value for " & Lookup & " Not Found"

End Function

Public Function getES(ID As String) As String

    Dim Projektpath          As String, result As String
    Dim row                  As Integer

    Projektpath = shPData.range("Projektpfad")

    row = shStoreData.range("A:A").Find(ID).row

    result = Projektpath & "\02_ES\" & shStoreData.Cells(row, 2).Value

    getES = result

End Function

Public Function getFormat(oFormat As String) As String

    Dim tmpstr()             As String
    tmpstr = Split(oFormat, "H")
    Dim breite
    Dim höhe
    If Not oFormat Like "*H*B" Then GoTo exitfunction
    breite = Left(tmpstr(1), Len(tmpstr(1)) - 1)
    höhe = tmpstr(0)
    Select Case Join(Array(breite, höhe), ",")
        Case Join(Array(1, 1), ",")
            getFormat = "A4"
        Case Join(Array(2, 1), ",")
            getFormat = "A3"
        Case Join(Array(2, 2), ",")
            getFormat = "A2"
        Case Join(Array(4, 2), ",")
            getFormat = "A1"
        Case Join(Array(4, 4), ",")
            getFormat = "A0"
        Case Else
            getFormat = höhe * 29.7 & "x" & breite * 21 & "cm"
    End Select

    Exit Function
exitfunction:
    getFormat = "---"

    Exit Function
End Function

Public Function deleteIndexesXml()

    Dim fso                  As Object

    Set fso = CreateObject("scripting.FileSystemObject")
    Dim lastrow              As Integer, row As Integer, col As Integer, lastcol As Integer

    lastrow = shGebäude.Cells(rows.Count, 2).End(xlUp).row
    lastcol = shGebäude.Cells(1, Columns.Count).End(xlToLeft).Column

    For col = 2 To lastcol Step 2
        For row = 6 To lastrow
            writelog "Info", "> empty cell " & row & " " & col & " " & IsEmpty(shGebäude.Cells(row, col)) & " " & shGebäude.Cells(row, col).Address
            If Not IsEmpty(shGebäude.Cells(row, col)) Then
                writelog "Info", "> " & shGebäude.Cells(row, col).Value
                ' only go if there is something in the cell
                deleteIndexXml row, col
            End If
        Next row
    Next col

    ' ----------------------------- PRINZIPSCHEMAS
    Dim j                    As Integer, i As Integer, filename As String
    Dim TinLine, Projektname, Projektpfad, EP, Gewerk As String, GewerkNr As String
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

End Function

Function deleteIndexXml(row As Integer, col As Integer, Optional i_xmlfile As String = "")

    Dim xmlfile              As String
    If i_xmlfile <> "" Then
        xmlfile = i_xmlfile
    Else
        xmlfile = i_xmlfile
    End If

load:                                            ' load xml file
    Dim oXml                 As MSXML2.DOMDocument60
    Set oXml = New MSXML2.DOMDocument60
    oXml.load xmlfile
    Dim nodes                As IXMLDOMNodeList
    Dim node                 As IXMLDOMNode
    Dim root                 As IXMLDOMNode

    Set root = oXml.SelectSingleNode("//tinPlan1")
    Set nodes = oXml.SelectNodes("//tinPlan1/*[contains(local-name(), 'IN')]")
    For Each node In nodes
        root.RemoveChild node
    Next
    On Error Resume Next
    oXml.save xmlfile
    Set oXml = Nothing
    On Error GoTo 0

End Function

Public Function getXML(PCol As Collection) As String
    ' get the xml file path for the genearted PK

    Dim Projektpath          As String, result As String
    On Error GoTo ErrHandler
    Projektpath = shPData.range("ADM_ProjektpfadCAD")
    'On Error Resume Next
    Dim buildings            As Boolean
    If shGebäude.range("D1").Value = "" Then
        buildings = False
    Else
        buildings = True
    End If
    If PCol(1) = 0 Then
        ' Plan
        If PCol.Count > 7 Then
            If PCol(7)(1) = "DE" Then
                ' Detail
                result = Projektpath & "\04_DE\TinPlan_DE_" & PCol(15) & ".xml"
            Else
                GoTo plan
            End If
        Else
plan:
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
    Dim Projektpath          As String, result As String

    Projektpath = shPData.range("ADM_ProjektpfadCAD")

    Dim buildings            As Boolean
    If shGebäude.range("D1").Value = "" Then
        buildings = False
    Else
        buildings = True
    End If
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
    ' löscht die gegebene zeile(row) im Worksheet (DATA [shstoredata])
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

Public Function CreateXmlAttribute(Name As String, Bez As String, Wert As String, str As String, nodChild As IXMLDOMElement, oXml As MSXML2.DOMDocument60, nodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim nodGrandChild        As IXMLDOMElement

    Set nodChild = oXml.createElement(str)
    nodElement.appendChild nodChild

    Set nodGrandChild = oXml.createElement("Name")
    nodGrandChild.Text = Name
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Bez")
    nodGrandChild.Text = Bez
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Wert")
    nodGrandChild.Text = Wert
    nodChild.appendChild nodGrandChild

    CreateXmlAttribute = True

End Function

Public Function CreateXmlIndexAttribute(Index As String, Name As String, Datum As String, Bez As String, NodName As String, nodChild As IXMLDOMElement, oXml As MSXML2.DOMDocument60, nodElement As IXMLDOMElement)
    ' create a TinLine XML Attribute with the given informations
    Dim nodGrandChild        As IXMLDOMElement

    Set nodChild = oXml.createElement(NodName)
    nodElement.appendChild nodChild

    Set nodGrandChild = oXml.createElement("Index")
    nodGrandChild.Text = Index
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Name")
    nodGrandChild.Text = Name
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Datum")
    nodGrandChild.Text = Datum
    nodChild.appendChild nodGrandChild

    Set nodGrandChild = oXml.createElement("Bez")
    nodGrandChild.Text = Bez
    nodChild.appendChild nodGrandChild

    CreateXmlIndexAttribute = True

End Function

Public Function GetArrLength(a As Variant) As Long
    ' get the length of an 1D array
    If IsEmpty(a) Then
        GetArrLength = 0
    Else
        GetArrLength = UBound(a) - LBound(a) + 1
    End If

End Function

Public Function getNewID(length As Integer, ws As Worksheet, Region As range, IDcol As Integer) As String
    ' get a new unique ID for a PK
    Dim i                    As Integer

    Dim rg                   As range
    Set rg = getRange(Region)
    Dim rows                 As Integer, r As Integer
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

    Dim arr(), tmparr()
    Dim tmprng               As range


    If RangeName = "PRO_Gebäude" Then
        Set tmprng = Globals.shGebäude.range(RangeName)
        arr() = tmprng.Resize(1, tmprng.Columns.Count)
        tmparr() = RemoveBlanksFromStringArray(arr(), True)
    Else
        Set tmprng = Globals.shPData.range(RangeName)
        arr() = tmprng.Resize(tmprng.rows.Count, 1)
        tmparr() = RemoveBlanksFromStringArray(arr())
    End If

    getList = tmparr()

End Function

Public Function getRange(Region As range, Optional Off As Integer = 1) As range
    ' Auswahl der aktuell gespeicherten Daten im Worksheet (DATA [shData]) ohne überschriften
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

    Dim arrUsername()        As String, UserName As String
    UserName = Application.UserName

    'arrUsername = Split(UserName, " ")

    'getUserName = Left(arrUsername(1), 2) & Left(arrUsername(0), 2)
    
    getUserName = UserName

End Function

Function IsInArray(stringToBeFound As String, arr() As String) As Boolean
    Dim i
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
    Dim myElement

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


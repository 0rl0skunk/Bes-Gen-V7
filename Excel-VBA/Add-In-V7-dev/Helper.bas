Attribute VB_Name = "Helper"
Attribute VB_Description = "Beinhaltet nützliche Funktionen welche nicht einem Modul zugeordnet werden können."

'@IgnoreModule VariableNotUsed
'@ModuleDescription "Beinhaltet nützliche Funktionen welche nicht einem Modul zugeordnet werden können."

Option Explicit

Public Enum IDType
    IDPlankopf = 0
    IDIndex = 1
    IDTask = 2
    IDPerson = 3
End Enum

Public Function GetPlanartNamedRange(ByVal Planart As String, ByVal Hauptgewerk As String) As String
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

Public Function GetUnterGewerkKF(UnterGewerk As String, ByVal Hauptgewerk As String, ByVal Planart As String) As String
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

Public Function CollectionToArray(ByVal myCol As Collection) As Variant
    ' convert a collection of elements to an array
    Dim result               As Variant
    Dim cnt                  As Long

    If myCol.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If

    ReDim result(myCol.Count - 1)
    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt
    CollectionToArray = result

End Function

Public Function WLookup(Lookup As Variant, range As range, Index As Long, Optional ByVal onError As String = "-") As String
    ' VLookup mit 'onError' wert welcher selbst zugeordnet werden kann.
    On Error GoTo err

    Lookup = CStr(Lookup)

    If IsError(Application.VLookup(Lookup, range, Index, False)) Then
        WLookup = onError
    Else
        WLookup = Application.VLookup(Lookup, range, Index, False)
    End If

    Exit Function
    writelog LogInfo, "Wlookup Value Found " & WLookup

err:

    WLookup = onError
    writelog LogError, "Wlookup Value for " & Lookup & " Not Found"

End Function

Public Function getDWG(ByVal PCol As Collection) As String
    ' get the dwg file path for the genearted PK
    Dim Projektpath          As String
    Dim result               As String


    Projektpath = shPData.range("ADM_ProjektpfadCAD")

    Dim buildings            As Boolean
    buildings = Not (shGebäude.range("D1").value = vbNullString)
    If PCol(1) = 0 Then
        ' Plan
        If PCol(7)(1) = "DE" Then
            ' Detail
            result = Projektpath & "\04_DE\DE_" & PCol(15) & ".dwg"
        Else
            If buildings Then
                result = Projektpath & "\01_EP\" & PCol(3)(0)(2) & "_" & PCol(3)(0)(1) & "\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\EP_" & PCol(3)(1)(1) & ".dwg"
            Else
                result = Projektpath & "\01_EP\" & Right$(PCol(3)(1)(2), 2) & "_" & PCol(3)(1)(1) & "\EP_" & PCol(3)(1)(1) & ".dwg"
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

Public Function getRow(ByVal PCol As Collection) As Long
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
    Dim length               As Long
    Dim ws                   As Worksheet
    Dim Region               As range
    Dim IDcol                As Long

    Select Case Typ
    Case 0                                       ' Plan
        length = 6
        Set ws = Globals.shStoreData
        Set Region = shStoreData.range("A1").CurrentRegion
        IDcol = 1
    Case 1                                       ' Index
        length = 4
        Set ws = Globals.shIndex
        Set Region = Globals.shIndex.range("A1").CurrentRegion
        IDcol = 1
    Case 2                                       ' Task
    Case 3                                       ' Person
        length = 6
        Set ws = Globals.shAdress
        Set Region = Globals.shAdress.range("ADR_Adressen")
        IDcol = 9
    End Select
    Dim i                    As Long

    Dim rg                   As range
    Set rg = getRange(Region)
    Dim rows                 As Long
    Dim r                    As Long

    rows = rg.rows.Count

    i = 4 * length
    Randomize
newID:
    getNewID = Hex$(Int(2 ^ i * Rnd(Rnd)))

    For r = 2 To rows + 1
        ' check if the ID already exists
        If getNewID = ws.Cells(r, IDcol).value Then GoTo newID
    Next r

End Function

Public Function getList(ByVal RangeName As String) As Variant()

    Dim arr()                As Variant
    Dim tmparr()             As Variant

    Dim tmprng               As range
    Globals.SetWBs
    Select Case RangeName
    Case "PRO_Gebäude"
        Set tmprng = Globals.shGebäude.range(RangeName)
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

Public Function getRange(ByVal Region As range, Optional ByVal Off As Long = 1) As range
    ' Auswahl der aktuell gespeicherten Daten im Worksheet (DATA [shData]) ohne überschriften
    On Error GoTo err

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

err:
    Set rng = Nothing

End Function

Public Function getUserName() As String

    Dim arrUsername()        As String
    Dim UserName             As String
    
    On Error GoTo ErrHandler
    
    UserName = Application.UserName
    
    arrUsername = Split(UserName, " ")
    getUserName = Left(arrUsername(1), 2) & Left(arrUsername(0), 2)
    Exit Function
    
ErrHandler:
    getUserName = UserName

End Function

Function ArrayIndex(ByVal arr As Variant, ByVal value As Variant) As Long
    Dim i                    As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If arr(i) = value Then ArrayIndex = i: Exit Function
        Next i
    End If
    ArrayIndex = -1
End Function

Function IsInArray(ByVal stringToBeFound As String, arr() As String) As Boolean
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

Public Function RemoveBlanksFromStringArray(ByRef inputArray() As Variant, Optional ByVal cols As Boolean = False) As Variant()

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

Function CountFiles(ByVal path As String) As Long

    Dim fso                  As Object
    Dim Folder               As Object
    Dim subfolder            As Object
    Dim amount               As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set Folder = fso.GetFolder(path)
    For Each subfolder In Folder.SubFolders
        amount = amount + CountFiles(subfolder.path)
    Next subfolder

    amount = amount + Folder.files.Count

    Set fso = Nothing
    Set Folder = Nothing
    Set subfolder = Nothing

    CountFiles = amount

End Function

Function SplitStringByLength(ByVal inputString As String, ByVal maxLength As Long) As Variant
    Dim inputArray()         As String
    Dim outputArray()        As String
    Dim currentLength        As Long
    Dim currentLine          As String
    Dim wordArray()          As String
    Dim word                 As Variant
    Dim i                    As Long

    inputArray = Split(inputString, " ")
    currentLength = 0
    currentLine = vbNullString

    ReDim outputArray(0)
    outputArray(0) = vbNullString

    For Each word In inputArray
        wordArray = Split(word, vbLf)
        For i = LBound(wordArray) To UBound(wordArray)
            If currentLength + Len(wordArray(i)) + 1 <= maxLength Then
                currentLine = currentLine & " " & wordArray(i)
                currentLength = currentLength + Len(wordArray(i)) + 1
            Else
                ReDim Preserve outputArray(UBound(outputArray) + 1)
                outputArray(UBound(outputArray)) = Trim$(currentLine)
                currentLine = wordArray(i)
                currentLength = Len(wordArray(i))
            End If
        Next i
    Next word

    If Len(Trim$(currentLine)) > 0 Then
        ReDim Preserve outputArray(UBound(outputArray) + 1)
        outputArray(UBound(outputArray)) = Trim$(currentLine)
    End If

    SplitStringByLength = outputArray
End Function


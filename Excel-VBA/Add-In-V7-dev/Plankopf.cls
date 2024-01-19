VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Plankopf Objekt f�r einfacheres Management von Plank�pfen."

'@Folder "Plankopf"
'@ModuleDescription "Plankopf Objekt f�r einfacheres Management von Plank�pfen."

Option Explicit

Private Const ErrSourceModule = "Plankopf.cls"

Implements IPlankopf

Private Type TPlankopf
    '--- Identifiers ---
    ID As String
    IDTinLine As String: TinLinePKNr As String
    Planart As String: PlanartKF As String
    PLANTYP As String
    TinLineTyp As String
    '--- Zeichner ---
    GezeichnetPerson As String
    GezeichnetDatum As String
    Gepr�ftPerson As String
    Gepr�ftDatum As String
    '--- Geb�ude ---
    Geb�ude As String: Geb�udeKF As String: Geb�udeCode As String
    Geb�udeteil As String: Geb�udeTeilKF As String
    Geschoss As String: GeschossKF As String: GeschossCode As String
    '--- Gewerk ---
    Gewerk As String: GewerkKF As String
    UnterGewerk As String: UnterGewerkKF As String
    '--- Planbezeichnung ---
    Plan�berschrift As String
    PlanBeschrieb As String
    CustomPlan�berschrift As Boolean
    '--- Layout ---
    LayoutName As String
    LayoutGr�sse As String
    LayoutGr�ssecm As String
    LayoutMasstab As String
    LayoutPlanstand As String
    '--- File Path ---
    Pfad As String
    PDFFileName As String
    DWGFileName As String
    DWGFilePath As String
    XMLFileName As String
    XMLFilePath As String
    '--- Index ---
    CurrentIndex As IIndex
    Indexes As New Collection
    SkipValidation As Boolean
    '--- Schema / Verteilung ---
    AnlageTyp As String: AnlageTypKF As String
    AnlageNummer As String
End Type

Private This                 As TPlankopf
Private Inputs               As TPlankopf
Private pProjekt             As IProjekt

Private Property Get IPlankopf_Gewerk() As String
    IPlankopf_Gewerk = This.Gewerk
End Property

Private Property Get IPlankopf_UnterGewerk() As String
    IPlankopf_UnterGewerk = This.UnterGewerk
End Property

Private Property Get IPlankopf_UnterGewerkKF() As String
    IPlankopf_UnterGewerkKF = This.UnterGewerkKF
End Property

Private Property Get IPlankopf_ID() As String
    IPlankopf_ID = This.ID
End Property

Private Property Get IPlankopf_TinLinePKNr() As String
    IPlankopf_TinLinePKNr = This.TinLinePKNr
End Property

Private Property Let IPlankopf_TinLinePKNr(ByVal NewPKNr As String)
    This.TinLinePKNr = NewPKNr
End Property

Private Property Let IPlankopf_IDTinLine(ByVal newID As String)
    This.IDTinLine = newID
End Property

Private Property Let IPlankopf_ID(ByVal NewValue As String)
    This.ID = NewValue
End Property

Private Property Get IPlankopf_IDTinLine() As String
    IPlankopf_IDTinLine = This.IDTinLine
End Property

Private Property Get IPlankopf_Plantyp() As String
    IPlankopf_Plantyp = This.PLANTYP
End Property

Private Property Get IPlankopf_Planart() As String
    IPlankopf_Planart = This.Planart
End Property

Private Property Get IPlankopf_GezeichnetPerson() As String
    IPlankopf_GezeichnetPerson = This.GezeichnetPerson
End Property

Private Property Get IPlankopf_GezeichnetDatum() As String
    IPlankopf_GezeichnetDatum = This.GezeichnetDatum
End Property

Private Property Get IPlankopf_Gezeichnet() As String
    IPlankopf_Gezeichnet = This.GezeichnetPerson & " ; " & This.GezeichnetDatum
End Property

Private Property Get IPlankopf_Gepr�ftPerson() As String
    IPlankopf_Gepr�ftPerson = This.Gepr�ftPerson
End Property

Private Property Get IPlankopf_Gepr�ftDatum() As String
    IPlankopf_Gepr�ftDatum = This.Gepr�ftDatum
End Property

Private Property Get IPlankopf_Gepr�ft() As String
    IPlankopf_Gepr�ft = This.Gepr�ftPerson & " ; " & This.Gepr�ftDatum
End Property

Private Property Get IPlankopf_Geb�ude() As String
    IPlankopf_Geb�ude = This.Geb�ude
End Property

Private Property Get IPlankopf_Geb�udeteil() As String
    IPlankopf_Geb�udeteil = This.Geb�udeteil
End Property

Private Property Get IPlankopf_Geschoss() As String
    IPlankopf_Geschoss = This.Geschoss
End Property

Private Property Get IPlankopf_Plan�berschrift() As String
    IPlankopf_Plan�berschrift = This.Plan�berschrift
End Property

Private Property Get IPlankopf_PlanBeschrieb() As String
    IPlankopf_PlanBeschrieb = This.PlanBeschrieb
End Property

Private Property Get IPlankopf_CustomPlan�berschrift() As Boolean
    IPlankopf_CustomPlan�berschrift = This.CustomPlan�berschrift
End Property

Private Property Get IPlankopf_LayoutName() As String
    IPlankopf_LayoutName = This.LayoutName
End Property

Private Property Get IPlankopf_PlanNummer() As String
    IPlankopf_PlanNummer = This.LayoutName
End Property

Private Property Get IPlankopf_LayoutGr�sse(Optional ByVal Layout As Boolean = False) As String
    If Layout Then
        IPlankopf_LayoutGr�sse = This.LayoutGr�ssecm
    Else
        IPlankopf_LayoutGr�sse = This.LayoutGr�sse
    End If
End Property

Private Property Get IPlankopf_LayoutMasstab() As String
    IPlankopf_LayoutMasstab = This.LayoutMasstab
End Property

Private Property Get IPlankopf_LayoutPlanstand() As String
    IPlankopf_LayoutPlanstand = This.LayoutPlanstand
End Property

Private Property Get IPlankopf_PDFFileName() As String
    IPlankopf_PDFFileName = This.PDFFileName
End Property

Private Property Get IPlankopf_DWGFileName() As String
    IPlankopf_DWGFileName = This.DWGFileName
End Property

Private Property Get IPlankopf_DWGFile() As String
    IPlankopf_DWGFile = This.DWGFilePath & "\" & This.DWGFileName & ".dwg"
End Property

Private Property Get IPlankopf_XMLFileName() As String
    IPlankopf_XMLFileName = This.XMLFileName
End Property

Private Property Get IPlankopf_XMLFile() As String
    IPlankopf_XMLFile = This.XMLFilePath & "\" & This.XMLFileName
End Property

Private Property Get IPlankopf_FolderName() As String
    IPlankopf_FolderName = This.DWGFilePath
End Property

Private Property Get IPlankopf_CurrentIndex() As IIndex
    If This.CurrentIndex Is Nothing Then Set This.CurrentIndex = IndexFactory.Create(This.ID, vbNullString, vbNullString, vbNullString, vbNullString, " ", vbNullString, vbNullString, True)
    Set IPlankopf_CurrentIndex = This.CurrentIndex
End Property

Private Property Get IPlankopf_Indexes() As Collection
    Set IPlankopf_Indexes = This.Indexes
End Property

Private Property Set IPlankopf_Indexes(ByVal NewIndexes As Collection)
    Set This.Indexes = NewIndexes
End Property

Private Property Get IPlankopf_Anlagetyp() As String
    IPlankopf_Anlagetyp = This.AnlageTyp
End Property

Private Property Get IPlankopf_AnlageNummer() As String
    IPlankopf_AnlageNummer = This.AnlageNummer
End Property

' --- Functions ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Filldata( _
       ByVal Projekt As IProjekt, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Gepr�ftPerson As String, _
       ByVal Gepr�ftDatum As String, _
       ByVal Geb�ude As String, _
       ByVal Geb�udeteil As String, _
       ByVal Geschoss As String, _
       ByVal Gewerk As String, _
       ByVal UnterGewerk As String, _
       ByVal Format As String, _
       ByVal Masstab As String, _
       ByVal Stand As String, _
       ByVal Planart As String, _
       Optional ByVal PLANTYP As String, _
       Optional ByVal TinLineID As String = "NEW", _
       Optional ByVal SkipValidation As Boolean = False, _
       Optional ByVal Plan�berschrift As String = "NEW", _
       Optional ByVal ID As String = "ID", _
       Optional ByVal Custom�berschrift As Boolean = False, _
       Optional ByVal AnlageTyp As String, _
       Optional ByVal AnlageNummer As String _
       ) As Boolean
    ' SkipValidation ist zu debug zwecken zu gebrauchen
    With Inputs
        .IDTinLine = TinLineID
        .GezeichnetPerson = GezeichnetPerson
        .GezeichnetDatum = GezeichnetDatum
        .Gepr�ftPerson = Gepr�ftPerson
        .Gepr�ftDatum = Gepr�ftDatum
        .Geb�ude = Geb�ude
        .Geb�udeteil = Geb�udeteil
        .Gewerk = Gewerk
        .UnterGewerk = UnterGewerk
        .Geschoss = Geschoss
        .Plan�berschrift = Plan�berschrift
        .LayoutGr�sse = Format
        .LayoutMasstab = Masstab
        .LayoutPlanstand = Stand
        .Planart = Planart
        .PLANTYP = PLANTYP
        .CustomPlan�berschrift = Custom�berschrift
        .SkipValidation = SkipValidation
        .AnlageNummer = AnlageNummer
        .AnlageTyp = AnlageTyp
    End With
    Set pProjekt = Projekt
    'On Error GoTo NotValid
    If ValidInputs Then
        This = Inputs
        If ID = "ID" Then
            This.ID = getNewID(IDPlankopf)
        Else
            This.ID = ID
        End If

        GetKF
        GetFileNames
        GetPlan�berschrift
        GetFormat
        Refresh
    End If
    On Error GoTo 0
    Filldata = True
    Exit Function

    On Error GoTo -1
    On Error GoTo 0
    Dim frm                  As New UserFormMessage
    frm.Typ TypWarning, "Es wurde kein Plankopf erstellt." & vbNewLine & _
                       "�ffne den Log um die Fehlenden Inhalte zu sehen.", _
            "Kein Plankopf erstellt", True
    frm.Show
    Filldata = False
    Exit Function

End Function

Private Sub GetFormat()
    Dim tmpstr()             As String
    If Len(This.LayoutGr�sse) < 4 Then Exit Sub
    tmpstr = Split(This.LayoutGr�sse, "H")
    Dim breite               As Long
    Dim h�he                 As Long
    breite = Left$(tmpstr(1), Len(tmpstr(1)) - 1)
    h�he = tmpstr(0)
    Select Case Join(Array(breite, h�he), ",")
    Case Join(Array(1, 1), ",")
        This.LayoutGr�ssecm = "A4"
    Case Join(Array(2, 1), ",")
        This.LayoutGr�ssecm = "A3"
    Case Join(Array(2, 2), ",")
        This.LayoutGr�ssecm = "A2"
    Case Join(Array(4, 2), ",")
        This.LayoutGr�ssecm = "A1"
    Case Join(Array(4, 4), ",")
        This.LayoutGr�ssecm = "A0"
    Case Else
        This.LayoutGr�ssecm = h�he * 29.7 & "x" & breite * 21 & "cm"
    End Select
End Sub

Private Sub GetKF()
    Dim ErrSource As String
    Dim ErrDescription As String
    Dim ErrSourceLog As String
    ErrSource = ErrSourceModule & " > " & "GetKF"
    On Error GoTo ErrHandler
    This.Geb�udeKF = Globals.shGeb�ude.range("1:1").Find(This.Geb�ude).Offset(1).value
    If This.Geb�udeKF = "GES" And Globals.shGeb�ude.range("D1").value = vbNullString Then
        This.Geb�udeKF = vbNullString
    End If
    If This.Geb�udeKF = vbNullString Then
        This.Geb�udeCode = vbNullString
    Else
        This.Geb�udeCode = Globals.shGeb�ude.range("1:1").Find(This.Geb�ude).Offset(2).value
    End If
    ErrSourceLog = ErrSource & " > " & "Geschoss-Code"
    ErrDescription = This.Geschoss & " nicht in Range Globals.shGeb�ude.Range(B:AQ) gefunden"
1   This.GeschossCode = Globals.shGeb�ude.Cells(Globals.shGeb�ude.range("B:AQ").Find(This.Geschoss).row, 1).value
    ErrSourceLog = ErrSource & " > " & "Geschoss KF"
    ErrDescription = This.Geschoss & " nicht in Range Globals.shGeb�ude.Range(B:AQ) gefunden"
2   This.GeschossKF = Globals.shGeb�ude.range("B:AQ").Find(This.Geschoss).Offset(, 1).value
    ErrSourceLog = ErrSource & " > " & "Geb�udeteil KF"
    ErrDescription = This.Geb�udeteil & " nicht in Range Globals.shPData.range(PRO_Geb�udeteil) gefunden"
3   This.Geb�udeTeilKF = Globals.shPData.range("PRO_Geb�udeteil").Find(This.Geb�udeteil).Offset(0, 1).value
    ErrSourceLog = ErrSource & " > " & "Gewerk KF"
    ErrDescription = This.Gewerk & " nicht in Range Globals.shPData.range(PRO_Hauptgewerk) gefunden"
4   This.GewerkKF = Globals.shPData.range("PRO_Hauptgewerk").Find(This.Gewerk).Offset(0, 1).value

    Dim col                  As Long
    Select Case This.PLANTYP
    Case "PLA"
        ErrSourceLog = ErrSource & " > " & "PLA > Column UnterGewerk KF"
        ErrDescription = This.GewerkKF & " nicht in Range Globals.shPData.range(10:10) gefunden"
5       col = Application.Match(This.GewerkKF & " PLA", Globals.shPData.range("10:10"), 0)
        ErrSourceLog = ErrSource & " > " & "PLA > UnterGewerk KF"
        ErrDescription = This.UnterGewerk & " nicht in Range Globals.shPData.Columns(" & col & ") gefunden"
6       This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).value
        ErrSourceLog = ErrSource & " > " & "PLA > Column Planart KF"
        ErrDescription = This.GewerkKF & " nicht in Range Globals.shPData.range(9:9) gefunden"
7       col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
        ErrSourceLog = ErrSource & " > " & "PLA > Planart KF"
        ErrDescription = This.Planart & " nicht in Range Globals.shPData.Columns(" & col & ") gefunden"
8       This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).value
        ' Welcher K�rzel kommt vor die Stockwerkbezeichnung beim DWG?
        Select Case This.GewerkKF
        Case "ELE"
            This.TinLineTyp = "EP"
        Case "BRA"
            This.TinLineTyp = "BS"
        Case "TUE"
            This.TinLineTyp = "TF"
        End Select
    Case "SCH"
        ErrSourceLog = ErrSource & " > " & "SCH > Column UnterGewerk KF"
        ErrDescription = This.GewerkKF & " nicht in Range Globals.shPData.range(10:10) gefunden"
9       col = Application.Match(This.GewerkKF & " SCH", Globals.shPData.range("10:10"), 0)
        ErrSourceLog = ErrSource & " > " & "SCH > UnterGewerk KF"
        ErrDescription = This.UnterGewerk & " nicht in Range Globals.shPData.Columns(" & col & ") gefunden"
10      This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).value
        This.TinLineTyp = vbNullString
        ' Kurzform f�r Anlagetyp
        ErrSourceLog = ErrSource & " > " & "SCH > Column Planart KF"
        ErrDescription = This.GewerkKF & " nicht in Range Globals.shPData.range(12:12) gefunden"
11      col = Application.Match("Anlagetyp " & This.UnterGewerk, Globals.shPData.range("12:12"), 0)
        ErrSourceLog = ErrSource & " > " & "SCH > Anlagetyp KF"
        ErrDescription = This.AnlageTyp & " nicht in Range Globals.shPData.Columns(" & col & ") gefunden"
12      This.AnlageTypKF = Globals.shPData.Columns(col).Find(This.AnlageTyp).Offset(0, 1).value
    Case "PRI"
        ErrSourceLog = ErrSource & " > " & "PRI > Column UnterGewerk KF"
        ErrDescription = This.GewerkKF & " nicht in Range Globals.shPData.range(10:10) gefunden"
13      col = Application.Match(This.GewerkKF & " PRI", Globals.shPData.range("10:10"), 0)
        ErrSourceLog = ErrSource & " > " & "PRI > UnterGewerk KF"
        ErrDescription = This.UnterGewerk & " nicht in Range Globals.shPData.Columns(" & col & ") gefunden"
14      This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).value
        ErrSourceLog = ErrSource & " > " & "PRI > Column Planart KF"
        ErrDescription = This.GewerkKF & " nicht in Range Globals.shPData.range(9:9) gefunden"
15      col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
        This.PlanartKF = "PR"
        This.TinLineTyp = "PR"
    End Select

    writelog LogInfo, "Kurzformen erstellt"

    Exit Sub
ErrHandler:
    writelog LogError, "Beim erstellen der Kurzformen" & vbNewLine & "PlankopfID: " & This.ID & vbNewLine & _
                      "Error: " & err.Number & vbNewLine & ErrDescription & vbNewLine & ErrSourceLog
    err.Clear

End Sub

Private Function GetFileNames()
    'BUG Hier kommen noch nicht die korrekten Pf�de raus.
    ' --- DWG FileName
    Select Case This.PLANTYP
    Case "PLA"
        This.DWGFileName = This.TinLineTyp & "_" & This.GeschossKF
    Case "SCH"
        This.DWGFileName = "1"
    Case "PRI"
        This.DWGFileName = This.TinLineTyp & "_" & This.UnterGewerkKF
    End Select
    ' --- XML FileName
    Select Case This.PLANTYP
    Case "PLA"
        This.XMLFileName = "TinPlan_" & This.TinLineTyp & "_" & This.GeschossKF & ".xml"
    Case "SCH"
        This.XMLFileName = "1.xml"
    Case "PRI"
        This.XMLFileName = "TinPlan_" & This.TinLineTyp & "_" & This.UnterGewerkKF & ".xml"
    End Select
    ' --- Pfad
    Select Case This.PLANTYP
    Case "PLA"
        If This.Geb�ude = "Gesamt" And Globals.shGeb�ude.range("D1").value = vbNullString Then
            Select Case This.TinLineTyp
            Case "EP"
                This.Pfad = pProjekt.ProjektOrdnerCAD & "\01_EP\" & This.GeschossCode & "_" & This.GeschossKF
            Case "BS"
                This.Pfad = pProjekt.ProjektOrdnerCAD & "\06_BS\" & This.GeschossCode & "_" & This.GeschossKF
            Case "TF"
                This.Pfad = pProjekt.ProjektOrdnerCAD & "\05_TF\" & This.GeschossCode & "_" & This.GeschossKF
            End Select
        Else                                     ' Mehrere Geb�ude
            Select Case This.TinLineTyp
            Case "EP"
                This.Pfad = pProjekt.ProjektOrdnerCAD & "\01_EP\" & This.Geb�udeCode & "_" & This.Geb�udeKF & "\" & This.GeschossCode & "_" & This.GeschossKF
            Case "BS"
                This.Pfad = pProjekt.ProjektOrdnerCAD & "\06_BS\" & This.Geb�udeCode & "_" & This.Geb�udeKF & "\" & This.GeschossCode & "_" & This.GeschossKF
            Case "TF"
                This.Pfad = pProjekt.ProjektOrdnerCAD & "\05_TF\" & This.Geb�udeCode & "_" & This.Geb�udeKF & "\" & This.GeschossCode & "_" & This.GeschossKF
            End Select
        End If
        This.DWGFilePath = This.Pfad
        This.XMLFilePath = This.DWGFilePath
    Case "SCH"
        This.Pfad = pProjekt.ProjektOrdnerCAD & "\02_ES" & This.LayoutName
        This.DWGFilePath = This.Pfad
        This.XMLFilePath = This.DWGFilePath
    Case "PRI"
        Dim ind              As Long
        Dim indStr           As String
        ind = ArrayIndex(Array("OBJ", "STS", "SLS", "KOM", "MES", "NOT", "BRA", "EVA", "GAT", "ZUT", "SEC", "ERD", "PRO"), This.UnterGewerkKF)
        indStr = ind
        If Len(indStr) < 2 Then indStr = "0" & indStr
        This.Pfad = pProjekt.ProjektOrdnerCAD & "\03_PR"
        This.DWGFilePath = This.Pfad & "\" & indStr & "_" & This.UnterGewerkKF
        This.XMLFilePath = This.DWGFilePath
    End Select
    ' --- FilePath
End Function

Private Sub GetPlan�berschrift()

    Dim Plan�berschrift      As String
    If This.CustomPlan�berschrift Then Exit Sub
    Dim Geb�udeteil As String
    Dim Geb�udeTeilKF As String
    If This.Geb�udeteil = "Gesamt" Then
        Geb�udeteil = vbNullString
        Geb�udeTeilKF = vbNullString
    Else
        Geb�udeteil = This.Geb�udeteil
        Geb�udeTeilKF = This.Geb�udeTeilKF
    End If
    Select Case This.PLANTYP
    Case "PLA"
        ' �berschrift f�r Elektropl�ne
        Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.Geschoss, Geb�udeteil)
        This.PlanBeschrieb = Plan�berschrift
        If Len(Plan�berschrift) > Globals.maxlen Then
            Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, Geb�udeteil)
            If Len(Plan�berschrift) > maxlen Then
                Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, Geb�udeTeilKF)
                If Len(Plan�berschrift) > maxlen Then
                    Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerkKF, This.GeschossKF, Geb�udeTeilKF)
                    If Len(Plan�berschrift) > maxlen Then
                        Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.PlanartKF, This.UnterGewerkKF, This.GeschossKF, Geb�udeTeilKF)
                    End If
                End If
            End If
        End If
    Case "PRI"
        ' �berschrift f�r Prinzipschemas
        Plan�berschrift = "Prinzipschema " & This.UnterGewerk
        This.PlanBeschrieb = Plan�berschrift
        If Len(Plan�berschrift) > Globals.maxlen Then
            Plan�berschrift = "Prinzipschema " & This.UnterGewerk
            If Len(Plan�berschrift) > maxlen Then
                Plan�berschrift = "PR " & This.UnterGewerk
            End If
        End If

    Case "SCH"
        ' �berschrift f�r Schemas
        Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Gewerk, This.AnlageTyp, This.Geb�ude, This.Geschoss, This.Geb�udeteil, "Anlagenummer:", This.AnlageNummer)
    End Select

    ' wurde die Plan�berschrift bearbeitet oder nicht?
    If This.Plan�berschrift <> Plan�berschrift And This.Plan�berschrift <> vbNullString Then
        Select Case MsgBox("Die Beschriftung: " & This.Plan�berschrift & " ist nicht Standard." & vbNewLine & "Soll diese mit der Standardm�ssigen �berschrift ersetzt werden?" & vbNewLine & Plan�berschrift, vbYesNo, "Beschriftung �berschreiben?")
        Case vbYes
            This.Plan�berschrift = Plan�berschrift
            This.CustomPlan�berschrift = False
            writelog LogInfo, "Plan�berschrift erstellt: " & This.Plan�berschrift
        Case vbNo
            This.CustomPlan�berschrift = True
            writelog LogWarning, "Plan�berschrift nicht erstellt: " & This.Plan�berschrift & " >> alte Beschriftung: " & Plan�berschrift
        End Select
    Else
        This.Plan�berschrift = Plan�berschrift
        writelog LogInfo, "Plan�berschrift erstellt: " & This.Plan�berschrift
    End If

End Sub

Private Function ValidInputs() As Boolean

    If Inputs.SkipValidation Then ValidInputs = True: Exit Function

    ValidInputs = False
    Dim ErrorSource          As String
    ErrorSource = "Plankopf > ValidInputs > "
    If Len(Inputs.GezeichnetPerson) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEZEICHNET PERSON' given"
    If Len(Inputs.GezeichnetDatum) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEZEICHNET DATUM' given"
    If Len(Inputs.Gepr�ftPerson) = 0 Then writelog LogWarning, ErrorSource & "no 'INPUT GEPR�FT PERSON' given"
    If Len(Inputs.Gepr�ftDatum) = 0 Then writelog LogWarning, ErrorSource & "no 'INPUT GEPR�FT DATUM' given"
    If Len(Inputs.Geb�ude) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEB�UDE' given"
    If Len(Inputs.Geb�udeteil) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEB�UDETEIL' given"
    If Len(Inputs.Geschoss) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GESCHOSS' given"
    If Len(Inputs.Plan�berschrift) = 0 Then writelog LogWarning, ErrorSource & "no 'INPUT PLAN�BERSCHRIFT' given"
    If Len(Inputs.LayoutGr�sse) = 0 Then writelog LogError, ErrorSource & "no 'INPUT FORMAT' given"
    If Len(Inputs.LayoutMasstab) = 0 Then writelog LogError, ErrorSource & "no 'INPUT MASSTAB' given"
    If Len(Inputs.LayoutPlanstand) = 0 Then writelog LogError, ErrorSource & "no 'INPUT PLANSTAND' given"
    ValidInputs = True

    writelog LogInfo, "Plankopf erstellt"

End Function

Public Sub Refresh()

    Dim dateToday            As String
    dateToday = Format$(Now(), "yymmdd")
    Select Case This.PLANTYP
    Case "PLA"
        This.LayoutName = WorksheetFunction.TextJoin _
                          ("-", True, This.Geb�udeKF, This.Geb�udeTeilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)

        This.PDFFileName = WorksheetFunction.TextJoin _
                           ("_", True, shPData.range("ADM_Projektnummer").value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday)
    Case "SCH"
        This.LayoutName = WorksheetFunction.TextJoin _
                          ("-", True, This.GewerkKF, This.AnlageTypKF, This.Geb�udeCode, This.GeschossCode, This.AnlageNummer)

        This.PDFFileName = WorksheetFunction.TextJoin _
                           ("_", True, shPData.range("ADM_Projektnummer").value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday)
    Case "PRI"
        This.LayoutName = WorksheetFunction.TextJoin _
                          ("-", True, This.Geb�udeKF, This.Geb�udeTeilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)

        This.PDFFileName = WorksheetFunction.TextJoin _
                           ("_", True, shPData.range("ADM_Projektnummer").value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday)
    Case Else
        This.LayoutName = WorksheetFunction.TextJoin _
                          ("-", True, This.Geb�udeKF, This.Geb�udeTeilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)

        This.PDFFileName = WorksheetFunction.TextJoin _
                           ("_", True, shPData.range("ADM_Projektnummer").value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday)
    End Select
    writelog LogInfo, "Planablage-Namen aktualisiert: " & This.PDFFileName
End Sub

Private Sub IPlankopf_AddIndex(NewValue As IIndex)

    This.Indexes.Add NewValue
    Set This.CurrentIndex = NewValue

End Sub

Private Sub IPlankopf_ClearIndex()

    Set This.Indexes = New Collection

End Sub


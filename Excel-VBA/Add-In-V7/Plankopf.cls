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
Option Explicit
'@Folder "Plankopf"
'@ModuleDescription "Plankopf Objekt f�r einfacheres Management von Plank�pfen."

Implements IPlankopf

Private Type TPlankopf
    '--- Identifiers ---
    ID As String
    IDTinLine As String
    Planart As String: PlanartKF As String
    Plantyp As String
    TinLineTyp As String
    '--- Zeichner ---
    GezeichnetPerson As String
    GezeichnetDatum As String
    Gepr�ftPerson As String
    Gepr�ftDatum As String
    '--- Geb�ude ---
    Geb�ude As String: Geb�udeKF As String: Geb�udeCode As String
    Geb�udeteil As String: Geb�udeteilKF As String
    Geschoss As String: GeschossKF As String: GeschossCode As String
    '--- Gewerk ---
    Gewerk As String: GewerkKF As String
    UnterGewerk As String: UnterGewerkKF As String
    '--- Planbezeichnung ---
    Plan�berschrift As String
    CustomPlan�berschrift As Boolean
    '--- Layout ---
    LayoutName As String
    LayoutGr�sse As String
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
End Type

Private This                 As TPlankopf
Private Inputs               As TPlankopf
Private pProjekt             As IProjekt

' -------------------------------
' --- Properties
' -------------------------------
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
    IPlankopf_Plantyp = This.Plantyp
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

Private Property Get IPlankopf_CustomPlan�berschrift() As Boolean
    IPlankopf_CustomPlan�berschrift = This.CustomPlan�berschrift
End Property

Private Property Get IPlankopf_LayoutName() As String
    IPlankopf_LayoutName = This.LayoutName
End Property

Private Property Get IPlankopf_PlanNummer() As String
    IPlankopf_PlanNummer = This.LayoutName
End Property

Private Property Get IPlankopf_LayoutGr�sse() As String
    IPlankopf_LayoutGr�sse = This.LayoutGr�sse
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
    IPlankopf_FolderName = This.Pfad
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

' -------------------------------
' --- Functions
' -------------------------------

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
       Optional ByVal Plantyp As String, _
       Optional ByVal TinLineID As String = "NEW", _
       Optional ByVal SkipValidation As Boolean = False, _
       Optional ByVal Plan�berschrift As String = "NEW", _
       Optional ByVal ID As String = "ID", _
       Optional ByVal Custom�berschrift As Boolean = False _
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
        .Plantyp = Plantyp
        .CustomPlan�berschrift = Custom�berschrift
        .SkipValidation = SkipValidation
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

Private Sub GetKF()

    On Error GoTo ErrHandler
    This.Geb�udeKF = Globals.shGeb�ude.range("1:1").Find(This.Geb�ude).Offset(1).value
    If This.Geb�udeKF = "GES" And Globals.shGeb�ude.range("D1").value = "" Then
        This.Geb�udeKF = vbNullString
    End If
    If This.Geb�udeKF = vbNullString Then
        This.Geb�udeCode = vbNullString
    Else
        This.Geb�udeCode = Globals.shGeb�ude.range("1:1").Find(This.Geb�ude).Offset(2).value
    End If
    This.GeschossCode = Globals.shGeb�ude.range("B:B").Find(This.Geschoss).Offset(, -1).value
    This.GeschossKF = Globals.shGeb�ude.range("B:B").Find(This.Geschoss).Offset(, 1).value
    This.Geb�udeteilKF = Globals.shPData.range("PRO_Geb�udeteil").Find(This.Geb�udeteil).Offset(0, 1).value
    This.GewerkKF = Globals.shPData.range("PRO_Hauptgewerk").Find(This.Gewerk).Offset(0, 1).value

    Dim col                  As Long
    Select Case This.Plantyp
        Case "PLA"
1           col = Application.Match(This.GewerkKF & " PLA", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).value
2           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).value
            ' Welcher K�rzel kommt vor die Stockwerkbezeichnung beim DWG?
            Select Case This.GewerkKF
                Case "ELE"
                    This.TinLineTyp = "EP"
                Case "BRA"
                    This.TinLineTyp = "BR"
                Case "TUE"
                    This.TinLineTyp = "TF"
            End Select
        Case "SCH"
3           col = Application.Match(This.GewerkKF & " SCH", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).value
4           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).value
            This.TinLineTyp = vbNullString
        Case "PRI"
5           col = Application.Match(This.GewerkKF & " PRI", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).value
6           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = "PR"
            This.TinLineTyp = "PR"
    End Select

    writelog LogInfo, "Kurzformen erstellt"
    Exit Sub
ErrHandler:
Debug.Print ">>>   "; err.description
    writelog LogError, "Beim erstellen der Kurzformen" & vbNewLine & _
                      err.Number & vbNewLine & err.description & vbNewLine & err.source

End Sub

Private Function GetFileNames()
'BUG Hier kommen noch nicht die korrekten Pf�de raus.
    ' --- DWG FileName
    Select Case This.Plantyp
        Case "PLA"
            This.DWGFileName = This.TinLineTyp & "_" & This.GeschossKF
        Case "SCH"
            This.DWGFileName = "1"
        Case "PRI"
            This.DWGFileName = This.TinLineTyp & "_" & This.UnterGewerkKF
    End Select
    ' --- XML FileName
    Select Case This.Plantyp
        Case "PLA"
            This.XMLFileName = "TinPlan_" & This.TinLineTyp & "_" & This.GeschossKF & ".xml"
        Case "SCH"
            This.XMLFileName = "1.xml"
        Case "PRI"
            This.XMLFileName = "TinPlan_" & This.TinLineTyp & "_" & This.UnterGewerkKF & ".xml"
    End Select
    ' --- Pfad
    Select Case This.Plantyp
        Case "PLA"
            If This.Geb�ude = "Gesamt" And Globals.shGeb�ude.range("D1").value = "" Then
                Select Case This.TinLineTyp
                    Case "EP"
                        This.Pfad = pProjekt.ProjektOrdnerCAD & "\01_EP\" & This.GeschossCode & "_" & This.GeschossKF
                    Case "BR"
                        This.Pfad = pProjekt.ProjektOrdnerCAD & "\06_BR\" & This.GeschossCode & "_" & This.GeschossKF
                    Case "TF"
                        This.Pfad = pProjekt.ProjektOrdnerCAD & "\05_TF\" & This.GeschossCode & "_" & This.GeschossKF
                End Select
            Else                                 ' Mehrere Geb�ude
                Select Case This.TinLineTyp
                    Case "EP"
                        This.Pfad = pProjekt.ProjektOrdnerCAD & "\01_EP\" & This.Geb�udeCode & "_" & This.Geb�udeKF & "\" & This.GeschossCode & "_" & This.GeschossKF
                    Case "BR"
                        This.Pfad = pProjekt.ProjektOrdnerCAD & "\06_BR\" & This.Geb�udeCode & "_" & This.Geb�udeKF & "\" & This.GeschossCode & "_" & This.GeschossKF
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
            Dim ind As Long
            Dim indstr As String
            ind = ArrayIndex(Array("OBJ", "STS", "SLS", "KOM", "MES", "NOT", "BRA", "EVA", "GAT", "ZUT", "SEC", "ERD", "PRO"), This.UnterGewerkKF)
            indstr = ind
            If Len(indstr) < 2 Then indstr = "0" & indstr
            This.Pfad = pProjekt.ProjektOrdnerCAD & "\03_PR"
            This.DWGFilePath = This.Pfad & "\" & indstr & "_" & This.UnterGewerkKF
            This.XMLFilePath = This.DWGFilePath
    End Select
    ' --- FilePath
End Function

Private Sub GetPlan�berschrift()

    Dim Plan�berschrift      As String
    If This.CustomPlan�berschrift Then Exit Sub
    If This.Plantyp = "PLA" Then
        ' �berschrift f�r Elektropl�ne
        Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.Geschoss, This.Geb�udeteil)
        If Len(Plan�berschrift) > Globals.maxlen Then
            Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, This.Geb�udeteil)
            If Len(Plan�berschrift) > maxlen Then
                Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, This.Geb�udeteilKF)
                If Len(Plan�berschrift) > maxlen Then
                    Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerkKF, This.GeschossKF, This.Geb�udeteilKF)
                    If Len(Plan�berschrift) > maxlen Then
                        Plan�berschrift = Application.WorksheetFunction.TextJoin(" ", True, This.PlanartKF, This.UnterGewerkKF, This.GeschossKF, This.Geb�udeteilKF)
                    End If
                End If
            End If
        End If
    Else
        ' �berschrift f�r Schemas
        Plan�berschrift = "Prinzipschema " & This.UnterGewerk
        If Len(Plan�berschrift) > Globals.maxlen Then
            Plan�berschrift = "Prinzipschema " & This.UnterGewerk
            If Len(Plan�berschrift) > maxlen Then
                Plan�berschrift = "PR " & This.UnterGewerk
            End If
        End If
    End If
    If This.Plan�berschrift <> Plan�berschrift And This.Plan�berschrift <> vbNullString Then
        Select Case MsgBox("Die Beschriftung: " & This.Plan�berschrift & " ist nicht Standard." & vbNewLine & "Soll diese mit der Standardm�ssigen �berschrift ersetzt werden?" & vbNewLine & Plan�berschrift, vbYesNo, "Beschriftung �berschreiben?")
            Case vbYes
                This.Plan�berschrift = Plan�berschrift
                writelog LogInfo, "Plan�berschrift erstellt: " & This.Plan�berschrift
            Case vbNo
                writelog Logwarning, "Plan�berschrift nicht erstellt: " & This.Plan�berschrift
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
    If Len(Inputs.Gepr�ftPerson) = 0 Then writelog Logwarning, ErrorSource & "no 'INPUT GEPR�FT PERSON' given"
    If Len(Inputs.Gepr�ftDatum) = 0 Then writelog Logwarning, ErrorSource & "no 'INPUT GEPR�FT DATUM' given"
    If Len(Inputs.Geb�ude) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEB�UDE' given"
    If Len(Inputs.Geb�udeteil) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEB�UDETEIL' given"
    If Len(Inputs.Geschoss) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GESCHOSS' given"
    If Len(Inputs.Plan�berschrift) = 0 Then writelog Logwarning, ErrorSource & "no 'INPUT PLAN�BERSCHRIFT' given"
    If Len(Inputs.LayoutGr�sse) = 0 Then writelog LogError, ErrorSource & "no 'INPUT FORMAT' given"
    If Len(Inputs.LayoutMasstab) = 0 Then writelog LogError, ErrorSource & "no 'INPUT MASSTAB' given"
    If Len(Inputs.LayoutPlanstand) = 0 Then writelog LogError, ErrorSource & "no 'INPUT PLANSTAND' given"
    ValidInputs = True

    writelog LogInfo, "Plankopf erstellt"

End Function

Public Sub Refresh()

    Dim dateToday            As String
    dateToday = Format(Now(), "yymmdd")
    'Select Case This.PlanTyp
    This.LayoutName = WorksheetFunction.TextJoin _
                      ("-", True, This.Geb�udeKF, This.Geb�udeteilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)
                      
    This.PDFFileName = WorksheetFunction.TextJoin _
                       ("_", True, shPData.range("ADM_Projektnummer").value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday)
    writelog LogInfo, "Planablage-Namen aktualisiert: " & This.PDFFileName
End Sub

' -------------------------------
' --- Subs
' -------------------------------
Private Sub IPlankopf_AddIndex(NewValue As IIndex)

    This.Indexes.Add NewValue
    Set This.CurrentIndex = NewValue

End Sub

Private Sub IPlankopf_ClearIndex()

    Set This.Indexes = New Collection

End Sub



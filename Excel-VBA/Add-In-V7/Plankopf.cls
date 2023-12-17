VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Plankopf Objekt für einfacheres Management von Planköpfen."
Option Explicit
'@Folder "Plankopf"
'@ModuleDescription "Plankopf Objekt für einfacheres Management von Planköpfen."


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
    GeprüftPerson As String
    GeprüftDatum As String
    '--- Gebäude ---
    Gebäude As String: GebäudeKF As String
    Gebäudeteil As String: GebäudeteilKF As String
    Geschoss As String: GeschossKF As String: GeschossCode As String
    '--- Gewerk ---
    Gewerk As String: GewerkKF As String
    UnterGewerk As String: UnterGewerkKF As String
    '--- Planbezeichnung ---
    Planüberschrift As String
    CustomPlanüberschrift As Boolean
    '--- Layout ---
    LayoutName As String
    LayoutGrösse As String
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

Private Property Get IPlankopf_ID() As String
    IPlankopf_ID = This.ID
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

Private Property Get IPlankopf_GeprüftPerson() As String
    IPlankopf_GeprüftPerson = This.GeprüftPerson
End Property

Private Property Get IPlankopf_GeprüftDatum() As String
    IPlankopf_GeprüftDatum = This.GeprüftDatum
End Property

Private Property Get IPlankopf_Geprüft() As String
    IPlankopf_Geprüft = This.GeprüftPerson & " ; " & This.GeprüftDatum
End Property

Private Property Get IPlankopf_Gebäude() As String
    IPlankopf_Gebäude = This.Gebäude
End Property

Private Property Get IPlankopf_Gebäudeteil() As String
    IPlankopf_Gebäudeteil = This.Gebäudeteil
End Property

Private Property Get IPlankopf_Geschoss() As String
    IPlankopf_Geschoss = This.Geschoss
End Property

Private Property Get IPlankopf_Planüberschrift() As String
    IPlankopf_Planüberschrift = This.Planüberschrift
End Property

Private Property Get IPlankopf_CustomPlanüberschrift() As Boolean
    IPlankopf_CustomPlanüberschrift = This.CustomPlanüberschrift
End Property

Private Property Get IPlankopf_LayoutName() As String
    IPlankopf_LayoutName = This.LayoutName
End Property

Private Property Get IPlankopf_PlanNummer() As String
    IPlankopf_PlanNummer = This.LayoutName
End Property

Private Property Get IPlankopf_LayoutGrösse() As String
    IPlankopf_LayoutGrösse = This.LayoutGrösse
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
       ByVal GeprüftPerson As String, _
       ByVal GeprüftDatum As String, _
       ByVal Gebäude As String, _
       ByVal Gebäudeteil As String, _
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
       Optional ByVal Planüberschrift As String = "NEW", _
       Optional ByVal ID As String = "NEW", _
       Optional ByVal CustomÜberschrift As Boolean = False _
       ) As Boolean
    ' SkipValidation ist zu debug zwecken zu gebrauchen
    With Inputs
        .IDTinLine = TinLineID
        .GezeichnetPerson = GezeichnetPerson
        .GezeichnetDatum = GezeichnetDatum
        .GeprüftPerson = GeprüftPerson
        .GeprüftDatum = GeprüftDatum
        .Gebäude = Gebäude
        .Gebäudeteil = Gebäudeteil
        .Gewerk = Gewerk
        .UnterGewerk = UnterGewerk
        .Geschoss = Geschoss
        .Planüberschrift = Planüberschrift
        .LayoutGrösse = Format
        .LayoutMasstab = Masstab
        .LayoutPlanstand = Stand
        .Planart = Planart
        .Plantyp = Plantyp
        .CustomPlanüberschrift = CustomÜberschrift
    End With
    Set pProjekt = Projekt
    'On Error GoTo NotValid
    If SkipValidation Or ValidInputs Then
        This = Inputs
        If ID = "NEW" Then
            This.ID = getNewID(6, Globals.shStoreData, shStoreData.range("A1").CurrentRegion, 1)
        Else
            This.ID = ID
        End If

        GetKF
        GetFileNames
        GetPlanüberschrift
        Refresh
    End If
    On Error GoTo 0
    Filldata = True
    Exit Function

    On Error GoTo -1
    On Error GoTo 0
    Dim frm                  As New UserFormMessage
    frm.Typ TypWarning, "Es wurde kein Plankopf erstellt." & vbNewLine & _
                    "öffne den Log um die Fehlenden Inhalte zu sehen.", _
                    "Kein Plankopf erstellt", True
    frm.Show
    Filldata = False
    Exit Function

End Function

Private Sub GetKF()

    On Error GoTo ERR
    This.GebäudeKF = Globals.shGebäude.range("1:1").Find(This.Gebäude).Offset(1).Value
    This.GeschossCode = Globals.shGebäude.range("B:B").Find(This.Geschoss).Offset(, -1).Value
    This.GeschossKF = Globals.shGebäude.range("B:B").Find(This.Geschoss).Offset(, 1).Value
    This.GebäudeteilKF = Globals.shPData.range("PRO_Gebäudeteil").Find(This.Gebäudeteil).Offset(0, 1).Value
    This.GewerkKF = Globals.shPData.range("PRO_Hauptgewerk").Find(This.Gewerk).Offset(0, 1).Value

    Dim col                  As Long
    Select Case This.Plantyp
        Case "PLA"
1           col = Application.Match(This.GewerkKF & " PLA", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).Value
2           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).Value
            ' Welcher Kürzel kommt vor die Stockwerkbezeichnung beim DWG?
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
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).Value
4           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = Globals.shPData.Columns(col).Find(This.Planart).Offset(0, 1).Value
            This.TinLineTyp = vbNullString
        Case "PRI"
5           col = Application.Match(This.GewerkKF & " PRI", Globals.shPData.range("10:10"), 0)
            This.UnterGewerkKF = Globals.shPData.Columns(col).Find(This.UnterGewerk).Offset(0, 1).Value
6           col = Application.Match(This.GewerkKF, Globals.shPData.range("9:9"), 0)
            This.PlanartKF = "PR"
            This.TinLineTyp = "PR"
    End Select

    writelog LogInfo, "Kurzformen erstellt"
    Exit Sub
ERR:
Debug.Print ">>>   "; ERR.description
    writelog LogError, "Beim erstellen der Kurzformen" & vbNewLine & _
                     ERR.Number & vbNewLine & ERR.description & vbNewLine & ERR.source

End Sub

Private Function GetFileNames()
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
            Select Case This.TinLineTyp
                Case "EP"
                    This.Pfad = pProjekt.ProjektOrdnerCAD & "\01_EP\"
                Case "BR"
                    This.Pfad = pProjekt.ProjektOrdnerCAD & "\06_BR\"
                Case "TF"
                    This.Pfad = pProjekt.ProjektOrdnerCAD & "\05_TF\"
            End Select
        Case "SCH"
            This.Pfad = pProjekt.ProjektOrdnerCAD & "\02_ES\" & This.LayoutName
        Case "PRI"
            This.Pfad = pProjekt.ProjektOrdnerCAD & "\03_PR\"
    End Select
    ' --- FilePath
    This.DWGFilePath = This.Pfad & "\" & This.DWGFileName
    This.XMLFilePath = This.Pfad & "\" & This.XMLFileName
End Function

Private Sub GetPlanüberschrift()

    Dim Planüberschrift      As String
    If This.CustomPlanüberschrift Then Exit Sub
    If This.Plantyp = "PLA" Then
        ' Überschrift für Elektropläne
        Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.Geschoss, This.Gebäudeteil)
        If Len(Planüberschrift) > Globals.maxlen Then
            Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, This.Gebäudeteil)
            If Len(Planüberschrift) > maxlen Then
                Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerk, This.GeschossKF, This.GebäudeteilKF)
                If Len(Planüberschrift) > maxlen Then
                    Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.Planart, This.UnterGewerkKF, This.GeschossKF, This.GebäudeteilKF)
                    If Len(Planüberschrift) > maxlen Then
                        Planüberschrift = Application.WorksheetFunction.TextJoin(" ", True, This.PlanartKF, This.UnterGewerkKF, This.GeschossKF, This.GebäudeteilKF)
                    End If
                End If
            End If
        End If
    Else
        ' Überschrift für Schemas
        Planüberschrift = "Prinzipschema " & This.UnterGewerk
        If Len(Planüberschrift) > Globals.maxlen Then
            Planüberschrift = "Prinzipschema " & This.UnterGewerk
            If Len(Planüberschrift) > maxlen Then
                Planüberschrift = "PR " & This.UnterGewerk
            End If
        End If
    End If
    If This.Planüberschrift <> Planüberschrift And This.Planüberschrift <> vbNullString Then
    Select Case MsgBox("Die Beschriftung: " & This.Planüberschrift & " ist nicht Standard." & vbNewLine & "Soll diese mit der Standardmässigen Überschrift ersetzt werden?" & vbNewLine & Planüberschrift, vbYesNo, "Beschriftung Überschreiben?")
    Case vbYes
    This.Planüberschrift = Planüberschrift
    writelog LogInfo, "Planüberschrift erstellt: " & This.Planüberschrift
    Case vbNo
    writelog logwarning, "Planüberschrift nicht erstellt: " & This.Planüberschrift
    End Select
    Else
    This.Planüberschrift = Planüberschrift
    writelog LogInfo, "Planüberschrift erstellt: " & This.Planüberschrift
    End If

End Sub

Private Function ValidInputs() As Boolean

    ValidInputs = False
    Dim ErrorSource          As String
    ErrorSource = "Plankopf > ValidInputs > "
    If Len(Inputs.GezeichnetPerson) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEZEICHNET PERSON' given"
    If Len(Inputs.GezeichnetDatum) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEZEICHNET DATUM' given"
    If Len(Inputs.GeprüftPerson) = 0 Then writelog logwarning, ErrorSource & "no 'INPUT GEPRÜFT PERSON' given"
    If Len(Inputs.GeprüftDatum) = 0 Then writelog logwarning, ErrorSource & "no 'INPUT GEPRÜFT DATUM' given"
    If Len(Inputs.Gebäude) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEBÄUDE' given"
    If Len(Inputs.Gebäudeteil) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GEBÄUDETEIL' given"
    If Len(Inputs.Geschoss) = 0 Then writelog LogError, ErrorSource & "no 'INPUT GESCHOSS' given"
    If Len(Inputs.Planüberschrift) = 0 Then writelog logwarning, ErrorSource & "no 'INPUT PLANÜBERSCHRIFT' given"
    If Len(Inputs.LayoutGrösse) = 0 Then writelog LogError, ErrorSource & "no 'INPUT FORMAT' given"
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
                      ("-", True, This.GebäudeKF, This.GebäudeteilKF, This.GeschossKF, This.PlanartKF, This.UnterGewerkKF)
    This.PDFFileName = WorksheetFunction.TextJoin _
                       ("_", True, shPData.range("ADM_Projektnummer").Value, "04" & This.GewerkKF, This.GeschossCode, This.LayoutName, dateToday) & ".PDF"

    writelog LogInfo, "Planablage-Namen aktualisiert" & This.PDFFileName
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



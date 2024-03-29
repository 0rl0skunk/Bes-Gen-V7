VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface f�r 'Plankopf.cls'."

'@Folder "Plankopf"
'@Interface
'@ModuleDescription "Interface f�r 'Plankopf.cls'."

Option Explicit

Public Property Get Ansicht() As String
End Property

Public Property Get ID() As String
End Property
Public Property Let ID(ByVal NewValue As String)
End Property

Public Property Get IDTinLine() As String
End Property

Public Property Get Geb�udeAdresse() As IAdresse
End Property

Public Property Get TinLinePKNr() As String
End Property
Public Property Let TinLinePKNr(ByVal NewPKNr As String)
End Property

Public Property Let IDTinLine(ByVal newID As String)
End Property

Public Property Get PLANTYP() As String
End Property

Public Property Get Planart() As String
End Property

Public Property Get GezeichnetPerson() As String
End Property

Public Property Get GezeichnetDatum() As String
End Property

Public Property Get Gezeichnet() As String
End Property

Public Property Get Gepr�ftPerson() As String
End Property

Public Property Get Gepr�ftDatum() As String
End Property

Public Property Get Gepr�ft() As String
End Property

Public Property Get Geb�ude() As String
End Property

Public Property Get Geb�udeteil() As String
End Property

Public Property Get Geschoss() As String
End Property

Public Property Get Gewerk() As String
End Property

Public Property Get UnterGewerk() As String
End Property

Public Property Get UnterGewerkKF() As String
End Property

Public Property Get Plan�berschrift() As String
End Property

Public Property Get PlanBeschrieb() As String
End Property

Public Property Get CustomPlan�berschrift() As Boolean
End Property

Public Property Get LayoutName() As String
End Property

Public Property Get Plannummer() As String
End Property

Public Property Get LayoutGr�sse(Optional ByVal Layout As Boolean = False) As String
End Property

Public Property Get LayoutMasstab() As String
End Property

Public Property Get LayoutPlanstand() As String
End Property

Public Property Get PDFFileName() As String
End Property

Public Property Get DWGFileName() As String
End Property

Public Property Get dwgFile() As String
End Property

Public Property Get XMLFileName() As String
End Property

Public Property Get XMLFile() As String
End Property

Public Property Get FolderName() As String
End Property

Public Property Get CurrentIndex() As IIndex
End Property

Public Property Get Indexes() As Collection
End Property
Public Property Set Indexes(ByVal NewIndexes As Collection)
End Property

Public Sub AddIndex(ByRef NewValue As IIndex)
End Sub

Public Sub ClearIndex()
End Sub

Public Property Get AnlageTyp() As String
End Property

Public Property Get AnlageNummer() As String
End Property

Public Property Get UnterProjekt() As String
End Property


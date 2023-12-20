VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface für 'Plankopf.cls'."
'@Folder "Plankopf"
'@Interface
'@ModuleDescription "Interface für 'Plankopf.cls'."


Option Explicit

Public Property Get ID() As String
End Property
Public Property Let ID(ByVal NewValue As String)
End Property

Public Property Get IDTinLine() As String
End Property

Public Property Get TinLinePKNr() As String
End Property

Public Property Let TinLinePKNr(ByVal NewPKNr As String)
End Property

Public Property Let IDTinLine(ByVal newID As String)
End Property

Public Property Get Plantyp() As String
End Property

Public Property Get Planart() As String
End Property

Public Property Get GezeichnetPerson() As String
End Property

Public Property Get GezeichnetDatum() As String
End Property

Public Property Get Gezeichnet() As String
End Property

Public Property Get GeprüftPerson() As String
End Property

Public Property Get GeprüftDatum() As String
End Property

Public Property Get Geprüft() As String
End Property

Public Property Get Gebäude() As String
End Property

Public Property Get Gebäudeteil() As String
End Property

Public Property Get Geschoss() As String
End Property

Public Property Get Gewerk() As String
End Property

Public Property Get UnterGewerk() As String
End Property

Public Property Get UnterGewerkKF() As String
End Property

Public Property Get Planüberschrift() As String
End Property

Public Property Get PlanBeschrieb() As String
End Property

Public Property Get CustomPlanüberschrift() As Boolean
End Property

Public Property Get LayoutName() As String
End Property

Public Property Get Plannummer() As String
End Property

Public Property Get LayoutGrösse(Optional ByVal Layout As Boolean = False) As String
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


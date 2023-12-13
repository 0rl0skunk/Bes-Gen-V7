VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlankopf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Plankopf"
'@Interface

Option Explicit

Public Property Get ID() As String
End Property

Public Property Let ID(ByVal NewValue As String)
End Property

Public Property Get IDTinLine() As String
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

Public Property Get Gepr�ftPerson() As String
End Property

Public Property Get Gepr�ftDatum() As String
End Property

Public Property Get Gepr�ft() As String
End Property

Public Property Get Geb�ude() As String
End Property

Public Property Get Geb�udeTeil() As String
End Property

Public Property Get Geschoss() As String
End Property

Public Property Get Gewerk() As String
End Property

Public Property Get UnterGewerk() As String
End Property

Public Property Get Klartext() As String
End Property

Public Property Get Plan�berschrift() As String
End Property

Public Property Get LayoutName() As String
End Property

Public Property Get PlanNummer() As String
End Property

Public Property Get LayoutGr�sse() As String
End Property

Public Property Get LayoutMasstab() As String
End Property

Public Property Get LayoutPlanstand() As String
End Property

Public Property Get PDFFileName() As String
End Property

Public Property Get DWGFileName() As String
End Property

Public Property Get DWGFile() As String
End Property

Public Property Get DWGFilePath() As String
End Property

Public Property Get XMLFileName() As String
End Property

Public Property Get XMLFilePath() As String
End Property

Public Property Get xmlfile() As String
End Property

Public Property Get currentIndex() As IIndex
End Property

Public Property Get indexes() As Collection
End Property

Public Property Set indexes(ByVal NewIndexes As Collection)
End Property

Public Sub AddIndex(ByRef NewValue As IIndex)
End Sub

Public Sub ClearIndex()
End Sub
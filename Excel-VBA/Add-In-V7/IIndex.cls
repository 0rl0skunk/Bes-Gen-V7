VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Index")
'@Interface
'@ModuleDescription "Interface für 'Index.cls'."

Option Explicit

Public Property Get IndexID() As String
End Property

Public Property Get Index() As String
End Property
Public Property Let Index(NewValue As String)
End Property

Public Property Get Gezeichnet() As String
End Property

Public Property Get Geprüft() As String
End Property

Public Property Get Klartext() As String
End Property

Public Property Get PlanID() As String
End Property
Public Property Let PlanID(ByVal NewValue As String)
End Property


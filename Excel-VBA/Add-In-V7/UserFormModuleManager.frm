VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormModuleManager 
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "UserFormModuleManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormModuleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Excel-Items")
Public pModules              As New Collection

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub UserForm_Initialize()

    Me.TitleLabel.Caption = "Module-Manager"

End Sub

Public Sub SetInstructions(ByVal Instruction As String)

    Me.LabelInstructions.Caption = Instruction

End Sub

Public Sub SetModules(ByVal Modules As Collection)

    Set pModules = Modules
    LoadModules

End Sub

Private Sub LoadModules()

    Dim li                   As ListItem
    Dim e                    As Long

    With Me.ListViewModule
        .ListItems.Clear
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            'max width = 300
            .Add , , "", 0
            .Add , , "FileName", 140
            .Add , , "Modified", 140
        End With

        For e = 1 To pModules.Count
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , pModules.Item(e)(0)
            li.ListSubItems.Add , , pModules.Item(e)(1)
        Next
    End With

End Sub


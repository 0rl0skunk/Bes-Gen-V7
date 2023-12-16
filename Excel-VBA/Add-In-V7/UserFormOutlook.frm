VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormOutlook 
   ClientHeight    =   14250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   30420
   OleObjectBlob   =   "UserFormOutlook.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule VariableNotUsed
Option Explicit

'@Folder("Outlook")
Private pMailTo              As New Collection
Private pMailCC              As New Collection
Private pPlanköpfe           As New Collection

Private Sub CommandButton1_Click()
    Dim appOutlook           As New Outlook.Application
    Dim Mail                 As MailItem

    Set Mail = appOutlook.CreateItem(olMailItem)

    Mail.To = MailRecepientsTO
    Mail.CC = MailRecepientsCC
    Mail.Subject = Globals.Projekt.Projektnummer & " | Planversand " & Format(Now, "DD.MM.YYYY")
    Mail.Body = Anrede & vbNewLine & vbNewLine & Me.TextBoxFreitext.Value & vbNewLine & Planliste
    Mail.Display 0
End Sub

Private Function Planliste() As String
    Dim e                    As IPlankopf
    For Each e In pPlanköpfe
        Planliste = Planliste & e.Plannummer & vbNewLine
    Next
    Planliste = "Im Anhang finden sie Folgende Pläne: " & vbNewLine & Planliste
End Function

Private Function Anrede() As String
    If pMailTo.Count > 1 Then
        Anrede = "Hallo Zusammen"
    Else
        If pMailTo.Item(1).Anrede = "Du" Then
            Anrede = "Hallo " & pMailTo.Item(1).Vorname
        Else
            Anrede = "Guten Tag " & pMailTo.Item(1).Anrede & " " & pMailTo.Item(1).Nachname
        End If
    End If
End Function

Private Function MailRecepientsTO() As String

    Dim Person               As IPerson
    For Each Person In pMailTo
        MailRecepientsTO = MailRecepientsTO & " ; " & Person.EMail
    Next

End Function

Private Function MailRecepientsCC() As String

    Dim Person               As IPerson
    For Each Person In pMailCC
        MailRecepientsCC = MailRecepientsCC & " ; " & Person.EMail
    Next

End Function

Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

Private Sub ListViewMailTo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim li                   As ListItem
    Set pMailTo = New Collection
    For Each li In Me.ListViewMailTo.ListItems
        If li.Checked Then
            pMailTo.Add PersonFactory.LoadFromDataBase(Globals.shAdress.range("ADR_Adressen").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next
End Sub

Private Sub ListViewMailCC_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim li                   As ListItem
    Set pMailCC = New Collection
    For Each li In Me.ListViewMailCC.ListItems
        If li.Checked Then
            pMailCC.Add PersonFactory.LoadFromDataBase(Globals.shAdress.range("ADR_Adressen").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next
End Sub

Private Sub ListViewPlankopf_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim li                   As ListItem
    Set pPlanköpfe = New Collection
    For Each li In Me.ListViewPlankopf.ListItems
        If li.Checked Then
            pPlanköpfe.Add PlankopfFactory.LoadFromDataBase(Globals.shStoreData.range("A:A").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next
End Sub

Private Sub UserForm_Initialize()
    LoadListViewPlan
    LoadListViewMail Me.ListViewMailTo
    LoadListViewMail Me.ListViewMailCC
End Sub

Private Sub LoadListViewPlan()

    Dim Pla                  As IPlankopf
    Dim li                   As ListItem

    Dim row                  As Long
    Dim lastrow              As Long


    With Me.ListViewPlankopf
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxes = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , vbNullString, 20            ' 0
            .Add , , "ID", 0                     ' 1
            .Add , , "Plannummer"                ' 2
            .Add , , "Geschoss"                  ' 3
            .Add , , "Gebäude"                   ' 4
            .Add , , "Gebäudeteil"               ' 5
            .Add , , "Gewerk", 0                 ' 6
            .Add , , "Untergewerk", 0            ' 7
            .Add , , "Planart", 0                ' 8
            .Add , , "Gezeichnet"                ' 9
            .Add , , "Geprüft"                   ' 10
            .Add , , "Index"                     ' 11
        End With
        If Globals.shStoreData Is Nothing Then Globals.SetWBs
        lastrow = Globals.shStoreData.range("A1").CurrentRegion.rows.Count
        For row = 3 To lastrow
            Set Pla = PlankopfFactory.LoadFromDataBase(row)
            'Planköpfe.Add Pla                    ', Pla.ID
            Set li = .ListItems.Add()
            li.ListSubItems.Add , , Pla.ID
            li.ListSubItems.Add , , Pla.Plannummer
            li.ListSubItems.Add , , Pla.Geschoss
            li.ListSubItems.Add , , Pla.Gebäude
            li.ListSubItems.Add , , Pla.GebäudeTeil
            li.ListSubItems.Add , , Pla.Gewerk
            li.ListSubItems.Add , , Pla.UnterGewerk
            li.ListSubItems.Add , , Pla.Planart
            li.ListSubItems.Add , , Pla.Gezeichnet
            li.ListSubItems.Add , , Pla.Geprüft
            li.ListSubItems.Add , , Pla.currentIndex.Index
        Next row
    End With

End Sub

Private Sub LoadListViewMail(ByRef control As ListView)

    Dim li                   As ListItem

    Dim row                  As range
    Dim lastrow              As Long


    With control
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxes = True
        .Gridlines = True
        .FullRowSelect = True
        With .ColumnHeaders
            .Clear
            .Add , , vbNullString, 20            ' 0
            .Add , , "ID", 0                     ' 0
            .Add , , "Anrede", 0                 ' 1
            .Add , , "Vorname"                   ' 2
            .Add , , "Nachname"                  ' 3
            .Add , , "Firma"                     ' 4
            .Add , , "E-Mail"                    ' 5
        End With
        If Globals.shAdress Is Nothing Then Globals.SetWBs
        lastrow = Globals.shAdress.range("ADR_Adressen").rows.Count
        For Each row In Globals.shAdress.range("ADR_Adressen").rows
            Set li = .ListItems.Add()
            With row.Resize(1, 1)
                li.ListSubItems.Add , , .Offset(0, 8).Value
                li.ListSubItems.Add , , .Offset(0, 7).Value
                li.ListSubItems.Add , , .Offset(0, 1).Value
                li.ListSubItems.Add , , .Offset(0, 0).Value
                li.ListSubItems.Add , , .Offset(0, 2).Value
                li.ListSubItems.Add , , .Offset(0, 6).Value
            End With
        Next row
    End With

End Sub


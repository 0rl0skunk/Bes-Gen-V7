VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormOutlook 
   ClientHeight    =   9240.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720.001
   OleObjectBlob   =   "UserFormOutlook.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "E-Mails direkt vom Beschriftungsgenerator erstellen und versenden."

'@Folder("Outlook")
'@ModuleDescription "E-Mails direkt vom Beschriftungsgenerator erstellen und versenden."
'@Version "Release V1.0.0"

Option Explicit

Private pMailTo              As New Collection
Private pMailCC              As New Collection
Private pPlank�pfe           As New Collection
Private icons                As UserFormIconLibrary

Private Sub CommandButton1_Click()

    Dim PrintPath            As String
    Dim appOutlook           As New Outlook.Application
    Dim Mail                 As MailItem
    Set Mail = appOutlook.CreateItem(olMailItem)

    If Me.CheckBoxPlot Then
        ' wenn die Pl�ne neu geplottet werden sollen
        Dim pPlankopf        As IPlankopf
        PrintPath = CreatePlotList(pPlank�pfe)
        For Each pPlankopf In Stapelplot.Planliste
            Mail.Attachments.Add PrintPath & "\" & pPlankopf.PDFFileName & ".pdf"
        Next
    End If
    Mail.To = MailRecepientsTO
    Mail.CC = MailRecepientsCC
    Mail.Subject = Me.TextBoxBetreff.value
    Mail.Body = Anrede & vbNewLine & vbNewLine & Me.TextBoxFreitext.value & vbNewLine & Planliste
    Mail.Display 0

    Unload Me

End Sub

Private Function Planliste() As String

    Dim e                    As IPlankopf
    For Each e In pPlank�pfe
        Planliste = Planliste & "- " & e.Plannummer & vbTab & e.PlanBeschrieb & vbNewLine
    Next
    Planliste = "Im Anhang finden sie Folgende Pl�ne: " & vbNewLine & Planliste

End Function

Private Function Anrede() As String

    On Error GoTo ErrHandler
    If pMailTo.Count > 1 Then
        Anrede = "Hallo Zusammen"
    Else
        If pMailTo.Item(1).Anrede = "Du" Then
            Anrede = "Hallo " & pMailTo.Item(1).Vorname
        Else
            Anrede = "Guten Tag " & pMailTo.Item(1).Anrede & " " & pMailTo.Item(1).Nachname & ","
        End If
    End If
    Exit Function

ErrHandler:
    Anrede = vbNullString

End Function

Private Function MailRecepientsTO() As String
    ' Formatiert die Personen im Format welches von Outlook verwendet wird.
    MailRecepientsTO = vbNullString
    Dim Person               As IPerson
    For Each Person In pMailTo
        MailRecepientsTO = MailRecepientsTO & " ; " & Person.EMail
    Next

End Function

Private Function MailRecepientsCC() As String
    ' Formatiert die Personen im Format welches von Outlook verwendet wird.
    MailRecepientsCC = vbNullString
    Dim Person               As IPerson
    For Each Person In pMailCC
        MailRecepientsCC = MailRecepientsCC & " ; " & Person.EMail
    Next

End Function

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub ListViewMailTo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ' Aktualisiert die Collection von den E-Mail Empf�nger gem�ss den neuen Angaben
    Dim li                   As ListItem
    Set pMailTo = New Collection
    For Each li In Me.ListViewMailTo.ListItems
        If li.Checked Then
            pMailTo.Add PersonFactory.LoadFromDataBase(Globals.shAdress.range("ADR_Adressen").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next

End Sub

Private Sub ListViewMailCC_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ' Aktualisiert die Collection von den E-Mail Empf�nger im CC gem�ss den neuen Angaben
    Dim li                   As ListItem
    Set pMailCC = New Collection
    For Each li In Me.ListViewMailCC.ListItems
        If li.Checked Then
            pMailCC.Add PersonFactory.LoadFromDataBase(Globals.shAdress.range("ADR_Adressen").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next

End Sub

Private Sub ListViewPlankopf_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ' Aktualisiert die Collection von den Pl�nen welche geschickt werden gem�ss den neuen Angaben
    Dim li                   As ListItem
    Set pPlank�pfe = New Collection
    For Each li In Me.ListViewPlankopf.ListItems
        If li.Checked Then
            pPlank�pfe.Add PlankopfFactory.LoadFromDataBase(Globals.shStoreData.range("A:A").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next

End Sub

Private Sub UserForm_Initialize()

    Globals.SetWBs
    LoadListViewPlan Me.ListViewPlankopf
    LoadListViewMail Me.ListViewMailTo
    LoadListViewMail Me.ListViewMailCC
    Me.TextBoxBetreff.value = Globals.Projekt.Projektnummer & " | Planversand " & Format$(Now, "DD.MM.YYYY")
    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconOutlook.Picture
    Me.TitleLabel.Caption = "E-Mail schreiben"
    Me.LabelInstructions.Caption = "E-Mail automatisch schreiben und Pl�ne anh�ngen"

End Sub

Private Sub LoadListViewMail(ByVal control As ListView)
    ' l�dt die erfassten Adressen in die Listview
    Dim li                   As ListItem
    Dim row                  As range
    Dim lastrow              As Long

    With control
        .ListItems.Clear
        .View = lvwReport
        .CheckBoxES = True
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
            .Add , , "E-Mail", 0                 ' 5
        End With
        If Globals.shAdress Is Nothing Then Globals.SetWBs
        lastrow = Globals.shAdress.range("ADR_Adressen").rows.Count
        For Each row In Globals.shAdress.range("ADR_Adressen").rows
            Set li = .ListItems.Add()
            With row.Resize(1, 1)
                li.ListSubItems.Add , , .Offset(0, 8).value
                li.ListSubItems.Add , , .Offset(0, 7).value
                li.ListSubItems.Add , , .Offset(0, 1).value
                li.ListSubItems.Add , , .Offset(0, 0).value
                li.ListSubItems.Add , , .Offset(0, 2).value
                li.ListSubItems.Add , , .Offset(0, 6).value
            End With
        Next row
    End With

End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormOutlook 
   ClientHeight    =   8880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720.001
   OleObjectBlob   =   "UserFormOutlook.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "E-Mails direkt vom Beschriftungsgenerator erstellen und Anzeigen."








'@Folder("Outlook")
'@ModuleDescription "E-Mails direkt vom Beschriftungsgenerator erstellen und versenden."

Option Explicit

Private pMailTo              As New Collection
Private pMailCC              As New Collection
Private pPlanköpfe           As New Collection
Private icons                As UserFormIconLibrary

Private Sub CheckBoxSelectAll_Click()
    ' alle Pläne auswählen
    Dim li As ListItem
    If Me.CheckBoxSelectAll.value Then
        For Each li In Me.ListViewPlankopf.ListItems
            li.Checked = Me.CheckBoxSelectAll.value
            pPlanköpfe.Add PlankopfFactory.LoadFromDataBase(Globals.shStoreData.range("A:A").Find(li.ListSubItems.Item(1).Text).row)
        Next li
    End If
End Sub

Private Sub CommandButton1_Click()

    Dim PrintPath            As String
    Dim appOutlook           As New Outlook.Application
    Dim Mail                 As MailItem
    Dim mailBody As String * 2048
    Dim mailStyle As String
    Dim strFreitext As String * 2048
    Set Mail = appOutlook.CreateItem(olMailItem)

    If Me.CheckBoxPlot Then
        ' wenn die Pläne neu geplottet werden sollen
        Dim pPlankopf        As IPlankopf
        PrintPath = CreatePlotList(pPlanköpfe)
        For Each pPlankopf In Stapelplot.Planliste
            Mail.Attachments.Add PrintPath & "\" & pPlankopf.PDFFileName & ".pdf"
        Next
    End If
    strFreitext = Replace(Me.TextBoxFreitext.value, vbNewLine, "<br/>")
    mailStyle = "<p STYLE='font-family:Calibri;font-size:11pt'/>"
    mailBody = "<p>" & Anrede & "</p>" & _
               "<p></p>" & _
               "<p>" & strFreitext & "</p>" & _
               "<p></p>"
    With Mail
        .To = MailRecepientsTO
        .CC = MailRecepientsCC
        .Subject = Me.TextBoxBetreff.value
        .Display 0
        .BodyFormat = olFormatHTML
        .HTMLBody = "<HTML><BODY>" & mailStyle & mailBody & "<p>" & Planliste & "</p>" & "</BODY></HTML>" & .HTMLBody
    
    End With
    writeToVersandliste
    Unload Me

End Sub

Private Sub writeToVersandliste()

    Dim lastrow As Long
    With Globals.shVersand
        lastrow = .ListObjects("Versandliste").range.rows.Count + 4
        If lastrow = 6 And .Cells(6, 1).value <> vbNullString Then lastrow = 7
        .Cells(lastrow, 1).value = JoinCollection(pPlanköpfe)
        .Cells(lastrow, 2).value = MailRecepientsTO
        .Cells(lastrow, 3).value = Format(Now, "dd.MM.YYYY")
        .Cells(lastrow, 6).value = "X"
    End With

End Sub

Private Function JoinCollection(planköpfe As Collection) As String
    Dim pla As IPlankopf
    Dim str As String
    For Each pla In planköpfe
        str = str & pla.Plannummer & " | " & pla.CurrentIndex.Index & vbNewLine
    Next
    JoinCollection = str
End Function

Private Function Planliste() As String

    Dim e                    As IPlankopf
    Dim str As String
    Dim strPlanliste As String
    ' Pläne im Anhang als liste formatieren
    For Each e In pPlanköpfe
        strPlanliste = strPlanliste & "<li>" & e.Plannummer & " | " & e.PlanBeschrieb & "</li>"
    Next
    strPlanliste = "<ul>" & strPlanliste & "</ul>"
    
    On Error GoTo ErrHandler
    If pMailTo.Count > 1 Then
        str = "<p>Im Anhang finden Sie Folgende Pläne:</p>"
    ElseIf pMailTo.Count = 1 Then
        If pMailTo.Item(1).Anrede = "Du" Then
            str = "<p>Im Anhang findest du Folgende Pläne:</p>"
        Else
            str = "<p>Im Anhang finden Sie Folgende Pläne:</p>"
        End If
    End If
    Planliste = str & vbNewLine & strPlanliste
    Exit Function

ErrHandler:
    Planliste = "<p>Im Anhang sind Folgende Pläne:</p>" & vbNewLine & strPlanliste

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
    ' Aktualisiert die Collection von den E-Mail Empfänger gemäss den neuen Angaben
    Dim li                   As ListItem
    Set pMailTo = New Collection
    For Each li In Me.ListViewMailTo.ListItems
        If li.Checked Then
            pMailTo.Add PersonFactory.LoadFromDataBase(Globals.shAdress.range("ADR_Adressen").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next

End Sub

Private Sub ListViewMailCC_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ' Aktualisiert die Collection von den E-Mail Empfänger im CC gemäss den neuen Angaben
    Dim li                   As ListItem
    Set pMailCC = New Collection
    For Each li In Me.ListViewMailCC.ListItems
        If li.Checked Then
            pMailCC.Add PersonFactory.LoadFromDataBase(Globals.shAdress.range("ADR_Adressen").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next

End Sub

Private Sub ListViewPlankopf_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ' Aktualisiert die Collection von den Plänen welche geschickt werden gemäss den neuen Angaben
    Dim li                   As ListItem
    Set pPlanköpfe = New Collection
    For Each li In Me.ListViewPlankopf.ListItems
        If li.Checked Then
            pPlanköpfe.Add PlankopfFactory.LoadFromDataBase(Globals.shStoreData.range("A:A").Find(li.ListSubItems.Item(1).Text).row)
        End If
    Next
    Me.CheckBoxSelectAll.value = False

End Sub

Private Sub UserForm_Initialize()

    Globals.SetWBs
    LoadListViewPlan Me.ListViewPlankopf
    LoadListViewMail Me.ListViewMailTo
    LoadListViewMail Me.ListViewMailCC
    Me.TextBoxBetreff.value = Globals.Projekt.ProjektBezeichnung & " | Planversand " & Format$(Now, "DD.MM.YYYY")
    Set icons = New UserFormIconLibrary
    Me.TitleIcon.Picture = icons.IconOutlook.Picture
    Me.TitleLabel.Caption = "E-Mail schreiben"
    Me.LabelInstructions.Caption = "E-Mail automatisch schreiben und Pläne anhängen"

End Sub

Private Sub LoadListViewMail(ByVal control As ListView)
    ' lädt die erfassten Adressen in die Listview
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



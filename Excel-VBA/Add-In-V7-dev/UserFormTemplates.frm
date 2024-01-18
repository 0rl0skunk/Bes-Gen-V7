VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTemplates 
   ClientHeight    =   9000.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9600.001
   OleObjectBlob   =   "UserFormTemplates.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













'@Folder "Templates"
'@Version "Release V1.0.0"

Option Explicit

Private icons                As UserFormIconLibrary

Private Sub CheckBox1_Enter()

    Me.LabelToolTip.Caption = " >> CheckBox <<" & vbNewLine & _
                              vbNullString & vbNewLine & _
                              "> Use Case" & vbNewLine & _
                              "Is used instead of a Toggle Button. Shows thw value in a cleaner way and is more flexible with the Label next to it." & vbNewLine & _
                              "> Naming Convention" & vbNewLine & _
                              "The ChecBox is shortened with 'CBX' followed by one or two words." & vbNewLine & _
                              "> example" & vbNewLine & _
                              "  CBXObjectSnapping > shows the current Value for Object Snapping."

End Sub

Private Sub ComboBox1_Enter()

    Me.LabelToolTip.Caption = " >> ComboBox <<" & vbNewLine & _
                              vbNullString & vbNewLine & _
                              "> Use Case" & vbNewLine & _
                              "It is used to limit the user Input to a predefined selection." & vbNewLine & _
                              "> Naming Convention" & vbNewLine & _
                              "The Combobox is shortened with 'CBB' followed by one or two words." & vbNewLine & _
                              "> example" & vbNewLine & _
                              "  CBBProjectState > Shows a Combobox with the available Project States (52 Ausführung etc.)"

End Sub

Private Sub CommandButton1_Enter()

    Me.LabelToolTip.Caption = " >> CommabdButton <<" & vbNewLine & _
                              vbNullString & vbNewLine & _
                              "> Use Case" & vbNewLine & _
                              "It is used to 'do something. The Caption should depict what it does > close, load, select Folder." & vbNewLine & _
                              "> Naming Convention" & vbNewLine & _
                              "The Combobox is shortened with 'CB' followed by one or two words." & vbNewLine & _
                              "> example" & vbNewLine & _
                              "  CBClose > This CommandButton is used tho close a UserForm."

End Sub

Private Sub CommandButtonClose_Click()

    Unload Me

End Sub

Private Sub LabelToolTip_Click()

    Me.LabelToolTip.Caption = " >> ToolTip 'LabelControl' <<" & vbNewLine & _
                              vbNullString & vbNewLine & _
                              "> Use Case" & vbNewLine & _
                              "It is used to display the ToolTip for the selected Control. It should simplify the usage of a UserForm while not being to complicated." & vbNewLine & _
                              "> Naming Convention" & vbNewLine & _
                              "LabelToolTip. Thst is the whole Name it alwaays has!" & vbNewLine & _
                              "> example" & vbNewLine & _
                              "  CBClose > This CommandButton is used tho close a UserForm."

End Sub

Private Sub OptionButton1_Enter()

    Me.LabelToolTip.Caption = " >> OptionButton <<" & vbNewLine & _
                              vbNullString & vbNewLine & _
                              "> Use Case" & vbNewLine & _
                              "Whene there are a small amount of states a Value can be in and also just one at a time. They should be grouped by a Frame named with the broader option description." & vbNewLine & _
                              "> Naming Convention" & vbNewLine & _
                              "The OptionButton is shortened with 'OB' followed by one or two words." & vbNewLine & _
                              "> example" & vbNewLine & _
                              "  OBFormatHorizontal / OBFormatVertical [FRFormat] > shows 2 Option Buttons to select between Horizontal and Vertical options."

End Sub

Private Sub TextBox1_Enter()

    Me.LabelToolTip.Caption = " >> Textbox <<" & vbNewLine & _
                              vbNullString & vbNewLine & _
                              "> Use Case" & vbNewLine & _
                              "It is used to input and outpur string Values from and to the User." & vbNewLine & _
                              "> Naming Convention" & vbNewLine & _
                              "The Textbox Name is shortened with 'TB' followed by an O for Output or an I for Input. Following this is one or two words describing the Input or Output of the TextBox." & vbNewLine & _
                              "> example" & vbNewLine & _
                              "  TBOLayoutName > Outputs the Layout Name to the user"

End Sub



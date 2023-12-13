Attribute VB_Name = "IndexFactory"
Option Explicit
'@Folder("Index")
' a Factory 'creates the new class and binds it to the interface so only the wanted methods are exposed to the user

Public Function Create( _
       ByVal IDPlan As String, _
       ByVal GezeichnetPerson As String, _
       ByVal GezeichnetDatum As String, _
       ByVal Klartext As String, _
       Optional ByVal Letter As String, _
       Optional ByVal Gepr�ftPerson As String = vbNullString, _
       Optional ByVal Gepr�ftDatum As String = vbNullString _
       ) As IIndex

    Dim newIndex             As index
    Set newIndex = New index
    newIndex.FillData _
        IDPlan:=IDPlan, _
        Letter:=Letter, _
        GezeichnetPerson:=GezeichnetPerson, _
        GezeichnetDatum:=GezeichnetDatum, _
        Gepr�ftPerson:=Gepr�ftPerson, _
        Gepr�ftDatum:=Gepr�ftDatum, _
        Klartext:=Klartext

    Set Create = newIndex

End Function



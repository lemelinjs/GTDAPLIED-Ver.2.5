VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormulaireDeTri 
   Caption         =   "UserForm2"
   ClientHeight    =   1830
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   4614
   OleObjectBlob   =   "FormulaireDeTri.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormulaireDeTri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 Dim celluleA1 As String
    inboxValue = chercheInbox.Value
    Sheets("1-Collecte-clarification-org.").Select
    Range("TableauCollect[[#Headers],[Collecter - inbox]]").Select
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveSheet.ListObjects("TableauCollect").Range.AutoFilter Field:=1, _
        Criteria1:="=*" & inboxValue & "*", Operator:=xlAnd
    Selection.Offset(1, 0).Select
End Sub

Private Sub CommandButton2_Click()
    Dim projectValue As String
    projectValue = chercherProjet.Value
    Sheets("1-Collecte-clarification-org.").Select
    Range("A2").Select
    Selection.AutoFilter
    Selection.AutoFilter
    ActiveSheet.ListObjects("TableauCollect").Range.AutoFilter Field:=5, _
        Criteria1:="=*" & projectValue & "*", Operator:=xlAnd
    Selection.Offset(1, 0).Select
End Sub

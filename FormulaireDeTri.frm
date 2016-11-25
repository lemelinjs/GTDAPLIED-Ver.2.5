VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormulaireDeTri 
   Caption         =   "Formulaire de projet"
   ClientHeight    =   2958
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   4626
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
    ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort.SortFields.Clear
    ActiveSheet.ListObjects("TableauCollect").Range.AutoFilter Field:=1, _
        Criteria1:="=*" & inboxValue & "*", Operator:=xlAnd
    Selection.Offset(1, 0).Select
End Sub

Private Sub CommandButton2_Click()
    Dim projectValue As String
    projectValue = chercherProjet.Value
    ' Trier
     Call sorterColE
    Sheets("1-Collecte-clarification-org.").Select
    ActiveSheet.ListObjects("TableauCollect").Range.AutoFilter Field:=5, _
        Criteria1:="=*" & projectValue & "*", Operator:=xlAnd
    Selection.Offset(1, 0).Select
End Sub
Sub sorterColE()
'
'

    ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort.SortFields.Add Key:=Range( _
        "TableauCollect[Code de projet et de tâches]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub CommandButton3_Click()
    Range("A2").Select
    Selection.AutoFilter
    Selection.AutoFilter
End Sub

Private Sub CommandButton4_Click()
    ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort.SortFields.Clear
    ActiveSheet.ListObjects("TableauCollect").Range.AutoFilter Field:=4, _
        Criteria1:="Oui - Projet"
    Selection.Offset(1, 0).Select
End Sub

Private Sub CommandButton5_Click()
    On Error GoTo getOut
    
    ' Déclarer les variables
    Dim nombreLignes As Integer
    Dim Number_of_Rows As Integer
    Dim found As Range
    Dim projetAChercher As String
    Dim ligneProjet As String
    Dim dernierLigne As Long
    
    
    ' procédure pour savoir sur quelle ligne est le projet en
    ' question - sous l'appellation found.row
    If Left(projetArchiver, 1) <> "p" Then
        projetAChercher = "p" & projetArchiver
    End If
    
        Set found = Sheets("1-Collecte-clarification-org.").Columns("E").Find(what:=projetAChercher, LookIn:=xlValues, lookat:=xlWhole)
        ligneProjet = found.Row
    
      
       
    
    'Faire un sort pour placer toutes les cellules d'un même projet
    'Contigües l'une à l'autre
    
    Call sorterColE
    
    'Initialiser le filtre
    ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort.SortFields.Clear
    If Left(projetArchiver, 1) = "p" Then
        projetArchiver = Mid(projetArchiver, 2, 8)
    End If
        
    ActiveSheet.ListObjects("TableauCollect").Range.AutoFilter Field:=5, _
        Criteria1:="=*" & projetArchiver & "*", Operator:=xlAnd
    
    With ActiveSheet.ListObjects(1)
        For Each Line In .Range.SpecialCells(xlCellTypeVisible).Areas
            Number_of_Rows = Number_of_Rows + Line.Rows.Count
            'Vérifier ici si les tâches sont encore ouvertes.
        Next
    End With
    Number_of_Rows = Number_of_Rows - 1 + ligneProjet
    Rows(ligneProjet & ":" & Number_of_Rows - 1).Select
    Selection.Cut
    Sheets("Archives").Select
    derniereLigne = Range("A" & Rows.Count).End(xlUp).Row
    Rows(derniereLigne + 1 & ":" & derniereLigne + 1).Select
    ActiveSheet.Paste
    Rows(derniereLigne + 1 & ":" & derniereLigne + 1 + ligneProjet).EntireRow.AutoFit
    Sheets("1-Collecte-clarification-org.").Select
    Rows(ligneProjet & ":" & Number_of_Rows - 1).Select
    Selection.Delete Shift:=xlUp
    ActiveWorkbook.Worksheets("1-Collecte-clarification-org.").ListObjects( _
        "TableauCollect").Sort.SortFields.Clear
    'MsgBox Number_of_Rows - 1 'pour enlever l'entete
getOut:
End Sub
Private Sub cancel()
    Exit Sub
End Sub

Private Sub Image1_BeforeDragOver(ByVal cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label4_Click()

End Sub

Private Sub ListBox1_Click()


End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub

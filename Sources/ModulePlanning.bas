Attribute VB_Name = "ModulePlanning"
Option Explicit

'***********************************************************************
' Procédure   : FormulaireProjet
' Description : Appelle l'UserForm permettant aux utilisateurs d'ajouter
'   projet à leur planning
'***********************************************************************

Sub FormulaireProjet()
    'Déclaration stanciation des variables
    Dim FormProjet As ufProjet
    Set FormProjet = New ufProjet
    
    'Affiche l'USerForm
    FormProjet.Show
    
    'Vide l'userForm
    Set FormProjet = Nothing
End Sub

'***********************************************************************
' Procédure   : AjoutProjet
' Description : Ajoute une ligne dans le planning à l'aide des informations
'   passées en paramètres
' Paramètres  :
'   + sFeuille   : La feuille sur laquelle ajouter la ligne
'   + sProjet    : Le nom du projet à ajouter
'   + sTache     : La tache du projet à ajouter
'   + sPriorite  : La priorité de la tâche à ajouter
'   + sDureePrev : La durée prévue pour cette tâche
'***********************************************************************

Sub AjoutProjet(sFeuille As String, sProjet As String, sTache As String, sPriorite As String, sDureePrev As String)
    'Déclaration des variables
    Dim iLigne As Integer
    Dim wkFeuille As Worksheet
    
    'Instanciation des variables
    iLigne = 4
    Set wkFeuille = ActiveWorkbook.Sheets(sFeuille)
    
    'Récupère l'indice de la dernière ligne
    Do While wkFeuille.Cells(iLigne, 4).Value <> ""
        iLigne = iLigne + 1
    Loop
    
    'Si la ligne est la première
    If iLigne = 4 Then
        'Ajoute les informations de la tâche
        wkFeuille.Cells(iLigne, 2).Value = sProjet
        wkFeuille.Cells(iLigne, 3).Value = sTache
        wkFeuille.Cells(iLigne, 4).Value = sPriorite
        wkFeuille.Cells(iLigne, 5).Value = sDureePrev
    Else
        'On insère une ligne
        wkFeuille.Rows(iLigne & ":" & iLigne).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'Ajoute les informations de la tâche
        wkFeuille.Cells(iLigne, 2).Value = sProjet
        wkFeuille.Cells(iLigne, 3).Value = sTache
        wkFeuille.Cells(iLigne, 4).Value = sPriorite
        wkFeuille.Cells(iLigne, 5).Value = sDureePrev
    End If
    
    'On se postionne dans la celulle A1
    ActiveSheet.Range("A1").Activate
End Sub

'***********************************************************************
' Procédure   : FormulaireTache
' Description : Appelle l'UserForm permettant aux utilisateurs d'ajouter
'   une tâche aux projets qu'ils ont renseignée
'***********************************************************************

Sub FormulaireTache()
    'Déclaration stanciation des variables
    Dim FormTache As ufTache
    Set FormTache = New ufTache
    
    'Affiche l'USerForm
    FormTache.Show
    
    'Vide l'userForm
    Set FormTache = Nothing
End Sub

'***********************************************************************
' Procédure   : AjoutTaches
' Description : Ajoute une ligne dans le planning à l'aide des informations
'   passées en paramètres
' Paramètres  :
'   + sFeuille   : La feuille sur laquelle ajouter la ligne
'   + sProjet    : Le nom du projet à ajouter
'   + sTache     : La tache du projet à ajouter
'   + sPriorite  : La priorité de la tâche à ajouter
'   + sDureePrev : La durée prévue pour cette tâche
'***********************************************************************

Sub AjoutTaches(sFeuille As String, sProjet As String, sTache As String, sPriorite As String, sDureePrev As String)
    'Déclaration des variables
    Dim iLigne As Integer
    Dim iLigneFin As Integer
    Dim wkFeuille As Worksheet
    
    'Instanciation des variables
    iLigne = 4
    Set wkFeuille = ActiveWorkbook.Sheets(sFeuille)
    
    'Parcourt toutes les lignes du planning
    Do While wkFeuille.Cells(iLigne, 4).Value <> ""
        'Si on trouve la ligne du projet
        If wkFeuille.Cells(iLigne, 2).Value = sProjet Then
            iLigneFin = iLigne + 1
            'On cherche l'indice de la dernière ligne du projet
            Do While wkFeuille.Cells(iLigneFin, 2).Value = ""
                iLigneFin = iLigneFin + 1
            Loop
            
            'Insère une ligne de tâche
            wkFeuille.Rows(iLigneFin & ":" & iLigneFin).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            'Ajoute les informations de la tâche
            wkFeuille.Cells(iLigneFin, 3).Value = sTache
            wkFeuille.Cells(iLigneFin, 4).Value = sPriorite
            wkFeuille.Cells(iLigneFin, 5).Value = sDureePrev
        End If
        
        'Passe à la ligne suivante
        iLigne = iLigne + 1
    Loop
    
    'On se postionne dans la celulle A1
    ActiveSheet.Range("A1").Activate
End Sub

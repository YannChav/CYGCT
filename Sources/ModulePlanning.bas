Attribute VB_Name = "ModulePlanning"
Option Explicit

'***********************************************************************
' Proc�dure   : FormulaireProjet
' Description : Appelle l'UserForm permettant aux utilisateurs d'ajouter
'   projet � leur planning
'***********************************************************************

Sub FormulaireProjet()
    'D�claration stanciation des variables
    Dim FormProjet As ufProjet
    Set FormProjet = New ufProjet
    
    'Affiche l'USerForm
    FormProjet.Show
    
    'Vide l'userForm
    Set FormProjet = Nothing
End Sub

'***********************************************************************
' Proc�dure   : AjoutProjet
' Description : Ajoute une ligne dans le planning � l'aide des informations
'   pass�es en param�tres
' Param�tres  :
'   + sFeuille   : La feuille sur laquelle ajouter la ligne
'   + sProjet    : Le nom du projet � ajouter
'   + sTache     : La tache du projet � ajouter
'   + sPriorite  : La priorit� de la t�che � ajouter
'   + sDureePrev : La dur�e pr�vue pour cette t�che
'***********************************************************************

Sub AjoutProjet(sFeuille As String, sProjet As String, sTache As String, sPriorite As String, sDureePrev As String)
    'D�claration des variables
    Dim iLigne As Integer
    Dim wkFeuille As Worksheet
    
    'Instanciation des variables
    iLigne = 4
    Set wkFeuille = ActiveWorkbook.Sheets(sFeuille)
    
    'R�cup�re l'indice de la derni�re ligne
    Do While wkFeuille.Cells(iLigne, 4).Value <> ""
        iLigne = iLigne + 1
    Loop
    
    'Si la ligne est la premi�re
    If iLigne = 4 Then
        'Ajoute les informations de la t�che
        wkFeuille.Cells(iLigne, 2).Value = sProjet
        wkFeuille.Cells(iLigne, 3).Value = sTache
        wkFeuille.Cells(iLigne, 4).Value = sPriorite
        wkFeuille.Cells(iLigne, 5).Value = sDureePrev
    Else
        'On ins�re une ligne
        wkFeuille.Rows(iLigne & ":" & iLigne).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'Ajoute les informations de la t�che
        wkFeuille.Cells(iLigne, 2).Value = sProjet
        wkFeuille.Cells(iLigne, 3).Value = sTache
        wkFeuille.Cells(iLigne, 4).Value = sPriorite
        wkFeuille.Cells(iLigne, 5).Value = sDureePrev
    End If
    
    'On se postionne dans la celulle A1
    ActiveSheet.Range("A1").Activate
End Sub

'***********************************************************************
' Proc�dure   : FormulaireTache
' Description : Appelle l'UserForm permettant aux utilisateurs d'ajouter
'   une t�che aux projets qu'ils ont renseign�e
'***********************************************************************

Sub FormulaireTache()
    'D�claration stanciation des variables
    Dim FormTache As ufTache
    Set FormTache = New ufTache
    
    'Affiche l'USerForm
    FormTache.Show
    
    'Vide l'userForm
    Set FormTache = Nothing
End Sub

'***********************************************************************
' Proc�dure   : AjoutTaches
' Description : Ajoute une ligne dans le planning � l'aide des informations
'   pass�es en param�tres
' Param�tres  :
'   + sFeuille   : La feuille sur laquelle ajouter la ligne
'   + sProjet    : Le nom du projet � ajouter
'   + sTache     : La tache du projet � ajouter
'   + sPriorite  : La priorit� de la t�che � ajouter
'   + sDureePrev : La dur�e pr�vue pour cette t�che
'***********************************************************************

Sub AjoutTaches(sFeuille As String, sProjet As String, sTache As String, sPriorite As String, sDureePrev As String)
    'D�claration des variables
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
            'On cherche l'indice de la derni�re ligne du projet
            Do While wkFeuille.Cells(iLigneFin, 2).Value = ""
                iLigneFin = iLigneFin + 1
            Loop
            
            'Ins�re une ligne de t�che
            wkFeuille.Rows(iLigneFin & ":" & iLigneFin).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            'Ajoute les informations de la t�che
            wkFeuille.Cells(iLigneFin, 3).Value = sTache
            wkFeuille.Cells(iLigneFin, 4).Value = sPriorite
            wkFeuille.Cells(iLigneFin, 5).Value = sDureePrev
        End If
        
        'Passe � la ligne suivante
        iLigne = iLigne + 1
    Loop
    
    'On se postionne dans la celulle A1
    ActiveSheet.Range("A1").Activate
End Sub

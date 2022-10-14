Attribute VB_Name = "ModuleGestion"
Option Explicit

'***********************************************************************
' Procédure   : AjouterPlanning
' Description : Ajoute une feuille de Planning dans le classeur Excel
'***********************************************************************

Sub AjouterPlanning(sFeuilleOrigine As String)
    'Déclaration des variables
    Dim Feuille As Worksheet
    Dim iLigne As Integer
    Dim sNomFeuille As String
    Dim sChaineIndice As String
    Dim tabIndices As Variant
    Dim Indice As Variant
    
    sNomFeuille = "Semaine " & ActiveWorkbook.Sheets("Accueil").Range("C8").Value
    sChaineIndice = ""
    
    'Vérifie que la feuille n'existe pas
    For Each Feuille In Sheets
        If UCase(Feuille.Name) = UCase(sNomFeuille) Then
            'Affiche un message d'erreur à l'utilisateur
            MsgBox ("La feuille contenant le planning de la semaine en cours existe déjà.")
            'Quitte la fonction
            Exit Sub
        End If
    Next Feuille
    'Copie la feuille de modèle planning pour la semaine en cours
    Sheets(sFeuilleOrigine).Copy After:=Sheets("Accueil")
    'Renomme la feuille
    Sheets(sFeuilleOrigine & " (2)").Name = sNomFeuille
    
    'Si la feuille ajoutée n'est pas la feuille Modèle
    iLigne = 4
    If sFeuilleOrigine <> "MODELE SEMAINE" Then
        'Parcourt toutes les lignes du planning
        Do While Sheets(sNomFeuille).Cells(iLigne, 3).Value <> ""
            'Si la colonne Durée réelle est complétée
            If Sheets(sNomFeuille).Cells(iLigne, 6).Value <> "" Then
                'On vérifie que la ligne a supprimée n'est pas la première du projet
                If Sheets(sNomFeuille).Cells(iLigne, 2).Value <> "" Then
                    'On cherche si ce projet possède d'autres tâches
                    If Sheets(sNomFeuille).Cells(iLigne + 1, 2).Value = "" Then
                        'On met le nom du projet sur cette ligne
                        Sheets(sNomFeuille).Cells(iLigne + 1, 2).Value = Sheets(sNomFeuille).Cells(iLigne, 2).Value
                    End If
                End If
                'Ajoute les indices dans l'ordre décroisant
                If sChaineIndice = "" Then
                    sChaineIndice = CStr(iLigne)
                Else
                    sChaineIndice = CStr(iLigne) & ";" & sChaineIndice
                End If
            End If
            iLigne = iLigne + 1
        Loop
    End If
    
    'Récupère la chaine des indices classés dans l'ordre décroissant
    tabIndices = Split(sChaineIndice, ";")
    For Each Indice In tabIndices
        'Supprime la ligne
        ActiveWorkbook.Sheets(sNomFeuille).Rows(CInt(Indice)).Delete
    Next Indice
    
    'Montre la feuille et l'active
    ActiveWorkbook.Sheets(sNomFeuille).Visible = True
    ActiveWorkbook.Sheets(sNomFeuille).Activate
    
    'Sauvegarde le classeur Excel
    ActiveWorkbook.Save
End Sub

'***********************************************************************
' Procédure   : FormulairePlanning
' Description : Appelle l'UserForm permettant aux utilisateurs d'ajouter
'   une feuille de planning
'***********************************************************************

Sub FormulairePlanning()
    'Déclaration et instanciation des variables
    Dim FormPlanning As ufNewPlanning
    Set FormPlanning = New ufNewPlanning
    
    'Affiche l'USerForm
    FormPlanning.Show
    
    'Vide l'userForm
    Set FormPlanning = Nothing
End Sub

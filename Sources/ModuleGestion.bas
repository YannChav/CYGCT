Attribute VB_Name = "ModuleGestion"
Option Explicit

'***********************************************************************
' Proc�dure   : AjouterPlanning
' Description : Ajoute une feuille de Planning dans le classeur Excel
'***********************************************************************

Sub AjouterPlanning(sFeuilleOrigine As String)
    'D�claration des variables
    Dim Feuille As Worksheet
    Dim iLigne As Integer
    Dim sNomFeuille As String
    Dim sChaineIndice As String
    Dim tabIndices As Variant
    Dim Indice As Variant
    
    sNomFeuille = "Semaine " & ActiveWorkbook.Sheets("Accueil").Range("C8").Value
    sChaineIndice = ""
    
    'V�rifie que la feuille n'existe pas
    For Each Feuille In Sheets
        If UCase(Feuille.Name) = UCase(sNomFeuille) Then
            'Affiche un message d'erreur � l'utilisateur
            MsgBox ("La feuille contenant le planning de la semaine en cours existe d�j�.")
            'Quitte la fonction
            Exit Sub
        End If
    Next Feuille
    'Copie la feuille de mod�le planning pour la semaine en cours
    Sheets(sFeuilleOrigine).Copy After:=Sheets("Accueil")
    'Renomme la feuille
    Sheets(sFeuilleOrigine & " (2)").Name = sNomFeuille
    
    'Si la feuille ajout�e n'est pas la feuille Mod�le
    iLigne = 4
    If sFeuilleOrigine <> "MODELE SEMAINE" Then
        'Parcourt toutes les lignes du planning
        Do While Sheets(sNomFeuille).Cells(iLigne, 3).Value <> ""
            'Si la colonne Dur�e r�elle est compl�t�e
            If Sheets(sNomFeuille).Cells(iLigne, 6).Value <> "" Then
                'On v�rifie que la ligne a supprim�e n'est pas la premi�re du projet
                If Sheets(sNomFeuille).Cells(iLigne, 2).Value <> "" Then
                    'On cherche si ce projet poss�de d'autres t�ches
                    If Sheets(sNomFeuille).Cells(iLigne + 1, 2).Value = "" Then
                        'On met le nom du projet sur cette ligne
                        Sheets(sNomFeuille).Cells(iLigne + 1, 2).Value = Sheets(sNomFeuille).Cells(iLigne, 2).Value
                    End If
                End If
                'Ajoute les indices dans l'ordre d�croisant
                If sChaineIndice = "" Then
                    sChaineIndice = CStr(iLigne)
                Else
                    sChaineIndice = CStr(iLigne) & ";" & sChaineIndice
                End If
            End If
            iLigne = iLigne + 1
        Loop
    End If
    
    'R�cup�re la chaine des indices class�s dans l'ordre d�croissant
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
' Proc�dure   : FormulairePlanning
' Description : Appelle l'UserForm permettant aux utilisateurs d'ajouter
'   une feuille de planning
'***********************************************************************

Sub FormulairePlanning()
    'D�claration et instanciation des variables
    Dim FormPlanning As ufNewPlanning
    Set FormPlanning = New ufNewPlanning
    
    'Affiche l'USerForm
    FormPlanning.Show
    
    'Vide l'userForm
    Set FormPlanning = Nothing
End Sub

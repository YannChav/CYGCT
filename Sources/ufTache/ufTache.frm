VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTache 
   Caption         =   "Ajout d'une tâche"
   ClientHeight    =   3180
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5064
   OleObjectBlob   =   "ufTache.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufTache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    'Déclaration des varaibles
    Dim iLigne As Integer
    Dim wkFeuille As Worksheet

    'Instanciation des variables
    iLigne = 4
    Set wkFeuille = ActiveSheet

    'Compléte la ListBox avec les projets
    With Me.LST_PROJET
        Do While wkFeuille.Cells(iLigne, 3).Value <> ""
            If wkFeuille.Cells(iLigne, 2).Value <> "" Then
                .AddItem wkFeuille.Cells(iLigne, 2).Value
            End If
            'Passe à la ligne suivante
            iLigne = iLigne + 1
        Loop
    End With

    'Compléte la ListBox avec les priorités
    With Me.LST_PRIORITE
        .AddItem "Journée"
        .AddItem "Semaine"
        .AddItem "Mois"
    End With

    'Initialise le champ durée
    Me.SAI_DUREE.Value = "00:00"
End Sub

Private Sub BTN_CANCEL_Click()
    'Masque l'USerForm
    Me.Hide
End Sub

Private Sub BTN_OK_Click()
    'Déclaration des variables
    Dim bProjet As Boolean
    Dim iLigne As Integer
    
    bProjet = False
    iLigne = 4
    
    'Vérifie que les champs soient complétés
    If Me.LST_PROJET.Value = "" Then
        'Affiche un message d'alerte à l'utilisateur
        MsgBox ("Veuillez saisir un projet pour ajouter une tâche.")
    ElseIf Me.SAI_TACHE.Value = "" Then
        'Affiche un message d'alerte à l'utilisateur
        MsgBox ("Veuillez saisir une tâche.")
    ElseIf Me.SAI_DUREE.Value = "00:00" Then
        'Affiche un message d'alerte à l'utilisateur
        MsgBox ("Veuillez saisir une durée provisoire pour créer une tâche.")
    ElseIf Me.LST_PRIORITE.Value = "" Then
        'Affiche un message d'alerte à l'utilisateur
        MsgBox ("Veuillez saisir une priorité pour créer une tâche")
    Else
        'Recherche le projet dans la liste de ceux disponible
        Do While ActiveSheet.Cells(iLigne, 3).Value <> ""
            If ActiveSheet.Cells(iLigne, 2).Value = Me.LST_PROJET.Value Then
                bProjet = True
            End If
            'Passe à la ligne suivante
            iLigne = iLigne + 1
        Loop
        'Si le projet n'existe pas on le crée
        If Not bProjet Then
            'Appelle la fonction pour ajouter un projet
            Call AjoutProjet(ActiveSheet.Name, Me.LST_PROJET.Value, Me.SAI_TACHE.Value, Me.LST_PRIORITE.Value, Me.SAI_DUREE.Value)
        Else
            'Sinon on crée la tâche
            Call AjoutTaches(ActiveSheet.Name, Me.LST_PROJET.Value, Me.SAI_TACHE.Value, Me.LST_PRIORITE.Value, Me.SAI_DUREE.Value)
        End If
        'Masque l'USerForm
        Me.Hide
    End If
End Sub

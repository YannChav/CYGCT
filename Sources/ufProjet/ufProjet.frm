VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProjet 
   Caption         =   "Ajout d'un projet"
   ClientHeight    =   3168
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5064
   OleObjectBlob   =   "ufProjet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProjet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_CANCEL_Click()
    'Masque l'USerForm
    Me.Hide
End Sub

Private Sub BTN_OK_Click()
    'V�rifie que les champs soient compl�t�s
    If Me.SAI_PROJET.Value = "" Then
        'Affiche un message d'alerte � l'utilisateur
        MsgBox ("Veuillez saisir un nom pour le projet")
    ElseIf Me.SAI_TACHE.Value = "" Then
        'Affiche un message d'alerte � l'utilisateur
        MsgBox ("Veuillez saisir une t�che pour cr�er un projet")
    ElseIf Me.SAI_DUREE.Value = "00:00" Then
        'Affiche un message d'alerte � l'utilisateur
        MsgBox ("Veuillez saisir une dur�e provisoire pour cr�er une t�che")
    ElseIf Me.LST_PRIORITE.Value = "" Then
        'Affiche un message d'alerte � l'utilisateur
        MsgBox ("Veuillez saisir une priorit� pour cr�er une t�che")
    Else
        'Appelle la fonction pour ajouter un projet
        Call AjoutProjet(ActiveSheet.Name, Me.SAI_PROJET.Value, Me.SAI_TACHE.Value, Me.LST_PRIORITE.Value, Me.SAI_DUREE.Value)
        'Masque l'USerForm
        Me.Hide
    End If
End Sub

Private Sub UserForm_Initialize()
    'Initialise le champ dur�e
    Me.SAI_DUREE.Value = "00:00"
    
    'Compl�te la ListBox avec les priorit�s
    With Me.LST_PRIORITE
        .AddItem "Journ�e"
        .AddItem "Semaine"
        .AddItem "Mois"
    End With
End Sub

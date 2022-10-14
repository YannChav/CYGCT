VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNewPlanning 
   Caption         =   "Nouvelle Semaine"
   ClientHeight    =   1680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4380
   OleObjectBlob   =   "ufNewPlanning.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNewPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_CANCEL_Click()
    'Masque l'USerForm
    Me.Hide
End Sub

Private Sub BTN_OK_Click()
    'Appelle la fonction pour dupliquer la feuille
    If Me.LST_FEUILLES.Value = "Vierge" Then
        Call AjouterPlanning("MODELE SEMAINE")
    Else
        Call AjouterPlanning(Me.LST_FEUILLES.Value)
    End If
    'Masque l'USerForm
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    'Initialise la liste des feuilles disponibles
    Dim Feuille As Worksheet
    
    'Ajoute le modèle de planning
    Me.LST_FEUILLES.AddItem "Vierge"
    'Ajoute chaque feuille de planning à la liste
    For Each Feuille In Sheets
        'Vérifie que la feuille est une feuille de planning
        If Len(Feuille.Name) > 8 Then
            If Left(Feuille.Name, 8) = "Semaine " Then
                'Ajoute la feuille
                Me.LST_FEUILLES.AddItem Feuille.Name
            End If
        End If
    Next Feuille
    
    'Met par défaut l'option Vierge
    Me.LST_FEUILLES.Value = "Vierge"
End Sub

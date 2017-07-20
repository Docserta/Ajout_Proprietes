VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_SelFic 
   Caption         =   "Choix des paramètres"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   OleObjectBlob   =   "FRM_SelFic.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_SelFic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Btn_Annule_Click()
    Me.CB_Annule = True
    Me.Hide
End Sub

Private Sub Btn_OK_Click()

Me.Hide
End Sub

Private Sub Btn_parcourir_Click()
'Recupere le fichier de Nomenclature
Dim NomComplet As String
    
    'Recherche du fichier de nomenclature
    NomComplet = CATIA.FileSelectionBox("Sélection du fichier de nomenclature", "*.xls", CatFileSelectionModeOpen)
    If NomComplet = "" Then Exit Sub 'on vérifie que qque chose a bien été selectionné
    Me.Tbx_FicNom = NomComplet


End Sub


Private Sub Btn_parcourirParam_Click()
'Recupere le fichier de paramètre
'Dim NomComplet As String

    'Ouverture du fichier de paramètres
    pubPathEnv = CATIA.FileSelectionBox("Selectionner le fichier de paramètres", "*.txt", CatFileSelectionModeOpen)
    If pubPathEnv = "" Then Exit Sub 'on vérifie que qque chose a bien été selectionné

    GetAttributs pubPathEnv
    Me.FicNom = pubNomEnv
    Me.Lbl_NomFic = pubPathEnv
    'Sauvegarde du nom du fichier dans les préférences
    If Me.FicNom <> "" Then
        SauvPref
    End If

End Sub

Private Sub Logo_eXcent_Click()

'Chargement de la boite eXcent
    Load Frm_eXcent
     Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    Unload Frm_eXcent
End Sub

Private Sub UserForm_Initialize()
Dim NomPref As String

    Me.CB_Annule = False
    NomPref = RestPref
    'charge les attributs définis dans le fichier de préférences
    If NomPref <> "" Then
        GetAttributs NomPref
    End If
    IniColTemplateOrdo pubNomTemplateOrdo
    Me.FicNom = pubNomEnv
    Me.Lbl_NomFic = NomPref
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        End
    End If
End Sub

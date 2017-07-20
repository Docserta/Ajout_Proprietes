VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Proprietes 
   Caption         =   "Choix des paramètres"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   OleObjectBlob   =   "FRM_Proprietes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Proprietes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CbxHeight As Double = 15
Private Const CbxLeft  As Double = 156
Private Const CbxTop As Double = 144
Private Const CbxWidth  As Double = 13.5
Private Const TbxHeight As Double = 18
Private Const TbxLeft  As Double = 174
Private Const TbxTop As Double = 144
Private Const TbxWidth  As Double = 114
Private Const LeftDecale As Double = 150
Private Const TopDecale As Double = 24

Private Sub Btn_Annule_Click()
    Me.CB_Annule = True
    Me.Hide
End Sub

Private Sub Btn_OK_Click()
'recupère le contenu de la collection des controles du formulaire
Dim fControls As Controls
Dim fControltx As Control
Dim i As Integer

Set fControls = Me.Controls
'Collecte la valeur des textBox
    For i = 1 To pubAttributs.Count
        Set fControltx = fControls.Item("Tbx" & i)
    Next i

Me.Hide
End Sub

Private Sub Btn_parcourir_Click()
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
    'Ajoute les controls pour chaque paramètres
    ConfigForm

   
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Tbx_Version = VMacro
    Frm_eXcent.Show
    Unload Frm_eXcent
End Sub

Private Sub UserForm_Initialize()
'Ajoute les txtbox des paramètres de l'environnement si celui ci est enregistré dans le fichier de préférences
Dim NomPref As String

    Me.CB_Annule = False
    NomPref = RestPref
    'charge les attributs définis dans le fichier de préférences
    If NomPref <> "" Then
        GetAttributs NomPref
        'Ajoute les controls pour chaque paramètres
        ConfigForm
    End If
    Me.FicNom = pubNomEnv
    Me.Lbl_NomFic = NomPref
       
End Sub
Private Sub FRM_Proprietes_QueryClose(Cancel As Integer, CloseMode As Integer)
'If CloseMode = 0 Then Cancel = True
    If CloseMode = 0 Then
        Me.CB_Annule = True
        Me.Hide
    End If
End Sub

Private Sub ConfigForm()
'Redimentionne le formulaire
'Ajoute les champs des atributs
'utilise les variables publiques
'pub
'pubNomEnv
'cibleNomProp

Dim mControls As Controls
Dim cTbxBox As Control
Dim TbxLeft1 As Double
Dim TbxTop1 As Double
Dim oattribut As c_Attribut
Dim i As Long
    
    Set oattribut = New c_Attribut
    Set mControls = Me.Controls
    TbxLeft1 = TbxLeft
    TbxTop1 = TbxTop
    
    'vide la collection oAttribut
'    For i = pubAttributs.Count To 1 Step -1
'        pubAttributs.Remove i
'    Next i
    'Supprime les controles des attributs
    'Cas ou on sélectionne une seconde fois un fichier de paramètres
    For Each cTbxBox In Me.Controls
        If Left(cTbxBox.Name, 3) = "Tbx" Then Me.Controls.Remove cTbxBox.Name
    Next
    
'    Me.FicNom = pubNomEnv
'    Me.Lbl_NomFic = cibleNomProp
    
    Nbre_Param = 1
    If pubAttributs.Count > 0 Then
        For Each oattribut In pubAttributs.Items
            'Calcul de la position des controles (ils se répartissent sur 2 colonnes
            If Nbre_Param Mod 2 = 0 Then
                TbxLeft1 = TbxLeft1 + LeftDecale
            Else
                TbxLeft1 = TbxLeft1 - LeftDecale
                TbxTop1 = TbxTop1 + TopDecale
            End If
                  
            Set cTbxBox = Me.Controls.Add("forms.TextBox.1", "Tbx" & Nbre_Param, True)
            With cTbxBox
                .Height = TbxHeight
                .Left = TbxLeft1
                .Top = TbxTop1
                .Width = TbxWidth
                .Value = oattribut.Nom
            End With
            cTbxBox.Locked = True
            Nbre_Param = Nbre_Param + 1
        Next
        Me.Height = TbxTop1 + 50
        Set mControls = Nothing
        Nbre_Param = Nbre_Param - 1
    End If
    Set oattribut = Nothing
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        End
    End If
End Sub

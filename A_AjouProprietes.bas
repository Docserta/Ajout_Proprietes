Attribute VB_Name = "A_AjouProprietes"
Option Explicit
Sub CATMain()

' *****************************************************************
' * Création d'attributs sur les parts et products
' * Lance une boite de dialogue permettant de choisir les attributs a ajouter
' * puis ajoute les attibuts sur chaques parts et products
' * Création CFR le 21/10/2016
' * modification le : 12/11/16
' *     Collecte des Produits dans la classe c_produits (les groupes sont des produits mais ne font pas parti de la collection des Documents
' *
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "A_AjouProprietes", VMacro

Dim mdocs As Documents
Dim mdoc As Document
Dim mprod As Product
Dim mparams As Parameters
Dim mbarre As ProgressBarre
Dim indP As Long 'indice de progression dans la barre
Dim oattribut As c_Attribut

    'Configure les champs en fonction de la langue
    InitLanguage
    
'Initialisation des classe
    Set oattribut = New c_Attribut
    'Set pubAttributs = New c_Attributs
    Set oProduits = New c_Products

    Set mdocs = CATIA.Documents
    
'Test si un CatProduct est actif
    If check_Env("Product") Then Set mdoc = CATIA.ActiveDocument
       
'Test si le product général est vide
    If mdocs.Count = 0 Then
        MsgBox "Ce product est vide !", vbCritical, "Erreur"
        Exit Sub
    End If


'Chargement du formulaire
    Load FRM_Proprietes
    FRM_Proprietes.Show
    'Bouton "annuler" choisi, on decharge le formulaire et on quite
    If FRM_Proprietes.CB_Annule Then
        Unload FRM_Proprietes
        Exit Sub
    End If
    'GetAttributs FRM_Proprietes.FicNom
    Unload FRM_Proprietes
    
'Chargement de la barre de progression
    Set mbarre = New ProgressBarre
    mbarre.Titre = "Création des paramètres"
    mbarre.Progression = 1
    mbarre.Affiche

'Chargement en mode conception
    mdoc.Product.ApplyWorkMode DESIGN_MODE

'Collecte des Produits (les groupes sont des produits mais ne font pas parti de la collection des Documents)
    On Error Resume Next
    For Each mdoc In mdocs
        On Error Resume Next
        Set mprod = mdoc.Product
        If mprod Is Nothing Then
            err.Clear
            On Error GoTo 0
        Else
            'collecte du product du doc en cours
            Set oProduit = New c_Product
            On Error Resume Next
            Set oProduit = oProduits.Item(mprod.PartNumber) 'Recherche si le produit est déja présent dans la collection (evite les doublons)
            If err.Number <> 0 Then
                oProduit.Nom = mprod.PartNumber
                oProduit.Produit = mprod
                oProduit.Parent = mprod.PartNumber
                oProduits.Add oProduit.Nom, oProduit.Produit, oProduit.Parent
            End If
            Set oProduit = Nothing
            
            
            'Collecte des sous produits du product en cours
            For Each mprod In mdoc.Product.Products
                If CompIsactive(mprod, mdoc.Product.Name) Then
                    Set oProduit = New c_Product
                    On Error Resume Next
                    Set oProduit = oProduits.Item(mprod.PartNumber) 'Recherche si le produit est déja présent dans la collection (evite les doublons)
                    If err.Number <> 0 Then
                        oProduit.Nom = mprod.PartNumber
                        oProduit.Produit = mprod
                        oProduit.Parent = mprod.PartNumber
                        oProduits.Add oProduit.Nom, oProduit.Produit, oProduit.Parent
                    End If
                    Set oProduit = Nothing
                End If
            Next
        End If
    Next
'Ajoute les attributs
    Set oProduit = New c_Product
    For Each oProduit In oProduits.Items
        mbarre.Progression = 100 / oProduits.Count * indP
        On Error Resume Next
        Set mparams = oProduit.Produit.ReferenceProduct.UserRefProperties
        If Not (mparams Is Nothing) Then
'        'If Err.Number <> 0 Then
'            'Err.Clear
            For Each oattribut In pubAttributs.Items
                'Création D 'un paramètre vide s'il n'existe pas déjà
                CreateParam mparams, oattribut.Nom, ""
            Next
        End If
        indP = indP + 1
    Next
    
    mbarre.Cache
   
MsgBox "Fin d'import des propriétées dans Catia", vbInformation, "Fin de traitement"
   
'Libération des classes
Set mbarre = Nothing
Set oattribut = Nothing
Set pubAttributs = Nothing
Set oProduit = Nothing
Set oProduits = Nothing

End Sub



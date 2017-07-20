Attribute VB_Name = "C_ModifProprietes"
Option Explicit

Sub CATMain()
' *****************************************************************
' * Remonte les infos du fichier excel modifié par l'utilisateur
' * et met à jour les attributs des parts et products
' *
' * Création CFR le 21/10/2016
' * modification le : 27/03/17
' *                 Ajout d'un fonction SupZero supprimant les zéro non significatifs des part Number
' *
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "C_ModifProprietes", VMacro

Dim mdocs As Documents
Dim mdoc As Document
Dim mDocFils As Document
Dim mprod As Product
Dim mparams As Parameters
Dim oLigNom As c_LNomencl
Dim oLigNoms As c_LNomencls
Dim mbarre As ProgressBarre
Dim pEtape As Long, pNbEt As Long, pItem As Long, pItems As Long
Dim oAttributs As c_Attributs
Dim oattribut As c_Attribut

'Initialisation des classes
    Set oattribut = New c_Attribut
    Set oAttributs = New c_Attributs
    Set oLigNom = New c_LNomencl
    Set oLigNoms = New c_LNomencls
    Set oProduits = New c_Products
    
    Set mdocs = CATIA.Documents
    
'Configure les champs en fonction de la langue
    InitLanguage
    
'Test si un CatProduct est actif
    If check_Env("Product") Then Set mdoc = CATIA.ActiveDocument

'Chargement du formulaire
    Load FRM_SelFic
    FRM_SelFic.Lbl_TypFicNom = "Sélectionnez le fichier des attributs modifiées"
    FRM_SelFic.Lbl_NomFicNom = "Nom du fichier des attributs modifiées"
    FRM_SelFic.Show
    'Bouton "annuler" choisi, on decharge le formulaire et on quite
    If FRM_SelFic.CB_Annule Then
        Unload FRM_SelFic
        End
    End If
    cibleNomCatia = FRM_SelFic.Tbx_FicNom
    Unload FRM_SelFic
    
'Chargement de la barre de progression
    Set mbarre = New ProgressBarre
    mbarre.Titre = "Import des valeurs modifièes des paramètres"
    mbarre.Progression = 1
    mbarre.Affiche
    pNbEt = 3: pEtape = 1: pItem = 1: pItems = 1
    
'Chargement en mode conception
    mdoc.Product.ApplyWorkMode DESIGN_MODE
    
'Collecte de la valeur des attributs dans le fichier eXcel
    Set oLigNoms = GetBomXl(cibleNomCatia, mbarre)

'Collecte des Produits (les groupes sont des produits mais ne font pas parti de la collection des Documents)
    For Each mdoc In mdocs
        On Error Resume Next
        Set mprod = mdoc.Product
        If mprod Is Nothing Then 'Saute les doc parasites tel que les catalogues
            err.Clear
            On Error GoTo 0
        Else
            'Ajoute le product en cours
            Set oProduit = New c_Product
                oProduit.Nom = mprod.PartNumber
                oProduit.Produit = mprod
                oProduit.Parent = mprod.PartNumber
                oProduits.Add oProduit.Nom, oProduit.Produit, oProduit.Parent
            Set oProduit = Nothing
            'Ajoute les sous-products du product en cours
            For Each mprod In mdoc.Product.Products
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
            Next
        End If
    Next
    
    pEtape = 3: pItem = 1: pItems = oProduits.Count
    
    
'Mise à jour des attributs dans les products et les parts
    Set oProduit = New c_Product
    For Each oProduit In oProduits.Items
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Mise à jour des attributs dans les products et les parts"
        On Error Resume Next
        Set mparams = oProduit.Produit.ReferenceProduct.UserRefProperties
        If Not (mparams Is Nothing) Then
            On Error Resume Next
            
            Set oLigNom = oLigNoms.Item(SupZero(oProduit.Nom))
                If err.Number = 0 Then 'si le composants catia n'est pas référencé dans la nomenclature
                    Set oAttributs = oLigNom.Attributs
                    'traitement du paramètre "Revision"
                    oProduit.Produit.Revision = oLigNom.Rev
                    'traitement du paramètre "Definition"
                    oProduit.Produit.Definition = oLigNom.Def
                    'traitement du paramètre "Nomenclature"
                    oProduit.Produit.Nomenclature = oLigNom.Nom
                     'traitement du paramètre "source"
                    oProduit.Produit.Source = ReverseSource(oLigNom.Source)
                    'traitement du paramètre "DescriptionRef"
                    oProduit.Produit.DescriptionRef = oLigNom.Desc
                    'Les attributs personalisés
                    For Each oattribut In oAttributs.Items
                        'Création D 'un paramètre vide s'il n'existe pas déjà
                        CreateParamExist mparams, oattribut.Nom, oattribut.Valeur
                    Next
                End If
        Else
            err.Clear
            On Error GoTo 0
        End If
    Next
    pItem = pItem + 1
    mbarre.Cache

MsgBox "Import de attributs terminé", vbInformation, "Fin d'import"

'Libération des classes
Set mbarre = Nothing

End Sub

Public Function GetBomXl(cibleNomCatia As String, mbarre) As c_LNomencls
'Renvoi la collection des attributs des ensembles et des parts
'Ouvre le fichier excel de la nomenclature regroupant les ensembles et les part

Dim objexcel
Dim objWBk
Dim LigActive As Long
Dim ColActive As Integer
Dim NoLigEntete As Long 'N° de la ligne des entète d'attributs dans le fichier eXcel
Dim NoLigFinNom As Long 'N° de la ligne de fin de la nomenclature
Dim NoLigFinEns As Long 'N° de la ligne de fin des ensembles
Dim NoLigDebDet As Long 'N° de la ligne de début des pièces
Dim NoLigDebSSE As Long 'N° de la ligne de début des ensembles
Dim oLigNom As c_LNomencl
Dim oLigNoms As c_LNomencls
Dim pEtape As Long, pNbEt As Long, pItem As Long, pItems As Long
Dim oattribut As c_Attribut
Dim oAttributs As c_Attributs
Dim oAttributEnv As c_Attribut
Dim oAttributEnvs As c_Attributs

    'Initialisation de l'objet Excel
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWBk = objexcel.Workbooks.Open(CStr(cibleNomCatia))
    objexcel.Visible = False
    'Détection de l'onglet "recapitulatif"
    On Error Resume Next
    objWBk.Sheets(nSheetReacp).Activate
    If err.Number <> 0 Then
        MsgBox "L'onglet " & nSheetReacp & " n'a pas été trouvé dans le fichier Excel", vbCritical, "Fichier incorrect"
        err.Clear
        GoTo erreur
    End If
    
    'Recherche la position des lignes dans le fichier Excel
    NoLigEntete = 3
    NoLigFinNom = NoDerniereLigne(objWBk)
    NoLigDebDet = NoDebLstPieces(objWBk) + 2
    NoLigFinEns = NoLigDebDet - 2
    NoLigDebSSE = 4

'initialisation des classes
    Set oLigNom = New c_LNomencl
    Set oLigNoms = New c_LNomencls
    Set oAttributEnv = New c_Attribut
    Set oAttributEnvs = New c_Attributs
    Set oattribut = New c_Attribut
    Set oAttributs = New c_Attributs

'Collecte de la liste des attributs dans les cellules de la ligne d'entète de la nomenclature
    'Collecte les attributs standards (présent sur tous les pats et products
        'Révision, Definition, Nomenclature, source et product description)
            'puis les attributs personalisés (ceux du fichier texte des propriètes)
    ColActive = NbColPrmStd + 1 'On saute les champs Qte, Référence qui ne sont pas modifiables
    While objWBk.ActiveSheet.cells(NoLigEntete, ColActive) <> ""
        oAttributEnv.Nom = objWBk.ActiveSheet.cells(NoLigEntete, ColActive)
        oAttributEnvs.Add oAttributEnv.Nom, False
        ColActive = ColActive + 1
    Wend

'Collecte de la valeur des attributs des ensembles
    LigActive = NoLigDebSSE
    pEtape = 1: pItem = 1: pItems = NoLigFinEns - LigActive
    While LigActive < NoLigFinEns
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Collecte de la valeur des attributs des ensembles"
        Set oAttributs = New c_Attributs
        ColActive = NbColPrmNomModif  'On saute le champs "Qte" qui n'est pas modifiable
        oLigNom.Ref = objWBk.ActiveSheet.cells(LigActive, ColActive) 'on garde le champs ref qui sert d'index
        oLigNom.Comp = "E"
        ColActive = ColActive + 1
        ' Révision
        oLigNom.Rev = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ColActive = ColActive + 1
        ' Definition
        oLigNom.Def = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ColActive = ColActive + 1
        ' Nomenclature
        oLigNom.Nom = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ColActive = ColActive + 1
        ' source
        oLigNom.Source = objWBk.ActiveSheet.cells(LigActive, ColActive) 'FormatSource(objWBk.ActiveSheet.cells(LigActive, ColActive))
        ColActive = ColActive + 1
        ' product description
        oLigNom.Desc = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ' Les attributs personalisables
        ColActive = NbColPrmStd + 1
        For Each oAttributEnv In oAttributEnvs.Items
            oattribut.Nom = oAttributEnv.Nom
            oattribut.Valeur = objWBk.ActiveSheet.cells(LigActive, ColActive)
            ColActive = ColActive + 1
            oAttributs.Add oattribut.Nom, False, oattribut.Valeur
        Next
        LigActive = LigActive + 1
        oLigNom.Attributs = oAttributs
        Set oAttributs = Nothing
        oLigNoms.Add oLigNom.Ref, oLigNom.Comp, oLigNom.Qte, oLigNom.Rev, oLigNom.Def, oLigNom.Nom, oLigNom.Source, oLigNom.Desc, oLigNom.Attributs
        pItem = pItem + 1
    Wend
    
'collecte de la valeur des attributs des pièces
    LigActive = NoLigDebDet
    pEtape = 2: pItem = 1: pItems = NoLigFinNom - LigActive
    While LigActive < NoLigFinNom
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "collecte de la valeur des attributs des pièces"
        Set oAttributs = New c_Attributs
        ColActive = NbColPrmNomModif 'On saute le champs "Qte" qui n'est pas modifiable
        oLigNom.Ref = objWBk.ActiveSheet.cells(LigActive, ColActive) 'on garde le champs ref qui sert d'index
        oLigNom.Comp = "D"
        ColActive = ColActive + 1
        ' Révision
        oLigNom.Rev = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ColActive = ColActive + 1
        ' Definition
        oLigNom.Def = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ColActive = ColActive + 1
        ' Nomenclature
        oLigNom.Nom = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ColActive = ColActive + 1
        ' source
        oLigNom.Source = objWBk.ActiveSheet.cells(LigActive, ColActive) 'FormatSource(objWBk.ActiveSheet.cells(LigActive, ColActive))
        ColActive = ColActive + 1
        ' product description
        oLigNom.Desc = objWBk.ActiveSheet.cells(LigActive, ColActive)
        ' Les attributs personalisables
        ColActive = NbColPrmStd + 1
        For Each oAttributEnv In oAttributEnvs.Items
            oattribut.Nom = oAttributEnv.Nom
            oattribut.Valeur = objWBk.ActiveSheet.cells(LigActive, ColActive)
            ColActive = ColActive + 1
            oAttributs.Add oattribut.Nom, False, oattribut.Valeur
        Next
        LigActive = LigActive + 1
        oLigNom.Attributs = oAttributs
        Set oAttributs = Nothing
        oLigNoms.Add oLigNom.Ref, oLigNom.Comp, oLigNom.Qte, oLigNom.Rev, oLigNom.Def, oLigNom.Nom, oLigNom.Source, oLigNom.Desc, oLigNom.Attributs
        pItem = pItem + 1
    Wend

Set GetBomXl = oLigNoms

erreur:
'Libération des classes
    Set oLigNoms = Nothing
    Set oLigNom = Nothing
    Set oAttributEnv = Nothing
    Set oAttributEnvs = Nothing
    Set oattribut = Nothing
    Set oAttributs = Nothing
End Function


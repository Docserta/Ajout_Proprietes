Attribute VB_Name = "B_ExtractNom"
 Option Explicit
    
Sub CATMain()

' *****************************************************************
' * Extraction de la nomenclature vers un fichier excel
' * Crée un onglet regroupant les sous ensembles et les parts de détails
' * Récupère les attributs et leur valeur
' * Création CFR le 21/10/2016
' * modification le :
' *
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "B_ExtractNom", VMacro

Dim mdoc As ProductDocument
Dim mprod As Product
Dim ExtracDescr As Boolean
Dim Reponse As Integer
Dim i As Integer, No As Integer
Dim oattribut As c_Attribut

    'Configure les champs en fonction de la langue
    InitLanguage

    'Test si un CatProduct est actif
    If check_Env("Product") Then Set mdoc = CATIA.ActiveDocument
    
    Set oattribut = New c_Attribut
    Set pubAttributs = New c_Attributs
    'Chargement du formulaire
    Load FRM_Proprietes
    FRM_Proprietes.Show
    'Bouton "annuler" choisi, on decharge le formulaire et on quite
    If FRM_Proprietes.CB_Annule Then
        Unload FRM_Proprietes
        Exit Sub
    End If
    GetAttributs FRM_Proprietes.Lbl_NomFic
    Unload FRM_Proprietes
      
    'Extraction de la nomenclature du product général et sauvegarde dans un fichier excel
    GenBomCatia mdoc, pubAttributs
    'Export des sous ensembles et des détails dans le fichier excel des propriétés modifiables
    PutBomXl cibleNomCatia
          
'Libération des classes
Set oattribut = Nothing
Set mprod = Nothing
Set mdoc = Nothing

End Sub

Public Sub PutBomXl(cibleNomCatia As String)
'Formate le fichier excel de la nomenclature généré par Catia
'Regoupe les ensembles
'Regroupe les détails

Dim mdoc As ProductDocument
Dim mparams As Parameters
Dim objexcel
Dim objWBkNomCatia
Dim objWBkProp
Dim LigActive As Long
Dim ColActive As Integer
Dim NoLigFinNom As Long
Dim NoLigFinEns As Long
Dim NoLigDebDet As Long
Dim NoLigDebSSE As Long
Dim NomSSE As String
Dim cLigNomDet As c_LNomencl
Dim cLigNomDets As c_LNomencls
Dim cLigNomEn As c_LNomencl
Dim cLigNomENs As c_LNomencls
Dim cAttribut As c_Attribut
Dim cAttributs As c_Attributs
Dim cAttributEnv As c_Attribut
Dim cAttributEnvs As c_Attributs 'Collections des attributs de l'environnement (ex GSE)
Dim i As Long
Dim pEtape As Long, pNbEt As Long, pItem As Long, pItems As Long
Dim mbarre As ProgressBarre
Dim pos As Integer

Set mdoc = CATIA.ActiveDocument

'Initialisation des classes
    Set cAttributEnv = New c_Attribut
    Set cAttributEnvs = New c_Attributs
    Set cLigNomEn = New c_LNomencl
    Set cLigNomENs = New c_LNomencls
    
'Chargement de la barre de progression
    Set mbarre = New ProgressBarre
    mbarre.Titre = "Extraction des paramètres"
    mbarre.Progression = 1
    mbarre.Affiche
    pNbEt = 5: pEtape = 1: pItem = 1: pItems = 1

'Initialisation des objets Excel
    'Ouverture de la nom générée par Catia
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWBkNomCatia = objexcel.Workbooks.Open(CStr(cibleNomCatia))
    objexcel.Visible = False
    'Création de la nom des propriétés modifiables
    nNomProp = Left(mdoc.Name, InStr(1, mdoc.Name, ".CATProduct", vbTextCompare) - 1) & nExtNomProp & ".xls"
    cibleNomProp = mdoc.Path & "\" & nNomProp
    Set objWBkProp = objexcel.Workbooks.Add()

    'Recherche la position des lignes dans le fichier Excel
    NoLigFinNom = NoDerniereLigne(objWBkNomCatia)
    NoLigFinEns = NoDebRecap(objWBkNomCatia)
    NoLigDebDet = NoLigFinEns + 4
    NoLigDebSSE = 3
    
    'collecte de la liste des propriétés spécifiques à l'environnement (nom des attributs)
    LigActive = 4
    ColActive = NbColPrmStd + 1 'Première colonne après les attributs standards
    While objWBkNomCatia.ActiveSheet.cells(LigActive, ColActive).Value <> ""
        cAttributEnv.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, ColActive).Value
        cAttributEnv.Ordre = ColActive
        cAttributEnvs.Add cAttributEnv.Nom, cAttributEnv.Ordre
        ColActive = ColActive + 1
    Wend
           
    'Collecte des sous ensembles
    LigActive = 5

    'Creation de la liste des SSe (reference)
    pEtape = 1: pItem = 1: pItems = NoLigFinEns - LigActive
    While LigActive < NoLigFinEns
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Création de la liste des sous ensembles"
        NomSSE = TestEstSSE(objWBkNomCatia.ActiveSheet.cells(LigActive, 1).Value)
        If NomSSE <> "False" Then
            cLigNomEn.Ref = NomSSE
            'Ajout du sous ensemble a la collection
            cLigNomENs.Add cLigNomEn.Ref
        End If
        LigActive = LigActive + 1
        pItem = pItem + 1
    Wend
    
'collecte des propriétés de chaque SSE
    LigActive = 5
    pEtape = 2: pItem = 1: pItems = NoLigFinEns - LigActive
    While LigActive < NoLigFinEns
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "collecte des propriétés de chaque sous ensemble"
        For Each cLigNomEn In cLigNomENs.Items
            If objWBkNomCatia.ActiveSheet.cells(LigActive, 2).Value = cLigNomEn.Ref Then
                Set cAttributs = New c_Attributs
                Set cAttribut = New c_Attribut
                'Quantité
                pos = 1 'cAttributEnvs.Item(nQt).Ordre
                cLigNomEn.Qte = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' Révision
                pos = 3 ' cAttributEnvs.Item(nRev).Ordre
                cLigNomEn.Rev = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' Definition
                pos = 4 ' cAttributEnvs.Item(nDef).Ordre
                cLigNomEn.Def = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' Nomenclature
                pos = 5 'cAttributEnvs.Item(nNom).Ordre
                cLigNomEn.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' source
                pos = 6 'cAttributEnvs.Item(nSrce).Ordre
                cLigNomEn.Source = FormatSource(objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value)
                ' product description
                pos = 7 'cAttributEnvs.Item(nDesc).Ordre
                cLigNomEn.Desc = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                cLigNomEn.Comp = "E"
                'collecte des attributs liés a l'environnement
                For Each cAttributEnv In cAttributEnvs.Items
                    cAttribut.Nom = cAttributEnv.Nom
                    cAttribut.Ordre = cAttributEnv.Ordre
                    cAttribut.Valeur = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
                    cAttributs.Add cAttribut.Nom, cAttribut.Ordre, cAttribut.Valeur
                Next
                'Ajout de la collection des attributs a la ligne de nomenclature
                cLigNomEn.Attributs = cAttributs
                'vidage de la collection des attributs
                Set cAttribut = Nothing
            End If
        Next
        LigActive = LigActive + 1
        pItem = pItem + 1
    Wend
    Set cLigNomEn = Nothing
    
'collecte du product de tète
    Set cLigNomEn = New c_LNomencl
    Set cAttributs = New c_Attributs
    Set cAttribut = New c_Attribut
    'Part Number
    cLigNomEn.Ref = mdoc.Product.Name
    ' Quantité
    cLigNomEn.Qte = 1
    ' Révision
    cLigNomEn.Rev = mdoc.Product.Revision
    ' Definition
    cLigNomEn.Def = mdoc.Product.Definition
    ' Nomenclature
    cLigNomEn.Nom = mdoc.Product.Nomenclature
    ' source
    cLigNomEn.Source = FormatSource(mdoc.Product.Source)
    ' product description
    cLigNomEn.Desc = mdoc.Product.DescriptionRef
    cLigNomEn.Comp = "E"
    'collecte des attributs du product de tète liés a l'environnement
    On Error Resume Next
    Set mparams = mdoc.Product.UserRefProperties
    If err.Number = 0 Then
        Set cAttributs = New c_Attributs
        Set cAttribut = New c_Attribut
        For Each cAttributEnv In cAttributEnvs.Items
            cAttribut.Nom = cAttributEnv.Nom
            cAttribut.Ordre = cAttributEnv.Ordre
            cAttribut.Valeur = TestParamExist(mparams, cAttributEnv.Nom)
            cAttributs.Add cAttribut.Nom, cAttribut.Ordre, cAttribut.Valeur
        Next
    Else
        err.Clear
        On Error GoTo 0
    End If
    
    'Ajout de la collection des attributs a la ligne de nomenclature
    cLigNomEn.Attributs = cAttributs
    'Ajout du sous ensemble a la collection
    cLigNomENs.Add cLigNomEn.Ref, cLigNomEn.Comp, cLigNomEn.Qte, cLigNomEn.Rev, cLigNomEn.Def, cLigNomEn.Nom, cLigNomEn.Source, cLigNomEn.Desc, cLigNomEn.Attributs
    'vidage de la collection des attributs
    Set cAttribut = Nothing
    Set cLigNomEn = Nothing
    
'Collecte des détails
    Set cLigNomDet = New c_LNomencl
    Set cLigNomDets = New c_LNomencls
    LigActive = NoLigDebDet + 1
    
    pEtape = 3: pItem = 1: pItems = NoLigFinNom - LigActive
    While LigActive < NoLigFinNom
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Collecte des propriétés des détails"
        Set cAttributs = New c_Attributs
        Set cAttribut = New c_Attribut
        'Collecte de la valeur des attributs Standards
        'Part Number
        'Set cAttributEnv = cAttributEnvs.Item(nRef)
        'cLigNomDet.Ref = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
        cLigNomDet.Ref = objWBkNomCatia.ActiveSheet.cells(LigActive, 2).Value
        ' Quantité
        'Set cAttributEnv = cAttributEnvs.Item(nQt)
        'cLigNomDet.Qte = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
        cLigNomDet.Qte = objWBkNomCatia.ActiveSheet.cells(LigActive, 1).Value
        ' Révision
        'Set cAttributEnv = cAttributEnvs.Item(nRev)
        'cLigNomDet.Rev = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
        cLigNomDet.Rev = objWBkNomCatia.ActiveSheet.cells(LigActive, 3).Value
        ' Definition
        'Set cAttributEnv = cAttributEnvs.Item(nDef)
        'cLigNomDet.Def = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
        cLigNomDet.Def = objWBkNomCatia.ActiveSheet.cells(LigActive, 4).Value
        ' Nomenclature
        'Set cAttributEnv = cAttributEnvs.Item(nNom)
        'cLigNomDet.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
        cLigNomDet.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, 5).Value
        ' source
        'Set cAttributEnv = cAttributEnvs.Item(nSrce)
        'cLigNomDet.Source = FormatSource(objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value)
        cLigNomDet.Source = FormatSource(objWBkNomCatia.ActiveSheet.cells(LigActive, 6).Value)
        ' product description
        'Set cAttributEnv = cAttributEnvs.Item(nDesc)
        'cLigNomDet.Desc = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
        cLigNomDet.Desc = objWBkNomCatia.ActiveSheet.cells(LigActive, 7).Value
        cLigNomDet.Comp = "D"
        'collecte de la valeur des attributs spécifiques a l'environnement
        For Each cAttributEnv In cAttributEnvs.Items
            cAttribut.Nom = cAttributEnv.Nom
            cAttribut.Ordre = cAttributEnv.Ordre
            cAttribut.Valeur = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
            cAttributs.Add cAttribut.Nom, cAttribut.Ordre, cAttribut.Valeur
        Next
        'Ajout de la collection des attributs a la ligne de nomenclature
        cLigNomDet.Attributs = cAttributs
        'vidage de la collection des attributs
        Set cAttribut = Nothing
        LigActive = LigActive + 1
        cLigNomDets.Add cLigNomDet.Ref, cLigNomDet.Comp, cLigNomDet.Qte, cLigNomDet.Rev, cLigNomDet.Def, cLigNomDet.Nom, cLigNomDet.Source, cLigNomDet.Desc, cLigNomDet.Attributs
        pItem = pItem + 1
    Wend

'Copie des infos dans le fichier excel des propriétés
    'renommage de l'onglet récapitulatif
    objWBkProp.Sheets.Item(1).Name = nSheetReacp
    LigActive = 2
    
    'Copie des entètes de colonnes pour les sous ensembles
    objWBkProp.ActiveSheet.cells(LigActive, 1) = "Liste des sous ensembles"
    LigActive = LigActive + 1
    XlEnteteCol objWBkProp.ActiveSheet, cAttributEnvs, LigActive
    LigActive = LigActive + 1
    
    'copie des sous ensembles
    pEtape = 4: pItem = 1: pItems = cLigNomENs.Count
    For Each cLigNomEn In cLigNomENs.Items
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Ecriture des ensembles"
        objWBkProp.ActiveSheet.cells(LigActive, 1) = cLigNomEn.Qte
        objWBkProp.ActiveSheet.cells(LigActive, 2) = cLigNomEn.Ref
        objWBkProp.ActiveSheet.cells(LigActive, 3) = cLigNomEn.Rev
        objWBkProp.ActiveSheet.cells(LigActive, 4) = cLigNomEn.Def
        objWBkProp.ActiveSheet.cells(LigActive, 5) = cLigNomEn.Nom
        objWBkProp.ActiveSheet.cells(LigActive, 6) = cLigNomEn.Source
        objWBkProp.ActiveSheet.cells(LigActive, 7) = cLigNomEn.Desc
        For i = 1 To cLigNomEn.Attributs.Count
            Select Case cLigNomEn.Attributs.Item(i).Nom
                Case nQt, nRef, nRev, nDef, nNom, nSrce, nDesc 'On saute les 4 paramètre non modifiable
                Case Else
                objWBkProp.ActiveSheet.cells(LigActive, NbColPrmStd + i) = cLigNomEn.Attributs.Item(i).Valeur
            End Select
        Next
        LigActive = LigActive + 1
        pItem = pItem = 1
    Next
    LigActive = LigActive + 1
    
    'Copie des entètes de colonnes pour les pièces
    objWBkProp.ActiveSheet.cells(LigActive, 1) = "Liste des pièces"
    LigActive = LigActive + 1
    XlEnteteCol objWBkProp.ActiveSheet, cAttributEnvs, LigActive
    LigActive = LigActive + 1
    
    'copie des détails
    pEtape = 5: pItem = 1: pItems = cLigNomDets.Count
    For Each cLigNomDet In cLigNomDets.Items
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Ecriture des Détails"
        objWBkProp.ActiveSheet.cells(LigActive, 1) = cLigNomDet.Qte
        objWBkProp.ActiveSheet.cells(LigActive, 2) = cLigNomDet.Ref
        objWBkProp.ActiveSheet.cells(LigActive, 3) = cLigNomDet.Rev
        objWBkProp.ActiveSheet.cells(LigActive, 4) = cLigNomDet.Def
        objWBkProp.ActiveSheet.cells(LigActive, 5) = cLigNomDet.Nom
        objWBkProp.ActiveSheet.cells(LigActive, 6) = cLigNomDet.Source
        objWBkProp.ActiveSheet.cells(LigActive, 7) = cLigNomDet.Desc
        
        For i = 1 To cLigNomDet.Attributs.Count
            objWBkProp.ActiveSheet.cells(LigActive, NbColPrmStd + i) = cLigNomDet.Attributs.Item(i).Valeur
        Next
        LigActive = LigActive + 1
        pItem = pItem + 1
    Next
    
    'affichage du fichier de Modification des propriétés et sauvegarde
    objWBkProp.SaveAs (cibleNomProp)
    objexcel.Visible = True
    

'Libération des classes
'Fermeture du fichier excel de nomenclature catia
    objWBkNomCatia.Close
Set cLigNomDet = Nothing
Set cLigNomDets = Nothing
Set cLigNomEn = Nothing
Set cLigNomENs = Nothing
Set cAttributEnv = Nothing
Set cAttributEnvs = Nothing
Set mbarre = Nothing

End Sub


Private Sub XlEnteteCol(ByVal oWbSheet, cAttribs As c_Attributs, Lig As Long)
'Copie les entètes de colonnes dans le fichier excel passé en argumet
'oWbSheet = Feuille excel
'cAttribs = classe des noms d'attributs
'Lig = ligne du fichier excel danslaquelle sont ecrites les entètes
Dim i As Long

Dim Zone As String
    oWbSheet.cells(Lig, 1) = nQt
    oWbSheet.cells(Lig, 2) = nRef
    oWbSheet.cells(Lig, 3) = nRev
    oWbSheet.cells(Lig, 4) = nDef
    oWbSheet.cells(Lig, 5) = nNom
    oWbSheet.cells(Lig, 6) = nSrce
    oWbSheet.cells(Lig, 7) = nDesc
    For i = 1 To cAttribs.Count
        oWbSheet.cells(Lig, cAttribs.Item(i).Ordre) = cAttribs.Item(i).Nom
    Next
    'Mise en forme
    Zone = "A" & Lig & ":" & NumCar(cAttribs.Count + NbColPrmStd) & Lig

    With oWbSheet.Range(Zone)
        .Font.Size = 10
        .Font.Bold = True
        .Interior.Color = 15917714
    End With
    
    
End Sub
Private Sub GenBomCatia(ByVal mdoc As Document, ByVal oAttributs As c_Attributs)
'Génère le fichier Bom par Catia
'mDoc = document dont on veux extraire la nomenclature
'oAttributs collection des attributs

Dim vAssConv
Dim AssConv As AssemblyConvertor
Dim arrayOfVariantOfBSTR1()
Dim arrayOfVariantOfBSTR2()
Dim No As Integer
Dim i As Long
Dim oattribut As New c_Attribut

'verifie si un fichier de nomenclature est déja présent et l'efface
    nNomCatia = Left(mdoc.Name, InStr(1, mdoc.Name, ".CATProduct", vbTextCompare) - 1) & nExtNomCatia & ".xls"
    cibleNomCatia = mdoc.Path & "\" & nNomCatia
    If Not (EffaceFicNom(mdoc.Path, nNomCatia)) Then
        End
    End If
    
'Extraction de la nomenclature du product général et sauvegarde dans un fichier excel
    'Construit la liste des propriétés a extraire
    ReDim arrayOfVariantOfBSTR1(oAttributs.Count + NbColPrmStd)
    ReDim arrayOfVariantOfBSTR2(oAttributs.Count + NbColPrmStd)
    No = 0
    '0=Qté 1=Partnumber 2=Révision 3=Definition 4=Nomenclature 5=source 6=product description
        arrayOfVariantOfBSTR1(No) = nQt
        arrayOfVariantOfBSTR2(No) = nQt
        arrayOfVariantOfBSTR1(No + 1) = nRef
        arrayOfVariantOfBSTR2(No + 1) = nRef
        arrayOfVariantOfBSTR1(No + 2) = nRev
        arrayOfVariantOfBSTR2(No + 2) = nRev
        arrayOfVariantOfBSTR1(No + 3) = nDef
        arrayOfVariantOfBSTR2(No + 3) = nDef
        arrayOfVariantOfBSTR1(No + 4) = nNom
        arrayOfVariantOfBSTR2(No + 4) = nNom
        arrayOfVariantOfBSTR1(No + 5) = nSrce
        arrayOfVariantOfBSTR2(No + 5) = nSrce
        arrayOfVariantOfBSTR1(No + 6) = nDesc
        arrayOfVariantOfBSTR2(No + 6) = nDesc
        No = NbColPrmStd
    'Les autres paramètres ensuite
    For i = 1 To oAttributs.Count
        Set oattribut = oAttributs.Item(i)
        arrayOfVariantOfBSTR1(No) = oattribut.Nom
        arrayOfVariantOfBSTR2(No) = oattribut.Nom
        No = No + 1
    Next

    Set AssConv = mdoc.Product.GetItem("BillOfMaterial")

    Set vAssConv = AssConv
    vAssConv.SetCurrentFormat arrayOfVariantOfBSTR1

    Set vAssConv = AssConv
    vAssConv.SetSecondaryFormat arrayOfVariantOfBSTR2
    AssConv.[Print] "XLS", CStr(cibleNomCatia), mdoc.Product

'Libération des classes
Set AssConv = Nothing
Set vAssConv = Nothing
End Sub







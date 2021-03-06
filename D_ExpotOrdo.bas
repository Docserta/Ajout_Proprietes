Attribute VB_Name = "D_ExpotOrdo"
Option Explicit
Sub CATMain()

' *****************************************************************
' * Export de la nomenclature dans le template Nomenclature
' * Lance une boite de dialogue permettant de choisir les attributs a ajouter
' *
' * Cr�ation CFR le 21/10/2016
' * modification le : 20/12/2016
' *                    Modification de la formule de Qt a cmd des pi�ce directement sous l'assemblage g�n�ral
' *
' *****************************************************************

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "D_ExpotOrdo", VMacro

Dim oattribut As c_Attribut         'attribut personnalis�
Dim oAttributs As c_Attributs       'Collections des attributs personnalis�s
Dim olignomCatias As c_LignomOrdos    'Collection des lignes de nomenclature g�n�r� par catia
Dim mbarre As ProgressBarre
Dim objexcel
Dim objWBk
Dim ColActive As Integer

    'Configure les champs en fonction de la langue
    InitLanguage
    
'initialisation des classes
    Set olignomCatias = New c_LignomOrdos

'Chargement du formulaire
    Load FRM_SelFic
    FRM_SelFic.Lbl_TypFicNom = "S�lectionnez la nomenclature g�n�r�e par Catia"
    FRM_SelFic.Lbl_NomFicNom = "Nom du fichier de la nomenclature Catia"
    FRM_SelFic.Show
    'Bouton "annuler" choisi, on decharge le formulaire et on quite
    If FRM_SelFic.CB_Annule Then
        Unload FRM_SelFic
        End
    End If
    cibleNomCatia = FRM_SelFic.Tbx_FicNom
    Unload FRM_SelFic

'Initialisation des nom de fichiers et des chemins
    CheminSourcesMacro = Get_Active_CATVBA_Path
    
'Chargement de la barre de progression
    Set mbarre = New ProgressBarre
    mbarre.Titre = "Export de la nomenclature ordo."
    mbarre.Progression = 1
    mbarre.Affiche
      
'Initialisation de l'objet Excel (nomenclature Catia)
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWBk = objexcel.Workbooks.Open(CStr(cibleNomCatia))
    objexcel.Visible = True
    
    'Recherche la position des lignes dans le fichier Excel
    NoLigEntete = 4
    NoLigFinNom = NoDerniereLigne(objWBk)
    NoLigDebSSE = 4
    NoLigDebDet = NoDebRecap(objWBk) + 2
    NoLigFinEns = NoLigDebDet - 2

'Collecte de la liste des sous ensembles
    Set oLigSSens = GetListSSE(objWBk)

'Collecte de la valeur des attributs dans le fichier eXcel et calcul des Qte de sous ensembles
    Set olignomCatias = GetBomOrdoXl(objWBk, mbarre)
    
'Export du de la nomenclature Ordo
    PutNomOrdo olignomCatias, mbarre
    
MsgBox "Fin de l'export de la nomenclature ordo !", vbInformation, "Fin de traitement"

erreur:

fin:
'Lib�ration des classes
    Set mbarre = Nothing
    Set objexcel = Nothing
    Set olignomCatias = Nothing

End Sub

Private Function GetBomOrdoXl(objWBk, mbarre) As c_LignomOrdos
'construit la liste des lignes de nomenclatures pour chaque sous ensemble avec le nom du parent de chaque �l�ment
'Ouvre le fichier excel de la nomenclature g�n�r�e par catia

Dim LigActive As Long
Dim ColActive As Integer
Dim nCellActive As String
Dim NomSSE As String    'Nom du sous ensemble parent
Dim NomRef As String   'Nom de la reference
Dim QteRef As Long   'Qte de la reference
Dim oLigNom As c_LignomOrdo
Dim oLigNoms As c_LignomOrdos
Dim oLigSSEn As c_LNomencl
'Dim oLigSSens As c_LNomencls
Dim oattribut As c_Attribut
Dim oAttributs As c_Attributs
Dim oAttributEnv As c_Attribut
Dim oAttributEnvs As c_Attributs
Dim pEtape As Long, pNbEt As Long, pItem As Long, pItems As Long

'initialisation des classes
    Set oLigNoms = New c_LignomOrdos
    Set oAttributEnv = New c_Attribut
    Set oAttributEnvs = New c_Attributs
    Set oattribut = New c_Attribut
    Set oAttributs = New c_Attributs
    
    'collecte de la liste des propri�t�s sp�cifiques � l'environnement (nom des attributs)
    LigActive = 4
    ColActive = 4 'Saut des champs "Qte", "Part number", "source"
    While objWBk.ActiveSheet.cells(LigActive, ColActive).Value <> ""
        oAttributEnv.Nom = objWBk.ActiveSheet.cells(LigActive, ColActive).Value
        oAttributEnv.Ordre = ColActive
        oAttributEnvs.Add oAttributEnv.Nom, oAttributEnv.Ordre
        ColActive = ColActive + 1
    Wend
    
'Collecte de la valeur des attributs de chaque ligne de nomenclature avec le nom de l'ensemble parent
    LigActive = NoLigEntete - 1
    'Collecte de l'ensemble g�n�ral
    Set oLigNom = New c_LignomOrdo
    NomSSE = TestEstSSE(objWBk.ActiveSheet.cells(LigActive, 1).Value)
    NomSSE = suppZero(NomSSE)
    oLigNoms.Add NomSSE, "E", 1
    pEtape = 1: pNbEt = 2: pItem = 0: pItems = NoLigFinEns - LigActive
    While LigActive < NoLigFinEns
        pItem = pItem + 1
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Collecte des infos dans la nomenclature Catia"
        nCellActive = objWBk.ActiveSheet.cells(LigActive, 1).Value
        'Saute les lignes d'ent�te et les lignes vides
        If nCellActive = "" Or nCellActive = "Quantit�" Then
            LigActive = LigActive + 1
        Else
            Set oLigNom = New c_LignomOrdo
            Set oAttributs = New c_Attributs
            'Stocke et formate le nom du sous ensemble parent
            If TestEstSSE(objWBk.ActiveSheet.cells(LigActive, 1).Value) <> "False" Then
                NomSSE = TestEstSSE(objWBk.ActiveSheet.cells(LigActive, 1).Value)
                LigActive = LigActive + 2 'Passage a la liste des composants du sous ensemble
                'suppression des z�ro non significatifs
                NomSSE = suppZero(NomSSE)
                oLigNom.Comp = "E"
            End If
            ColActive = 1
            oLigNom.Qte = objWBk.ActiveSheet.cells(LigActive, ColActive)
            ColActive = ColActive + 1
            oLigNom.Ref = objWBk.ActiveSheet.cells(LigActive, ColActive)
            ColActive = ColActive + 1
            oLigNom.Source = FormatSource(objWBk.ActiveSheet.cells(LigActive, ColActive))
            ColActive = ColActive + 1
            oLigNom.Parent = NomSSE
            oLigNom.Comp = "D"
            'Collecte des autres propri�t�s
            For Each oAttributEnv In oAttributEnvs.Items
                oattribut.Nom = oAttributEnv.Nom
                oattribut.Valeur = objWBk.ActiveSheet.cells(LigActive, ColActive)
                ColActive = ColActive + 1
                oAttributs.Add oattribut.Nom, False, oattribut.Valeur
            Next
            oLigNom.Attributs = oAttributs
            Set oAttributs = Nothing
            oLigNoms.Add oLigNom.Ref, oLigNom.Comp, oLigNom.Qte, oLigNom.Source, oLigNom.Desc, oLigNom.Attributs, oLigNom.Parent
            Set oLigNom = Nothing
            
            LigActive = LigActive + 1
            pItem = pItem + 1
        End If
    Wend
        
    'Calcul des quantit� des sous ensembles
    LigActive = NoLigEntete - 1
    While LigActive < NoLigFinEns
        NomRef = TestEstSSE(objWBk.ActiveSheet.cells(LigActive, 1).Value)
        If NomRef <> "False" Then
            LigActive = LigActive + 2 'Passage a la liste des composants du sous ensemble
        End If
        
        NomRef = objWBk.ActiveSheet.cells(LigActive, 2).Value
        QteRef = CInt(objWBk.ActiveSheet.cells(LigActive, 1).Value)
        While NomRef <> ""
            On Error Resume Next
            Set oLigSSEn = oLigSSens.Item(NomRef)
            If oLigSSEn Is Nothing Then
                err.Clear
                On Error GoTo 0
            Else
                oLigSSEn.Qte = CalculQte(NomRef, oLigNoms)  '* QteRef
                oLigSSens.Item(NomRef).Qte = oLigSSEn.Qte
            End If
            LigActive = LigActive + 1
            NomRef = objWBk.ActiveSheet.cells(LigActive, 2).Value
            QteRef = CInt(objWBk.ActiveSheet.cells(LigActive, 1).Value)
            Set oLigSSEn = Nothing
        Wend
        LigActive = LigActive + 1
    Wend
    objWBk.Close False
 
Set GetBomOrdoXl = oLigNoms

erreur:
'Lib�ration des classes
    Set oLigNoms = Nothing
    Set oLigNom = Nothing
    Set oattribut = Nothing
    Set oAttributs = Nothing
    Set oAttributEnv = Nothing
    Set oAttributEnvs = Nothing

End Function

Public Sub PutNomOrdo(olignomCatias As c_LignomOrdos, mbarre)
' *****************************************************************
' * Export dans le template excel Ordo de la nomenclature dans C:\temp
' *
' * Cr�ation CFR le 23/11/16
' * Derni�re modification le :
' *****************************************************************

Dim objExcelNomOrdo
Dim objWBkOrdo
Dim oLignomOrdo As c_LignomOrdo
Dim oLignomsSEn As c_LNomencl
Dim oattribut As c_Attribut
Dim oAttributs As c_Attributs
Dim SSERef As String
'Dim CellQteAss As String
Dim CellQteSSE As String
Dim LigActiveOrdo As Long
Dim LigActiveNom As Long
Dim FormuleSSE As String
Dim Formulepiece As String
Dim pEtape As Long, pNbEt As Long, pItem As Long, pItems As Long

'Initialisation des classes
    Set oLignomOrdo = New c_LignomOrdo
    Set oLignomsSEn = New c_LNomencl
    Set oattribut = New c_Attribut
    Set oAttributs = New c_Attributs

'Creation d'un objet eXcel et ouverture du fichier Template ordo
    Set objExcelNomOrdo = CreateObject("EXCEL.APPLICATION")
    Set objWBkOrdo = objExcelNomOrdo.Workbooks.Open(CStr(CheminSourcesMacro & pubNomTemplateOrdo))
    objExcelNomOrdo.Visible = True
    objWBkOrdo.ActiveSheet.Visible = True

'    CellQteAss = "$E$3" 'Cellule contenant la quantit� d'assemblage
    LigActiveOrdo = 3
    LigActiveNom = 5


    pEtape = 2: pNbEt = 2: pItem = 0: pItems = olignomCatias.Count
    'Ecriture de la nomenclature
    For Each oLignomOrdo In olignomCatias.Items
        pItem = pItem + 1
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Ecriture des infos dans la nomenclature Catia"
        'Test si c'est l'ensemble g�n�ral
        If oLignomOrdo.Parent = "" Then
            ' Ecriture du nom de l'assemblage
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_nMachine).Value = olignomCatias.Item(1).Ref
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Qte).Value = olignomCatias.Item(1).Qte
            'Stockage de la cellule de la quantit� de SSe a commander (pour les lignes pieces de l'ensemble g�n�ral)
            'CellQteSSE = "$" & NumCar(Ord_Qte) & "$" & LigActiveOrdo
            CellQteSSE = 1
            SSERef = oLignomOrdo.Ref
            LigActiveOrdo = LigActiveOrdo + 1
        Else
            'test si c'est un Sous ensemble
            'On se sert du parent de la ligne de nomenclature en cours pour �crire la ligne de regroupement du sous ensemble
            'La ligne de nomenclature de la piece ou du sous ensemnle en cours est �crite dans le If suivant
            If SSERef <> oLignomOrdo.Parent Then 'Changment de sous ensemble
                SSERef = oLignomOrdo.Parent
                Set oAttributs = oLignomOrdo.Attributs
                'recherche de la Qte se sous ensemble
                Set oLignomsSEn = oLigSSens.Item(SSERef)
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_nSSE).Value = SSERef
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Qte).Value = oLignomsSEn.Qte   'Qte du sous ensemble
                'Formule de calcul de Qt� a commander
                'FormuleSSE = "=IF((E" & LigActiveOrdo & "-K" & LigActiveOrdo & ")<0,0,E" & LigActiveOrdo & "-K" & LigActiveOrdo & ")"
                FormuleSSE = "=IF((" & NumCar(Ord_Qte) & LigActiveOrdo & "-" & NumCar(Ord_QteStock) & LigActiveOrdo & ")<0,0," & NumCar(Ord_Qte) & LigActiveOrdo & "-" & NumCar(Ord_QteStock) & LigActiveOrdo & ")"
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_QteCmd).Formula = FormuleSSE
                'Stockage de la cellule de la quantit� de SSe a commander (pour les lignes pieces des sous ensembles)
                CellQteSSE = "$" & NumCar(Ord_QteCmd) & "$" & LigActiveOrdo
                'Mise forme de la ligne du sous Ensemble
                objWBkOrdo.ActiveSheet.Range(CStr("A" & LigActiveOrdo & ":U" & LigActiveOrdo)).Interior.Color = 16751052
                Set oAttributs = Nothing
                LigActiveOrdo = LigActiveOrdo + 1
            End If
            
            'test si c'est un composant
            If oLignomOrdo.Comp = "D" Then
                Set oAttributs = oLignomOrdo.Attributs
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Repere).Value = oLignomOrdo.Ref   'Rep
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Qte).Value = oLignomOrdo.Qte   'Qte
                'Formule de calcul de Qt� a commander
                Formulepiece = "=IF(((" & NumCar(Ord_Qte) & LigActiveOrdo & "*" & CellQteAss & "*" & CellQteSSE & ")-" & NumCar(Ord_QteStock) & LigActiveOrdo & ")<0,0,(" & NumCar(Ord_Qte) & LigActiveOrdo & "*" & CellQteAss & "*" & CellQteSSE & ")-" & NumCar(Ord_QteStock) & LigActiveOrdo & ")"
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_QteCmd).Formula = Formulepiece
                objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Type).Value = FormatSource(oLignomOrdo.Source)  'Source
                'Attributs sp�cifiques � l'environnement
                If oAttributs.Count > 0 Then
                    PutNomOrdoAttributs oAttributs, objWBkOrdo, LigActiveOrdo
                End If
                'Traitement des sp�cificit�s de l'environnement
                PutNomOrdoSpecif oAttributs, objWBkOrdo, LigActiveOrdo, pubNomEnv
                LigActiveOrdo = LigActiveOrdo + 1
                Set oAttributs = Nothing
            End If
        End If
    Next
End Sub

Private Sub PutNomOrdoAttributs(oAttributs As c_Attributs, objWBkOrdo, LigActiveOrdo As Long)
'Ecrit la valeur des attributs sp�cifiques de l'environnement dans la nomenclature ordo
'oAttributs = collection des attributs de la ligne de nomenclature en cours
'objWBkOrdo = template excel ordo
'LigActiveOrdo = N� de la ligne en cours dans le fichier excel
Dim oattribut As c_Attribut

    For Each oattribut In pubAttributs.Items
        If oattribut.Ordre <> 0 Then
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, oattribut.Ordre).Value = oAttributs.Item(oattribut.Nom).Valeur
        End If
    Next

End Sub

Private Sub PutNomOrdoSpecif(oAttributs As c_Attributs, objWBkOrdo, LigActiveOrdo As Long, Env As String)
'Traite les sp�cificit�es de l'environnement dans la nomenclature ordo
'oAttributs = collection des attributs de la ligne de nomenclature en cours
'objWBkOrdo = template excel ordo
'LigActiveOrdo = N� de la ligne en cours dans le fichier excel
'Env = Nom de l'environnement (Airbus, GSE, Snecma, etc...)
Dim oattribut As c_Attribut

    Select Case Env
        Case "SNECMA"
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Designation).Value = oAttributs.Item(nDesc).Valeur
        Case "GSE"
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Designation).Value = oAttributs.Item(nDesc).Valeur
        Case "EXCENT"
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Designation).Value = oAttributs.Item(nDesc).Valeur
        Case "AIRBUS_UK"
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Designation).Value = oAttributs.Item(nDesc).Valeur
        Case "AIRBUS_FR"
            objWBkOrdo.ActiveSheet.cells(LigActiveOrdo, Ord_Designation).Value = oAttributs.Item(nDesc).Valeur
            
    End Select

End Sub

Private Function GetListSSE(objWBk) As c_LNomencls
'Collecte de la liste des sous ensembles dans le fichier eXcel
Dim LigActive As Long
Dim QteRef As Long   'Qte de la reference
Dim oLstSSE As c_LNomencls

Set oLstSSE = New c_LNomencls
    
    LigActive = NoLigEntete - 1
    While LigActive < NoLigFinEns
        If TestEstSSE(objWBk.ActiveSheet.cells(LigActive, 1).Value) <> "False" Then
            'force la quantit� de l'ensemble g�n�ral � 1
            If LigActive = NoLigEntete - 1 Then QteRef = 1 Else QteRef = 0
            oLstSSE.Add suppZero(TestEstSSE(objWBk.ActiveSheet.cells(LigActive, 1).Value)), "E", QteRef
        End If
        LigActive = LigActive + 1
    Wend
    Set GetListSSE = oLstSSE
End Function

Private Function suppZero(str As String) As String
'Supprime les z�ros en t�te de la string
    While Left(str, 1) = "0"
        str = Right(str, Len(str) - 1)
    Wend
    suppZero = str
End Function

Private Function CalculQte(NSSe As String, oLigNoms As c_LignomOrdos) As Integer
'Fonction r�cursive permettant de calculer la quantite d'un sous-ensemble
'en fonction de la quantit�s des sous ensembles dans lequel il est utilis�
'NSSe = nom du sous ensemble dont on veux calculer la Qte
'olignom = collection des sous ensembles avec leur qte et leur parent
Dim oLignomSSE As c_LignomOrdo
        For Each oLignomSSE In oLigNoms.Items
            If oLignomSSE.Ref = NSSe Then
                If oLignomSSE.Parent = "" Then
                    CalculQte = 1
                Else
                    CalculQte = CalculQte + oLignomSSE.Qte * CalculQte(oLignomSSE.Parent, oLigNoms)
                End If
            End If
        Next
End Function

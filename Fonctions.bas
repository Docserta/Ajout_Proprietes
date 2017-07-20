Attribute VB_Name = "Fonctions"
Public Function CompIsactive(mprod As Product, nParent As String) As Boolean
'V�rifie si le produit (composant) pass� en argument est actif ou non
Dim mparams As Parameters
Dim mparam As Parameter
Dim ncomplet As String

    ncomplet = nParent & "\" & mprod.Name & "\" & nParamActivate
    Set mparams = mprod.Parameters
    
    Set mparam = mparams.Item(ncomplet)
    'Set mparam = mparams.Item(nParamActivate)
    CompIsactive = mparam.Value

End Function

Public Function check_Env(Env As String) As Boolean
'Check si l'environnement est conforme aux pr�requis de lancement des macros
'Env = "Part", "Product"
On Error Resume Next
Dim mPart As PartDocument
Dim mprod As ProductDocument

Select Case Env
    Case "Parts"
        Set mPart = CATIA.ActiveDocument
        If err.Number <> 0 Then
            MsgBox "Activez un CATPart avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            err.Clear
            End
        Else
            check_Env = True
        End If
    Case "Product" 'Test si un CatProduct est actif
        Set mprod = CATIA.ActiveDocument
        If err.Number <> 0 Then
            MsgBox "Activez un Catproduct avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            err.Clear
            End
        Else
            check_Env = True
        End If
    End Select
    
    On Error GoTo 0
End Function

Public Function Check_PartProd(mdoc) As Boolean
'V�rifie si le document pass� en argument est un part ou un product
'sert a sauter les .ctcfg etc
On Error Resume Next
Dim mPart As PartDocument
Dim mprod As ProductDocument
Check_PartProd = False
Set mPart = mdoc
    If err.Number <> 0 Then
        err.Clear
    Else
        Check_PartProd = True
    End If
    Set mprod = mdoc
    If err.Number <> 0 Then
        err.Clear
    Else
        Check_PartProd = True
    End If
End Function

Public Sub CreateParamExist(mparams As Parameters, nParam As String, vParam As String)
'test si le param�tre pass� en argument existe dans le product.
'si oui lui affecte la valeur vParam
'sinon le cr�e et lui affecte la valeur vParam
'mParams Collection des param�tres du product
'nParam Nom du param�tre
'vParam valeur du param�tre

Dim oParam As StrParam
On Error Resume Next
    Set oParam = mparams.Item(nParam)
    If (err.Number <> 0) Then
        ' Le param�tre n'existe pas, on le cr�e
        err.Clear
        Set oParam = mparams.CreateString(nParam, vParam)
    Else
        oParam.Value = vParam
    End If
End Sub

Public Sub CreateParam(mparams As Parameters, nParam As String, vParam As String)
'test si le param�tre pass� en argument existe dans le product.
'sinon le cr�e et lui affecte la valeur vParam
'si oui, ne fait rien
'mParams Collection des param�tres du product
'nParam Nom du param�tre
'vParam valeur du param�tre

Dim oParam As StrParam
On Error Resume Next
    Set oParam = mparams.Item(nParam)
    If (err.Number <> 0) Then
        ' Le param�tre n'existe pas, on le cr�e
        err.Clear
        Set oParam = mparams.CreateString(nParam, vParam)
    Else
        'On ne fait rien
    End If
End Sub

Public Function FileExist(StrFile As String) As Boolean
'Teste si le fichier existe et revois vrai ou faux
'StrFile = nom du chemin complet jusqu'au r�pertoire a tester ex "c:\temp\test"
Dim fs, f
    Set fs = CreateObject("scripting.filesystemobject")
    On Error Resume Next
    Set f = fs.GetFile(StrFile)
    If err.Number <> 0 Then
        err.Clear
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Function NoDebLstPieces(objWBk) As Long
'recherche la 1ere ligne des attributs des pi�ces
' la ligne commence par "Liste des pi�ces"
    Dim NomSeparateur As String
    Dim NoLigne As Long
    NoLigne = 1
    NomSeparateur = "Liste des pi�ces"
    While Left(objWBk.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebLstPieces = NoLigne

End Function

Public Function NoDebRecap(objWBk As Variant) As Long
'recherche la 1ere ligne du r�capitulatif des pi�ces
' la ligne commence par "Nomenclature de" ou "Recapitulation of:"
    Dim NomSeparateur As String
    Dim NoLigne As Integer
    NoLigne = 1
    If Langue = "EN" Then
        NomSeparateur = "Recapitulation of:"
    ElseIf Langue = "FR" Then
        NomSeparateur = "R�capitulatif sur"
    End If
    While Left(objWBk.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebRecap = NoLigne
End Function

Public Function NoDerniereLigne(objWBk As Variant) As Long
'recherche la derni�re ligne du fichier excel
'On part du principe que 2 lignes vide indiquent la fin du fichier
Dim NoLigne As Integer, NbLigVide As Integer
    NoLigne = 1
    NbLigVide = 0
    While NbLigVide < 2
        If objWBk.ActiveSheet.cells(NoLigne, 1).Value = "" Then
            NbLigVide = NbLigVide + 1
        Else
            NbLigVide = 0
        End If
    NoLigne = NoLigne + 1
    Wend
    NoDerniereLigne = NoLigne - 2
End Function

Public Function SupZero(str As String) As String
'supprime les z�ros non significatif du partNumber
'test si la string contien un nombre
If IsNumeric(str) Then
    SupZero = CDbl(str)
Else
    SupZero = str
End If
End Function

Public Function TestEstSSE(Ligne As String) As String
'test si la ligne correspond a une ent�te de sous ensemble
' la ligne commence par "Nomenclature de" ou "Bill of Material"
Dim NomSeparateur As String
Dim tmpNomSSE As String

    If Langue = "EN" Then
        NomSeparateur = "Bill of Material: "
    ElseIf Langue = "FR" Then
        NomSeparateur = "Nomenclature de "
    End If
    On Error Resume Next 'Test si la chaine est vide ou inf�rieur a len(noms�parateur)
    tmpNomSSE = Right(Ligne, Len(Ligne) - Len(NomSeparateur))
    If err.Number <> 0 Then
         TestEstSSE = "False"
    Else
        If Left(Ligne, Len(NomSeparateur)) = NomSeparateur Then
            TestEstSSE = tmpNomSSE
        Else
            TestEstSSE = "False"
        End If
    End If
End Function

Public Function TestParamExist(mparams As Parameters, nParam As String) As String
'test si le param�tre pass� en argument existe dans le part.
'si oui renvoi sa valeur,
'sinon la cr�e et lui affecte une chaine vide
Dim oParam As StrParam
On Error Resume Next
    Set oParam = mparams.Item(nParam)
If (err.Number <> 0) Then
    ' Le param�tre n'existe pas, on le cr�e
    err.Clear
    Set oParam = mparams.CreateString(nParam, "")
    oParam.Value = ""
End If
TestParamExist = oParam.Value
End Function

Public Sub GetAttributs(nFile As String)
'Collecte la liste des attributs dans le fichier text
'les stocke dans la collection public pubAttributs
'nFile = chemin + non du fichier des attributs
Dim fs, f
Dim oattribut As c_Attribut
Dim Linestr As String
On Error GoTo erreur:

    Set pubAttributs = New c_Attributs
    Set oattribut = New c_Attribut
    If nFile <> "" Then
        'Ouverture du fichier de param�tres
        Set fs = CreateObject("Scripting.FileSystemObject")
        'teste si le fichier existe
        If FileExist(nFile) Then
            Set f = fs.opentextfile(nFile, ForReading, 1)
            'lecture du nom de l'environnement
            pubNomEnv = SplitSemicolon(f.ReadLine, 2)
            'Lecture su nom du template ordo
            pubNomTemplateOrdo = SplitSemicolon(f.ReadLine, 2)
            'lecture des param�tres
            Do While Not f.AtEndOfStream
                Linestr = f.ReadLine
                If SplitSemicolon(Linestr, 1) = "Attrib" Then
                    oattribut.Nom = SplitSemicolon(Linestr, 2)
                    'oattribut.Valeur = SplitSemicolon(Linestr, 3)
                    On Error GoTo erreur:
                    oattribut.Ordre = CInt(SplitSemicolon(Linestr, 3))
                    pubAttributs.Add oattribut.Nom, oattribut.Ordre, oattribut.Valeur
                End If
            Loop
        Else
            MsgBox "Le fichier d'attributs d�fini dans vos pr�f�rences " & Chr(10) & "(" & nFile & ")" & Chr(10) & " est introuvable. s�lectionnez un autre fichier.", vbInformation
        End If
    End If
    GoTo fin:
    
erreur:
    MsgBox " Une erreur a �t� d�tect�e dans le fichier des attributs : " & nFile, vbCritical, "Erreur dans fichier attributs"
    
fin:
'Lib�ration des classes
Set oattribut = Nothing
Set f = Nothing
Set fs = Nothing
End Sub


Public Sub InitLanguage()
'Configure les champs en fonction de la langue
    Langue = Language
    If Langue = "EN" Then
        nQt = "Quantity"
        nRef = "Part Number"
        nRev = "Revision"
        nDef = "Definition"
        nNom = "Nomenclature"
        nDesc = "Product Description"
        nSrce = "Source"
        nParamActivate = "Component Activation State"
    Else
        nQt = "Quantit�"
        nRef = "R�f�rence"
        nRev = "R�vision"
        nDef = "D�finition"
        nNom = "Nomenclature"
        nDesc = "Description du produit"
        nSrce = "Source"
        nParamActivate = "Etat d'activation du composant"
    End If
End Sub

Public Function Language() As String
'D�tecte la langue de l'interface Catia
'Ouvre un part vierge et test le nom du "Main Body"
Dim ofile, oFolder, ofs
Dim EmptyPartFolder, EmptyPartFile
Dim oEmptyPart  As PartDocument

On Error Resume Next
Set ofs = CreateObject("Scripting.FileSystemObject")
Set oFolder = ofs.GetFolder(CATIA.Parent.Path)
Set EmptyPartFolder = ofs.GetFolder(oFolder.ParentFolder.ParentFolder.Path & "\startup\templates") ' dossier relatif des mod�les vides
Set EmptyPartFile = ofs.GetFile(EmptyPartFolder.Path & "\empty.CATPart")

If err.Number = 0 Then
    On Error GoTo 0
    Set oEmptyPart = CATIA.Documents.Open(EmptyPartFile.Path)
    If oEmptyPart.Part.MainBody.Name = "PartBody" Then
        Language = "EN"
    Else
        Language = "FR"
    End If
Else
    err.Clear
    On Error GoTo 0
End If

    oEmptyPart.Close
 Set oEmptyPart = Nothing
 Set EmptyPartFile = Nothing
 Set EmptyPartFolder = Nothing
 Set oFolder = Nothing
 Set ofs = Nothing

End Function


Public Sub LogUtilMacro(ByVal mPath As String, ByVal mFic As String, ByVal mMacro As String, ByVal mModule As String, ByVal mVersion As String)
'Log l'utilisation de la macro
'Ecrit une ligne dans un fichier de log sur le serveur
'mPath = localisation du fichier de log ("\\serveur\partage")
'mFic = Nom du fichier de log ("logUtilMacro.txt")
'mMacro = nom de la macro ("NomGSE")
'mVersion = Version de la macro ("version 9.1.4")
'mModule = Nom du module ("_Info_Outillage")

Dim mDate As String
Dim mUser As String
Dim nFicLog As String
Dim LigLog As String
Const ForWriting = 2, ForAppending = 8

    mDate = Date & " " & Time()
    mUser = ReturnUserName()
    nFicLog = mPath & "\" & mFic

    nliglog = mDate & ";" & mUser & ";" & mMacro & ";" & mModule & ";" & mVersion

    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile(nFicLog)
    If err.Number <> 0 Then
        Set f = fs.opentextfile(nFicLog, ForWriting, 1)
    Else
        Set f = fs.opentextfile(nFicLog, ForAppending, 1)
    End If
    
    f.Writeline nliglog
    f.Close
    On Error GoTo 0

End Sub

Function ReturnUserName() As String 'extrait d'un code de Paul, Dave Peterson Exelabo
'Renvoi le user name de l'utilisateur de la station
'fonctionne avec la fonction GetUserName dans l'ent�te de d�claration
    Dim Buffer As String * 256
    Dim BuffLen As Long
    BuffLen = 256
    If GetUserName(Buffer, BuffLen) Then _
    ReturnUserName = Left(Buffer, BuffLen - 1)
End Function

Public Function EffaceFicNom(nfold As String, nFic As String) As Boolean
'Effacement d'un fichier de nomenclature pr�-existant
'nfold = nom du r�pertoire
'nFic = nom du fichier

 On Error GoTo Err_EffaceFicNom
    Dim ofs, ofold, ofiles, ofile
    Set ofs = CreateObject("Scripting.FileSystemObject")
    Set ofold = ofs.GetFolder(nfold)
    Set ofiles = ofold.Files
    For Each ofile In ofiles
        If ofile.Name = nFic Then
            ofs.DeleteFile (CStr(nfold & "\" & nFic))
        End If
    Next
    EffaceFicNom = True
GoTo Quit_EffaceFicNom

Err_EffaceFicNom:
MsgBox "Il est possible que le fichier de nomenclature soit encore ouvert dans Excel. Veuillez le fermer et relancer la macro.", vbCritical, "erreur"
EffaceFicNom = False
Quit_EffaceFicNom:
End Function

Public Function FormatSource(str As String) As String
'formate le contenu du champs source
'remplace "Inconu" ou "Unknown" par une chaine vide.
'remplace les codes champs sources par une string
FormatSource = str
Select Case str
    Case "Inconnu", "Unknown"
        FormatSource = ""
    Case "Bought", catProductBought
        FormatSource = "Achet�"
    Case "Made", catProductMade
        FormatSource = "Fabriqu�"
End Select
End Function

Public Function ReverseSource(str As String) As CatProductSource
'Renvoi le code champs source
'remplace "Achet�" par catProductBought
' et "Fabriqu�" par catProductMade
ReverseSource = catProductSourceUnknown
Select Case str
    Case "Achet�"
        ReverseSource = catProductBought
    Case "Fabriqu�"
        ReverseSource = catProductMade
End Select
End Function

Public Function Get_Active_CATVBA_Path() As String
Dim APC_Obj As New MSAPC.Apc
Dim TempName As String
Dim i As Long
   TempName = APC_Obj.VBE.ActiveVBProject.FileName
   For i = Len(TempName) To 1 Step -1
        If Mid(TempName, i, 1) = "\" Then
            TempName = Left(TempName, i)
            Exit For
        End If
   Next
   Get_Active_CATVBA_Path = TempName
End Function

Public Function RestPref() As String
'Restaure les pr�f�rences a partir d'un fichier texte dans c:\temp
Dim fs, f
Dim LigTxt As String, Rlig As String, Llig As String
Dim pos As Integer
On Error GoTo err
    
    Set fs = CreateObject("scripting.filesystemobject")
    Set f = fs.opentextfile(nFicPref, ForReading, 1)
    
    LigTxt = f.ReadLine
    pos = InStr(1, LigTxt, "=") 'calcule la position du "="
    RestPref = Right(LigTxt, Len(LigTxt) - pos) 'Valeur de la constante (a droite du "=")
    GoTo fin
    
err:
    RestPref = ""
fin:
    f.Close
    Set fs = Nothing
End Function


Public Sub SauvPref()
'Sauvegarde les options dans un fichier texte dans c:\temp
Dim LigEncours As String
Dim fs, f
    Set fs = CreateObject("scripting.filesystemobject")
    Set f = fs.CreateTextFile(nFicPref, True)
    LigEncours = "PathFicAttributs=" & pubPathEnv
    f.Writeline (LigEncours)
    f.Close
    Set fs = Nothing
End Sub

Public Function SplitSlash(str As String) As String
'Recup�re la partie finale d'une string apres le dernier "\"
Do While InStr(1, str, "\", vbTextCompare) > 0
    str = Right(str, Len(str) - InStr(1, str, "\", vbTextCompare))
Loop
SplitSlash = str
End Function

Public Function SplitSemicolon(ByVal str As String, pos As Integer) As String
'Renvois la nieme partie d'une string s�par�e par des ";"
Dim SplitStr() As String
Dim i As Integer
On Error GoTo erreur:

    Do While InStr(1, str, ";", vbTextCompare) > 0
        ReDim Preserve SplitStr(i)
        SplitStr(i) = Left(str, InStr(1, str, ";", vbTextCompare) - 1)
        str = Right(str, Len(str) - Len(SplitStr(i)) - 1)
        i = i + 1
    Loop
        ReDim Preserve SplitStr(i)
        SplitStr(i) = str
    If i > 0 Then
        SplitSemicolon = SplitStr(pos - 1)
    Else
        SplitSemicolon = ""
    End If
    GoTo fin

erreur:
    SplitSemicolon = ""
fin:

End Function

Public Function NomENS(Ligne As String) As String
' Renvoi le nom de l'ensemble
' la ligne commence par "Nomenclature de " ou "Bill of Material: "
' puis est suivie du nom de l'ensemble
Dim NomSeparateur As String
Dim tmpNomSSE As String

    If Langue = "EN" Then
        NomSeparateur = "Bill of Material: "
    ElseIf Langue = "FR" Then
        NomSeparateur = "Nomenclature de "
    End If
    On Error Resume Next 'Test si la chaine est vide ou inf�rieur a len(noms�parateur)
    NomENS = Right(Ligne, Len(Ligne) - Len(NomSeparateur))
    If err.Number <> 0 Then
         NomENS = ""
    End If
End Function

Public Function NumCar(Num As Integer) As String
'Converti un chiffre en lettre
'1 = A, 2 = B etc
'Attention la num�rotation de Array commence � 0 d'ou le double A dans la liste
Dim ListCar
If Num > 78 Then ' a changer si on ajoute des colonnes a la liste Array
    Num = 1
End If
ListCar = Array("A", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
                "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
                "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")
NumCar = ListCar(Num)
End Function


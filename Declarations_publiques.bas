Attribute VB_Name = "Declarations_publiques"
'Fonction de r�cup�ration du username
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Version de la macro
Public Const VMacro As String = "Version 4.6.7 du 21/03/17"
Public Const nMacro As String = "Ajout_Proprietes"
Public Const nPath As String = "\\srvxsiordo\xLogs\01_CatiaMacros"
Public Const nFicLog As String = "logUtilMacro.txt"

'Nom du fichier de sauvgarde des pr�ferences
Public Const nFicPref As String = "C:\temp\PrefMacroAjoutProprietes.txt"

Public Nbre_Param As Integer

Public Langue As String

Public Const ForReading As Integer = 1

Public pubAttributs As c_Attributs 'Collection des attributs collect�es dans le fichier texte
Public pubNomEnv As String          'Nom de l'environnement
Public pubPathEnv As String         'Chemin + nom du fichier des attributs de l'environnement
Public pubNomTemplateOrdo As String
'Public NomTemplateNomOrdo As String 'Nom du fichier template de la nomenclature excel Ordo

Public oProduit As c_Product
Public oProduits As c_Products

Public Const MaxAttributs As Integer = 20

'Variable fichier excel
Public nNomCatia As String      'Nom du fichier de la nomenclature g�n�r�e par Catia
Public cibleNomCatia As String      'Chemin + nom du fichier de nomenclature g�n�r�e par Catia
Public Const nExtNomCatia As String = "-NomCatia"    'Extention ajout�e au fichier excel nomenclature g�n�r�e par Catia
Public Const nSheetExtractCatia As String = "ExtractCatia" 'Nom de l'onglet "Feuile.1" dans le fichier de nomenclature g�n�r�e par Catia

Public nNomProp As String       'Nom du fichier excel de la nomenclature des propri�t�s modifiables
Public cibleNomProp As String   'Chemin + nom du fichier excel de la nomenclature des propri�t�s modifiables
Public Const nExtNomProp As String = "-Proprietes"     'Extention ajout�e au fichier excel des propri�t�s modifiables
Public Const nSheetReacp As String = "Recapitulatif"        'Nom de l'onglet "Recapitulatif" dans le fichier des propri�t�es modifiables


Public cibleNomOrdo As String   'Chemin + nom du template de la nomenclature excel Ordo

Public CheminSourcesMacro As String         'Chemin dans lequel est lanc� la macro

'Position des �l�ments dans le fichier excel de nomenclature Catia
Public NoLigEntete As Long 'N� de la ligne des ent�te d'attributs dans le fichier eXcel
Public NoLigFinNom As Long 'N� de la ligne de fin de la nomenclature
Public NoLigFinEns As Long 'N� de la ligne de fin des ensembles
Public NoLigDebDet As Long 'N� de la ligne de d�but des pi�ces
Public NoLigDebSSE As Long 'N� de la ligne de d�but des ensembles

'Position des colonnes dans la nomenclature Ordo
Public Ord_nMachine As Integer
Public Ord_nSSE As Integer
Public Ord_CodeX3 As Integer
Public Ord_Repere As Integer
Public Ord_nPlan As Integer
Public Ord_Qte As Integer
Public Ord_Indice As Integer
Public Ord_Planche As Integer
Public Ord_Designation As Integer
Public Ord_Marquage As Integer
Public Ord_Marque As Integer
Public Ord_Fournisseur As Integer
Public Ord_QteStock As Integer
Public Ord_QteCmd As Integer
Public Ord_Type As Integer
Public Ord_Traitmnt As Integer
Public CellQteAss As String

'Nom des champs de nomenclature (fran�ais/Anglais)
Public nQt As String
Public nRef As String
Public nRev As String
Public nDef As String
Public nNom As String
Public nDesc As String
Public nSrce As String

'Nom des param�tres (fran�ais/Anglais
Public nParamActivate As String

'Nombre de param�tres de base dans les nomenclatures (Qte, Reference, R�vision, Definition, Nomenclature, source et product description)
'Les param�tre particuliers de l'environnement d�bute apres
Public Const NbColPrmStd As Integer = 7
'Nombre de param�tres non modifiables (Qte, Reference)
Public Const NbColPrmNomModif As Integer = 2

Public oLigSSens As c_LNomencls

Public Sub IniColTemplateOrdo(NomTemplate As String)
'Param�tre la position des colonnes du template Ordo en fonction du template choisi

    Select Case NomTemplate
        Case "NomOrdostd"
            pubNomTemplateOrdo = "FML-Trame de nomenclature-03.xls"
            'Position des colonnes
            Ord_nMachine = 1
            Ord_nSSE = 2
            Ord_Repere = 3
            Ord_nPlan = 4
            Ord_Qte = 5
            Ord_Indice = 6
            Ord_Planche = 7
            Ord_Designation = 8
            Ord_Marque = 9
            Ord_Fournisseur = 10
            Ord_QteStock = 11
            Ord_QteCmd = 12
            Ord_Type = 13
            Ord_Traitmnt = 17
            'Position des colonnes permettant les calcul de quantit�
            CellQteAss = "$E$3" 'Cellule contenant la quantit� d'assemblage
        Case "NomOrdX3"
            pubNomTemplateOrdo = "FML-Trame de nomenclature-05.xls"
            'Position des colonnes
            Ord_nMachine = 1
            Ord_nSSE = 2
            Ord_CodeX3 = 3
            Ord_Repere = 4
            Ord_nPlan = 5
            Ord_Qte = 6
            Ord_Indice = 7
            Ord_Planche = 8
            Ord_Designation = 9
            Ord_Marquage = 10
            Ord_Marque = 11
            Ord_Fournisseur = 12
            Ord_QteStock = 13
            Ord_QteCmd = 14
            Ord_Type = 15
            Ord_Traitmnt = 19
            'Position des colonnes permettant les calcul de quantit�
            CellQteAss = "$F$3" 'Cellule contenant la quantit� d'assemblage
            
    End Select
    



End Sub







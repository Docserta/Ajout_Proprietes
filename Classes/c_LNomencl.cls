VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_LNomencl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'c_LNomencl
'Collection des infos d'une ligne de nomenclature
'Les attributs Reference,Qte, Revsion, Definition, Nomenclature, Source et Description sont toujours présents dans les parts et Products
'Les attributs personalisés sont collectés dans la collection c_attributs

Private mvarRef As String            'Référence de l'objet (Part Number)
Private mvarComp As String           'Type de composant Ensemble("E") ou pièce("P")
Private mvarQte As Long              'Quantité de composant
Private mvarRev As String            'Revision
Private mvarDef As String            'Definition
Private mvarNom As String            'Nomenclature
Private mvarsource As String         'Source (Fabriqué ou Acheté)
Private mvarDesc As String           'Description du composant (description ref)
Private mvarAttributs As c_Attributs 'collection des attributs du composant
Private mvarParent As String         'Nom de l'ensemble parent

Public Property Let Ref(ByVal vData As String)
    mvarRef = vData
End Property

Public Property Get Ref() As String
    Ref = mvarRef
End Property

Public Property Let Comp(ByVal vData As String)
    mvarComp = vData
End Property

Public Property Get Comp() As String
    Comp = mvarComp
End Property

Public Property Let Qte(ByVal vData As Long)
    mvarQte = vData
End Property

Public Property Get Qte() As Long
    Qte = mvarQte
End Property

Public Property Let Rev(ByVal vData As String)
    mvarRev = vData
End Property

Public Property Get Rev() As String
    Rev = mvarRev
End Property

Public Property Let Def(ByVal vData As String)
    mvarDef = vData
End Property

Public Property Get Def() As String
    Def = mvarDef
End Property

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let Source(ByVal vData As String)
    mvarsource = vData
End Property

Property Get Source() As String
    Source = mvarsource
End Property

Public Property Let Desc(ByVal vData As String)
    mvarDesc = vData
End Property

Public Property Get Desc() As String
    Desc = mvarDesc
End Property

Public Property Let Attributs(ByVal oData As c_Attributs)
    Set mvarAttributs = oData
End Property

Public Property Get Attributs() As c_Attributs
    Set Attributs = mvarAttributs
End Property

Public Property Let Parent(ByVal vData As String)
    mvarParent = vData
End Property

Property Get Parent() As String
    Parent = mvarParent
End Property

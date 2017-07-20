VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Product

Private mvarNom As String
Private mvarProduct As Product
Private mvarParent As String

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let Produit(ByVal vData As Product)
    Set mvarProduct = vData
End Property

Public Property Get Produit() As Product
    Set Produit = mvarProduct
End Property

Public Property Let Parent(ByVal vData As String)
    mvarParent = vData
End Property

Public Property Get Parent() As String
    Parent = mvarParent
End Property


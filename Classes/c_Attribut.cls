VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Attribut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Attribut

Private mvarNom As String
Private mvarOrdre As Integer
Private mvarValeur As String

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let Ordre(ByVal vData As Integer)
    mvarOrdre = vData
End Property

Public Property Get Ordre() As Integer
    Ordre = mvarOrdre
End Property

Public Property Let Valeur(ByVal vData As String)
    mvarValeur = vData
End Property

Public Property Get Valeur() As String
    Valeur = mvarValeur
End Property

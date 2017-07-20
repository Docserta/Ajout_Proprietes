VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_LNomencls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_LNomencls
'Collection des lignes de nomenclature

Private mcol As Collection

Public Function Add(Ref As String, _
            Optional Comp As String = "", _
            Optional Qte As Long = 0, _
            Optional Rev As String = "", _
            Optional Def As String = "", _
            Optional Nom As String = "", _
            Optional Source As String = "", _
            Optional Desc As String = "", _
            Optional Attributs As c_Attributs = Nothing, _
            Optional Parent As String) As c_LNomencl
    
    Dim objNewMember As c_LNomencl
    Set objNewMember = New c_LNomencl
    
    objNewMember.Ref = Ref
    objNewMember.Comp = Comp
    objNewMember.Qte = Qte
    objNewMember.Rev = Rev
    objNewMember.Def = Def
    objNewMember.Nom = Nom
    objNewMember.Source = Source
    objNewMember.Desc = Desc
    objNewMember.Attributs = Attributs
    objNewMember.Parent = Parent

    mcol.Add objNewMember, Ref

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_LNomencl
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub







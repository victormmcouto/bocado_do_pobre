VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryAcompanhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Repository

Private tblmgmt As New TableManagement
Private me_repoCnjg As RepositoryConjuge
Private me_repoDpnt As RepositoryDependente

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksACOMPANHAMENTO.ListObjects(1))
End Sub

Public Function Repository_readByKey(ByVal key As Long, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim arr As Variant
    Dim acmpmnt As New Acompanhamento
    
    Set lr = tblmgmt.getListRow(key)
    
    arr = lr.Range.Value
    
    Set Repository_readByKey = acmpmnt.CreateFromArray(arr)
End Function

Public Function Repository_getTbl() As TableManagement
End Function

Public Function Repository_save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
    obj.setKey = CLng(obj.getAno & obj.getKeyCadastro)
    tblmgmt.InsertDataRow obj, autoIncrement
End Function

Public Function Repository_updateById(ByVal key As Long, ByRef obj As Object) As Boolean
End Function

Public Function Repository_deleteById(ByVal key As Long) As Boolean
End Function

Public Function Repository_deleteFkey(ByVal fk As Long, ByVal fkName As String) As Boolean
End Function



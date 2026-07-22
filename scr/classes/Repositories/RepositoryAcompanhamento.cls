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
    Set tblmgmt = tblmgmt.create(dbACOMPANHAMENTO.ListObjects(1))
End Sub

Public Function Repository_readByKey(ByVal key As Long, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim acmpmnt As New acompanhamento
    
    Set lr = tblmgmt.getListRow(key)
    
    Set Repository_readByKey = acmpmnt.CreateFromArray(lr.Range.Value)
End Function

Public Function Repository_getTbl() As TableManagement
End Function

Public Function Repository_save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
    obj.setKey = CLng(obj.getAno & obj.getKeyCadastro)
    tblmgmt.InsertDataRow obj, autoIncrement
End Function

Public Function Repository_updateById(ByVal key As Long, ByRef obj As Object) As Boolean
    If tblmgmt.hasKey(key) Then
        Repository_deleteById key
        Repository_save obj, False
    Else
        Err.Raise errInvalidKey, _
        "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
        "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteById(ByVal key As Long) As Boolean
    If tblmgmt.hasKey(key) Then
        tblmgmt.getListRow(key).Range.delete
    Else
        Err.Raise errInvalidKey, _
        "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
        "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteFkey(ByVal fk As Long, ByVal fkName As String) As Boolean
End Function

Public Function Repository_associacoes(ByVal fk As Long, ByVal fkColName As String) As Collection
    Dim index As Integer
    Dim arr As Variant
    Dim acmpmnt As acompanhamento
    Dim ano As Integer
    Dim tempCollection As New Collection
    
    arr = tblmgmt.getArrayAssociations(fk, fkColName)
    
    For index = 1 To UBound(arr)
        Set acmpmnt = Repository_readByKey(tblmgmt.getCellData(arr(index), "Key"))
        'Debug.Print acmpmnt.toString
        ano = tblmgmt.getCellData(arr(index), "Ano")
        tempCollection.Add acmpmnt, CStr(ano)
    Next index
    
    Set Repository_associacoes = tempCollection
End Function



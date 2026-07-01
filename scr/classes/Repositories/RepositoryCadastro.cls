VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Repository

Dim tblmgmt As New TableManagement
Dim me_repoAssistido As RepositoryAssistido
Dim me_repoAcompanhamento As RepositoryAcompanhamento

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksCADASTRO.ListObjects(1))
End Sub

Public Property Set setRepoAssistiddo(ByRef repoAssistido As Repository)
    Set me_repoAssistido = repoAssistido
End Property

Public Property Set setRepoAcompanhamento(ByRef repoAcompanhamento As Repository)
    Set me_repoAcompanhamento = repoAcompanhamento
End Property

Public Function Repository_readByKey(ByVal key As Long, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim cdst As New cadastro
    
    Set lr = tblmgmt.getListRow(key)
    
    Set cdst = cdst.CreateFromArray(lr.Range.Value)
    
    If complete Then
        Set astd = me_repoAssistido.Repository_readByKey(tblmgmt.getCellData(lr, "KeyAssistido"), True)
        Set cdst.setAssistido = astd
        Set cdst.setAcompanhamentos = me_repoAcompanhamento.Repository_associacoes(cdst.getKey, "KeyCadastro")
    End If
    
    Set Repository_readByKey = cdst
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
    me_repoAssistido.Repository_save obj.getAssistido
    
    tblmgmt.InsertDataRow obj, autoIncrement
    
    Repository_save = True
End Function

Public Function Repository_readById(ByVal key As Long) As Object
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
        me_repoAssistido.Repository_deleteById tblmgmt.getCellData(tblmgmt.getListRow(key), "KeyAssistido")
        tblmgmt.getListRow(key).delete
        Repository_deleteById = True
    Else
        Repository_deleteById = False
        Err.Raise errInvalidKey, _
                  "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
                  "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteFkey(ByVal fk As Long, ByVal fkName As String) As Boolean
End Function

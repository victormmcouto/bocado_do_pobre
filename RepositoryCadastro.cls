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

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksCADASTRO.ListObjects(1))
End Sub

Public Property Set setRepoAssistiddo(ByRef repoAssistido As Repository)
    Set me_repoAssistido = repoAssistido
End Property

Public Function Repository_factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim cdst As New Cadastro
    Dim arrDados() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim arrDados(1 To 1, 1 To lr.Range.Columns.count)
    
    arrDados = lr.Range.Value
    
    Set cdst = cdst.CreateFromArray(arrDados)
    
    If complete Then
        Set astd = me_repoAssistido.Repository_factory(tblmgmt.getCellData(lr, "KeyAssistido"), True)
        Set cdst.setAssistido = astd
    End If
    
    Set Repository_factory = cdst
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
    me_repoAssistido.Repository_save obj.getAssistido
    
    tblmgmt.InsertDataRow obj, autoIncrement
    
    Repository_save = True
End Function

Public Function Repository_readById(ByVal key As Integer) As Object
End Function

Public Function Repository_updateById(ByVal key As Integer, ByRef obj As Object) As Boolean
    If tblmgmt.hasKey(key) Then
        Repository_deleteById key
        Repository_save obj, False
    Else
        Err.Raise errInvalidKey, _
        "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
        "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteById(ByVal key As Integer) As Boolean
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

Public Function Repository_deleteFkey(ByVal fk As Integer, ByVal fkName As String) As Boolean
End Function

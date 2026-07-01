VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryDependente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Repository

Private tblmgmt As New TableManagement
Private me_repoAssistido As RepositoryAssistido

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksDEPENDENTE.ListObjects(1))
End Sub

Public Property Set setRepoAssistido(ByRef repoAssistido As Repository)
    Set me_repoAssistido = repoAssistido
End Property

Public Function Repository_readByKey(ByVal key As Long, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim dpnt As New Dependente
    Dim tempArr() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim tempArr(1 To 1, 1 To lr.Range.Columns.count)
    
    tempArr = lr.Range.Value
    
    If complete Then
        'Dim locInfo As Integer
        Dim keyAstd As Integer
        'Dim arrAsso As Variant
        
        keyAstd = me_repoAssistido.Repository_getTbl.getCellData(lr, "KeyAssistido")
        'locInfo = tblmgmt.getListColumn("KeyAssistido").index
        
        Set dpnt.setAssistido = me_repoAssistido.Repository_readByKey(keyAstd, False)
    End If
    
    Set Repository_readByKey = dpnt.CreateFromArray(tempArr)
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
    tblmgmt.InsertDataRow obj, autoIncrement
    Repository_save = True
End Function

Public Function Repository_updateById(ByVal key As Long, ByRef obj As Object) As Boolean
    If tblmgmt.hasKey(key) Then
        Repository_deleteById key
        If key = 0 Then
            Repository_save obj, True
        Else
            Repository_save obj, False
        End If
    Else
        Err.Raise errInvalidKey, _
        "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
        "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteById(ByVal key As Long) As Boolean
    If tblmgmt.hasKey(key) Then
        tblmgmt.getListRow(key).delete
        Repository_deleteById = True
    Else
        Err.Raise errInvalidKey, _
        "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
        "ID não encontrado!"
        Repository_deleteById = False
    End If
End Function

Public Function Repository_deleteFkey(ByVal fk As Long, ByVal fkName As String) As Boolean
End Function

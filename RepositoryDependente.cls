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

Public Function Repository_factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
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
        
        Set dpnt.setAssistido = me_repoAssistido.Repository_factory(keyAstd, False)
    End If
    
    Set Repository_factory = dpnt.CreateFromArray(tempArr)
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object) As Boolean
    tblmgmt.InsertDataRow obj
    Repository_save = True
End Function

Public Function Repository_readById(ByVal key As Integer) As Object
End Function

Public Function Repository_update(ByRef obj As Object) As Boolean
End Function

Public Function Repository_deleteById(ByVal key As Integer) As Boolean
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

Public Function Repository_deleteFkey(ByVal fk As Integer, ByVal fkName As String) As Boolean
End Function

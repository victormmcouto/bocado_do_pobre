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
Dim repoAssistido As New RepositoryAssistido

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksCADASTRO.ListObjects(1))
End Sub

Public Function Repository_factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
    Dim repoAstd As New RepositoryAssistido
    Dim repoCnjg As New RepositoryConjuge
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim cdst As New Cadastro
    Dim arrDados() As Variant
    
    Set repoAstd = repoAstd.Repository_create(wksASSISTIDO.ListObjects(1))
    Set repoCnjg = repoCnjg.Repository_create(wksCONJUGE.ListObjects(1))
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim arrDados(1 To 1, 1 To lr.Range.Columns.count)
    
    arrDados = lr.Range.Value
    
    Set cdst = cdst.CreateFromArray(arrDados)
    
    If complete Then
        Set astd = repoAstd.Repository_factory(tblmgmt.getCellData(lr, "KeyAssistido"), True)
        Set cdst.setAssistido = astd
    End If
    
    Set Repository_factory = cdst
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object) As Boolean
    repoAssistido.Repository_save obj.getAssistido
    
    tblmgmt.InsertDataRow obj
    
    Repository_save = True
End Function

Public Function Repository_read() As Object
End Function

Public Function Repository_update(ByRef obj As Object) As Boolean
End Function

Public Function Repository_delete(ByRef obj As Object) As Boolean
End Function

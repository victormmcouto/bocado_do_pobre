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

Public Function Repository_Create(ByVal tbl As ListObject) As Repository
    Set tblmgmt = tblmgmt.Create(tbl)
    Set Repository_Create = Me
End Function

Public Function Repository_factory(ByVal key As Integer) As Object
    Dim repoAstd As New RepositoryAssistido
    Dim repoCnjg As New RepositoryConjuge
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim cdst As New Cadastro
    Dim arrDados() As Variant
    
    Set repoAstd = repoAstd.Repository_Create(wksASSISTIDO.ListObjects(1))
    Set repoCnjg = repoCnjg.Repository_Create(wksCONJUGE.ListObjects(1))
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim arrDados(1 To 1, 1 To lr.Range.Columns.count)
    
    arrDados = lr.Range.Value
    
    Set astd = repoAstd.Repository_factory(tblmgmt.getCellData(lr, "KeyAssistido"))
    Set cdst = cdst.CreateFromArray(arrDados)
    Set cdst.setAssistido = astd
    
    Set Repository_factory = cdst
End Function

Public Function Repository_getTbl() As TableManagement
End Function


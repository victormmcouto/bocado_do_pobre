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

Public Function Repository_Create(ByVal tbl As ListObject) As Repository
    Set tblmgmt = tblmgmt.Create(tbl)
    Set Repository_Create = Me
End Function

Public Function Repository_factory(ByVal key As Integer) As Object
    Dim lr As ListRow
    Dim dpnt As New Dependente
    Dim tempArr() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim tempArr(1 To 1, 1 To lr.Range.Columns.count)
    
    tempArr = lr.Range.Value
    
    Dim repoAstd As New RepositoryAssistido
    Dim astd As New Assistido
    Dim locInfo As Integer
    Dim keyAstd As Integer
    Dim arrAsso As Variant
    
    Set repoAstd = repoAstd.Repository_Create(wksASSISTIDO.ListObjects(1))
    
    keyAstd = repoAstd.Repository_getTbl.getCellData(lr, "KeyAssistido")
    locInfo = tblmgmt.getListColumn("KeyAssistido").index
    
    Set dpnt.setAssistido = repoAstd.Repository_factory(keyAstd)
        
    Set Repository_factory = dpnt.CreateFromArray(tempArr)
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function



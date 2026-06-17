VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryAssistido"
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

Public Function Repository_factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim tempArr() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim tempArr(1 To 1, 1 To lr.Range.Columns.count)
    
    tempArr = lr.Range.Value
    
    Set astd = astd.CreateFromArray(tempArr)
    
    If complete Then
        Dim qnt As Integer
        Dim arrAsso() As Variant
        Dim repoDpnt As New RepositoryDependente
        
        Set repoDpnt = repoDpnt.Repository_Create(wksDEPENDENTE.ListObjects(1))
        
        arrAsso = repoDpnt.Repository_getTbl.getArrayAssociations(tblmgmt.getCellData(lr, "Key"), "KeyAssistido")
        
        If UBound(arrAsso) <> 0 Then
            For qnt = 1 To UBound(arrAsso)
                Dim dpnt As New Dependente
                
                Set dpnt = repoDpnt.Repository_factory(repoDpnt.Repository_getTbl.getCellData(arrAsso(qnt), "Key"), False)
                astd.getDependentes.Add dpnt
                
                Set dpnt = Nothing
            Next qnt
        End If
        
        Dim repoCnjg As New RepositoryConjuge
        Dim keyCnjg As Integer
        Set repoCnjg = repoCnjg.Repository_Create(wksCONJUGE.ListObjects(1))
        keyCnjg = tblmgmt.getCellData(lr, "KeyConjuge")
        If keyCnjg <> 0 Then Set astd.setConjuge = repoCnjg.Repository_factory(keyCnjg, False)
    End If
    
    Set Repository_factory = astd
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryConjuge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Repository

Private tblmgmt As New TableManagement

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksCONJUGE.ListObjects(1))
End Sub

Public Function Repository_factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim cnjg As New Conjuge
    Dim tempArr() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim tempArr(1 To 1, 1 To lr.Range.Columns.count)
    
    tempArr = lr.Range.Value
    
    Set cnjg = cnjg.CreateFromArray(tempArr)
    
    Set Repository_factory = cnjg
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object) As Boolean
    tblmgmt.InsertDataRow obj
    Repository_save = True
End Function

Public Function Repository_read() As Object
End Function

Public Function Repository_update(ByRef obj As Object) As Boolean
End Function

Public Function Repository_delete(ByRef obj As Object) As Boolean
End Function

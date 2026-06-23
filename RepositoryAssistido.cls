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
Private repoCnjg As New RepositoryConjuge
Private repoDpnt As New RepositoryDependente

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksASSISTIDO.ListObjects(1))
End Sub

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
        
        Set repoDpnt = repoDpnt.Repository_create(wksDEPENDENTE.ListObjects(1))
        
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
        Set repoCnjg = repoCnjg.Repository_create(wksCONJUGE.ListObjects(1))
        keyCnjg = tblmgmt.getCellData(lr, "KeyConjuge")
        If keyCnjg <> 0 Then Set astd.setConjuge = repoCnjg.Repository_factory(keyCnjg, False)
    End If
    
    Set Repository_factory = astd
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object) As Boolean
    If Not obj.getConjuge Is Nothing Then
        repoCnjg.Repository_save obj.getConjuge
    End If
    
    tblmgmt.InsertDataRow obj
        
    If obj.getTotalDependentes > 0 Then
        Dim index As Integer
        
        For index = 1 To obj.getTotalDependentes
            repoDpnt.Repository_save obj.getDependente(index)
        Next index
    End If
    
    Repository_save = True
End Function

Public Function Repository_read() As Object
End Function

Public Function Repository_update(ByRef obj As Object) As Boolean
End Function

Public Function Repository_delete(ByRef obj As Object) As Boolean
End Function

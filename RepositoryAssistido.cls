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

Public Function Repository_factory(ByVal key As Integer) As Object
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim tempArr() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim tempArr(1 To 1, 1 To lr.Range.Columns.count)
    
    tempArr = lr.Range.Value
    
    Set astd = astd.CreateFromArray(tempArr)

    Dim qnt As Integer
    Dim arrAsso() As Variant
    Dim tblmgmtDpnt As New TableManagement
    
    Set tblmgmtDpnt = tblmgmtDpnt.Create(wksDEPENDENTE.ListObjects(1))
    
    arrAsso = tblmgmtDpnt.getArrayAssociations(tblmgmt.getCellData(lr, "Key"), "KeyAssistido")
    
    If UBound(arrAsso) <> 0 Then
        For qnt = 1 To UBound(arrAsso)
            Dim dpnt As New Dependente
            
            ReDim tempArr(1 To 1, 1 To UBound(arrAsso))
            tempArr = arrAsso(qnt).Range.Value
            
            With astd.getDependentes
                 Set dpnt.setAssistido = astd
                .Add dpnt.CreateFromArray(tempArr)
            End With
            
            Set dpnt = Nothing
        Next qnt
    End If
    
    Dim tblmgmtCnjg As New TableManagement
    
    Set tblmgmtCnjg = tblmgmtCnjg.Create(wksCONJUGE.ListObjects(1))
    
    arrAsso = tblmgmtCnjg.getArrayAssociations(tblmgmt.getCellData(lr, "KeyConjuge"), "Key")
    
    If UBound(arrAsso) <> 0 Then
        ReDim tempArr(1 To 1, 1 To arrAsso(1).Range.count)
        Dim cnjg As New Conjuge
        
        tempArr = arrAsso(1).Range.Value
        
        cnjg.CreateFromArray tempArr
        Set astd.setConjuge = cnjg
    End If
    
    Set Repository_factory = astd
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

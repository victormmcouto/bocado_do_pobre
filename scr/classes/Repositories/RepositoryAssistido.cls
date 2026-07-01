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
Private me_repoCnjg As RepositoryConjuge
Private me_repoDpnt As RepositoryDependente

Private Sub Class_Initialize()
    Set tblmgmt = tblmgmt.create(wksASSISTIDO.ListObjects(1))
End Sub

Public Property Set setRepoConjuge(ByRef repoCnjg As Repository)
    Set me_repoCnjg = repoCnjg
End Property

Public Property Set setRepoDependente(ByRef repoDpnt As Repository)
    Set me_repoDpnt = repoDpnt
End Property

Public Function Repository_readByKey(ByVal key As Long, Optional ByVal complete As Boolean = True) As Object
    Dim lr As ListRow
    Dim astd As New Assistido
    Dim tempArr() As Variant
    
    Set lr = tblmgmt.getListRow(key)
    
    ReDim tempArr(1 To 1, 1 To lr.Range.Columns.count)
    
    tempArr = lr.Range.Value
    
    Set astd = astd.CreateFromArray(tempArr)
    
    If complete Then
        Dim qnt As Integer
        Dim arrAsso As Variant
        
        arrAsso = me_repoDpnt.Repository_getTbl.getArrayAssociations(tblmgmt.getCellData(lr, "Key"), "KeyAssistido")
        
        If UBound(arrAsso) <> 0 Then
            For qnt = 1 To UBound(arrAsso)
                Dim dpnt As New Dependente
                
                Set dpnt = me_repoDpnt.Repository_readByKey(me_repoDpnt.Repository_getTbl.getCellData(arrAsso(qnt), "Key"), False)
                astd.AddDependente dpnt
                
                Set dpnt = Nothing
            Next qnt
        End If
        
        Dim keyCnjg As Long
        keyCnjg = tblmgmt.getCellData(lr, "KeyConjuge")
        If keyCnjg <> 0 Then Set astd.setConjuge = me_repoCnjg.Repository_readByKey(keyCnjg, False)
    End If
    
    Set Repository_readByKey = astd
End Function

Public Function Repository_getTbl() As TableManagement
    Set Repository_getTbl = tblmgmt
End Function

Public Function Repository_save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
    If Not obj.getConjuge Is Nothing Then
        me_repoCnjg.Repository_save obj.getConjuge, autoIncrement
    End If
    
    tblmgmt.InsertDataRow obj, autoIncrement
        
    If obj.getTotalDependentes > 0 Then
        Dim index As Integer
        
        For index = 1 To obj.getTotalDependentes
            If obj.getDependente(index).getKey = 0 Then
                me_repoDpnt.Repository_save obj.getDependente(index), True
            Else
                me_repoDpnt.Repository_save obj.getDependente(index), False
            End If
        Next index
    End If
    
    Repository_save = True
End Function

Public Function Repository_readById(ByVal key As Long) As Object
End Function

Public Function Repository_updateById(ByVal key As Long, ByRef obj As Object) As Boolean
    If tblmgmt.hasKey(key) Then
        Repository_deleteById key
        Repository_save obj, False
    Else
        Err.Raise errInvalidKey, _
        "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
        "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteById(ByVal key As Long) As Boolean
    If tblmgmt.hasKey(key) Then
        Dim arrAssociation As Variant
        
        arrAssociation = me_repoDpnt.Repository_getTbl.getArrayAssociations(key, "KeyAssistido")
    
        If UBound(arrAssociation) > 0 Then
            Dim lrIndex As Integer
            
            For lrIndex = UBound(arrAssociation) To 1 Step -1
                arrAssociation(lrIndex).delete
            Next lrIndex
        End If
        
        Dim keyCnjg As Long
        
        keyCnjg = tblmgmt.getCellData(tblmgmt.getListRow(key), "KeyConjuge")
        
        If keyCnjg <> 0 Then
            me_repoCnjg.Repository_deleteById keyCnjg
        End If
        
        tblmgmt.getListRow(key).delete
        
        Repository_deleteById = True
    Else
        Repository_deleteById = False
        Err.Raise errInvalidKey, _
                  "RepositoryDependente.Repository_delebyId(key:=" & key & ")", _
                  "ID não encontrado!"
    End If
End Function

Public Function Repository_deleteFkey(ByVal fk As Long, ByVal fkName As String) As Boolean
    tblmgmt.setCellData tblmgmt.getListRow(fk, fkName), fkName, 0
End Function

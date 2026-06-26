VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private me_tbl As ListObject

Public Sub StartTable()
    On Error Resume Next
    Dim rows As Integer
    rows = me_tbl.ListRows.count
    If Err.Number <> 0 Then AddRows 1
    On Error GoTo 0
End Sub

Public Function create(ByRef tbl As ListObject) As TableManagement
    Set me_tbl = tbl
    StartTable
    Set create = Me
End Function

Public Property Get Repository_getTbl() As ListObject
    Set Repository_getTbl = me_tbl
End Property

Public Function getListColumn(ByVal index As Variant) As ListColumn
    Set getListColumn = me_tbl.ListColumns(index)
End Function

Public Function getListRow(ByVal key As Integer, Optional keyColName As String = "Key") As ListRow
    Dim lr As ListRow
    Dim current_key As Integer
    
    If hasKey(key) Then
        With me_tbl
            Set getListRow = .ListRows.item(getListColumn(keyColName).DataBodyRange.Find(key, LookAt:=xlWhole).Row - 1)
            'getListRow.Range.Select
        End With
    Else
        Err.Raise errInvalidKey, "TableManagement.getListRow(key:=" & key & ")", "Chave inválida!"
    End If
End Function

Public Property Get getLastRow() As Range
    With me_tbl.ListRows
        Set getLastRow = .item(.count).Range
    End With
End Property

Public Sub AddRows(ByVal qnt As Integer)
    Dim index As Integer
    
    For index = 1 To qnt
        me_tbl.ListRows.Add
    Next index
End Sub

Public Function LastRowIsEmpty() As Boolean
    With Application.WorksheetFunction
        If .CountBlank(getLastRow) = getLastRow.count Then
            LastRowIsEmpty = True
        Else
            LastRowIsEmpty = False
        End If
    End With
End Function

Public Property Get getTotalRows() As Integer
    getTotalRows = me_tbl.ListRows.count
End Property

Public Property Get getTotalColumns() As Integer
    getTotalColumns = me_tbl.ListColumns.count
End Property

Public Sub InsertDataRow(ByRef obj As Object, ByVal autoIncrement As Boolean)
    If Not LastRowIsEmpty Then AddRows 1
    With getLastRow
        If autoIncrement Then obj.setKey = getLastKey + 1
        .Value = obj.toArray
    End With
End Sub

Public Function getLastKey() As Integer
    Dim lastKey As Integer
    
    lastKey = Application.WorksheetFunction.Max(me_tbl.ListColumns(1).DataBodyRange)
    
    getLastKey = lastKey
End Function

Public Function getCellData(ByVal lr As ListRow, ByVal colNAme As String) As Variant
    getCellData = lr.Range(1, lr.Parent.ListColumns(colNAme).index).Value
End Function

Public Function setCellData(ByVal lr As ListRow, ByVal colNAme As String, ByVal newValue As Variant) As Variant
    lr.Range(1, lr.Parent.ListColumns(colNAme).index).Value = newValue
End Function

Public Function getArrayAssociations(ByVal fkey As Integer, ByVal colNAme As String) As Variant
    Dim tempArr() As ListRow
    Dim count As Integer
    Dim lr As ListRow
    
    count = 1
    
    'me_tbl.DataBodyRange.Select
    
    For Each lr In me_tbl.ListRows
        If getCellData(lr, colNAme) = fkey Then
            ReDim Preserve tempArr(1 To count)
            Set tempArr(count) = lr
            count = count + 1
        End If
    Next lr
    
    getArrayAssociations = tempArr
End Function

Public Function hasKey(ByVal key As Integer) As Boolean
    hasKey = Not getListColumn("Key").DataBodyRange.Find(key) Is Nothing
End Function

Public Sub updateRow(ByVal key As Integer, ByRef obj As Object)
    getListRow(key).Range.Value = obj.toArray()
End Sub

Public Sub deleteAssosciations(ByVal key As Integer, ByVal colNAme As String)
    Dim index As Integer
    Dim assosciations As Variant
    
    assosciations = getArrayAssociations(key, colNAme)
    
    For index = UBound(assosciations) To 1 Step -1
        assosciations(index).delete
    Next index
End Sub

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

Public Function Create(ByRef tbl As ListObject) As TableManagement
    Set me_tbl = tbl
    StartTable
    Set Create = Me
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
    
    For Each lr In me_tbl.ListRows
        current_key = CInt(lr.Range(1, me_tbl.ListColumns.item("Key").index).Value)
        If current_key = key Then
            Set getListRow = lr
            Exit Function
        End If
    Next lr
    
    Err.Raise errInvalidKey, "TableManagement.getListRow(key=" & key & ")", "Chave inválida!"
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

Public Sub InsertDataRow(ByRef obj As Object)
    If Not LastRowIsEmpty Then AddRows 1
    With getLastRow
        obj.setKey = getLastKey + 1
        .Value = obj.toArray
    End With
End Sub

Public Function getLastKey() As Integer
    Dim lastKey As Integer
    
    lastKey = Application.WorksheetFunction.Max(me_tbl.ListColumns(1).DataBodyRange)
    
    getLastKey = lastKey
End Function

Public Function getCellData(ByVal lr As ListRow, ByVal colName As String) As Variant
    getCellData = lr.Range(1, lr.Parent.ListColumns(colName).index).Value
End Function

Public Function getArrayAssociations(ByVal fkey As Integer, ByVal colName As String) As Variant
    Dim tempArr() As Variant
    Dim count As Integer
    Dim lr As ListRow
    
    count = 1
    
    For Each lr In me_tbl.ListRows
        If getCellData(lr, colName) = fkey Then
            ReDim Preserve tempArr(1 To count)
            Set tempArr(count) = lr
            count = count + 1
        End If
    Next lr
    
    getArrayAssociations = tempArr
End Function


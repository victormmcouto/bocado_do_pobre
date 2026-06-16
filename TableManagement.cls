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

Public Function Create(ByRef tbl) As TableManagement
    Set me_tbl = tbl
    Set Create = Me
End Function

Public Property Get getTbl() As ListObject
    Set getTbl = me_tbl
End Property

Public Property Get getListRow(ByVal key As Integer) As ListRow
    Dim lr As ListRow
    Dim current_key As Integer
    
    For Each lr In me_tbl.ListRows
        current_key = CInt(lr.Range(1, me_tbl.ListColumns.item(1).index).Value)
        If current_key = key Then
            Set getListRow = lr
            Exit Property
        End If
    Next lr
End Property

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

Public Sub InsertData(ByVal obj As Object)
    Select Case obj
        Case TypeName("Assistido"):
            
    End Select
End Sub


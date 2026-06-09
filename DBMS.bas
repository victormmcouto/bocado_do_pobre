Attribute VB_Name = "DBMS"
Option Explicit

Private tblCadastros As ListObject
Private tblParentes As ListObject

Public Sub RealizarCadastro()
    Dim index As Integer
    Dim totalRows As Integer
    Dim keyParenteAnterior As Integer
    
    InitializeTables
    
    With tblParentes
        keyParenteAnterior = .ListColumns("KeyParente").DataBodyRange.Cells(.DataBodyRange.Rows.count).Value
        
        For index = 1 To UBound(cadastro.parentes)
            AddRow tblParentes
            
            totalRows = .DataBodyRange.Rows.count
            
            With cadastro.parentes(index)
                tblParentes.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = keyParenteAnterior + 1
                tblParentes.ListColumns("NomeParente").DataBodyRange.Cells(totalRows).Value = .Nome
                tblParentes.ListColumns("GrauParentescoParente").DataBodyRange.Cells(totalRows).Value = .GrauParentesco
                tblParentes.ListColumns("DataNascimentoParente").DataBodyRange.Cells(totalRows).Value = .DataNascimento
                tblParentes.ListColumns("EscolaridadeParente").DataBodyRange.Cells(totalRows).Value = .Escolaridade
            End With
        Next index
    End With
    
    AddRow tblCadastros
End Sub

Private Sub InitializeTables()
    Dim totalRows As Integer
    If tblCadastros Is Nothing Then
        Set tblCadastros = wksCADASTROS.ListObjects(1)
    End If
    If tblParentes Is Nothing Then
        Set tblParentes = wksPARENTES.ListObjects(1)
    End If
        
    On Error Resume Next
    With tblCadastros
        totalRows = .DataBodyRange.Rows.count
        If Err.Number <> 0 Then .ListRows.Add
    End With
    On Error GoTo 0
    On Error Resume Next
    With tblParentes
        totalRows = .DataBodyRange.Rows.count
        If Err.Number <> 0 Then .ListRows.Add
    End With
    On Error GoTo 0
End Sub

Private Sub AddRow(ByRef tbl As ListObject)
    With tbl.ListRows
        If Application.WorksheetFunction.CountBlank(.item(.count).Range.Cells) <> .item(.count).Range.Cells.count Then
            .Add
        End If
    End With
End Sub

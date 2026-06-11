Attribute VB_Name = "FuncoesWks"
Option Explicit

Public Function PROCINDICE(ByVal rngDados As Range, _
                           ByVal target As String, _
                           ByVal colName As String, _
                           ByVal ocorrencia As Integer) As String
    Dim tbl As ListObject
    Dim strNomePrimeiraCol As String
    Dim lr As ListRow
    Dim count As Integer
    Dim i As Long
    
    Set tbl = rngDados.ListObject
    strNomePrimeiraCol = tbl.HeaderRowRange.Cells(1, 1).Value
    
    count = 0
    
    For Each lr In tbl.ListRows
        If lr.Range(1, tbl.ListColumns(strNomePrimeiraCol).index).Value = target Then
            count = count + 1
            If count = ocorrencia Then
                PROCINDICE = lr.Range(1, tbl.ListColumns(colName).index).Value
                Exit Function
            End If
        End If
    Next lr
    
    PROCINDICE = "------"  ' não encontrou a ocorrência
End Function

Public Function RemoverEspacos(str As String) As String
    RemoverEspacos = "31" & Replace(str, " ", "")
End Function

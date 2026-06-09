Attribute VB_Name = "DBMS"
Option Explicit

Private tblCadastros As ListObject
Private tblParentes As ListObject

Public Sub RealizarCadastro()
    Dim index As Integer
    Dim totalRows As Integer
    Dim keyParentesAnterior As Integer
    
    InitializeTables
    
    With tblParentes
        keyParentesAnterior = .ListColumns("KeyParente").DataBodyRange.Cells(.DataBodyRange.Rows.count).Value
        
        For index = 1 To UBound(cadastro.parentes)
            AddRow tblParentes
            
            totalRows = .DataBodyRange.Rows.count
            
            With cadastro.parentes(index)
                tblParentes.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = keyParentesAnterior + 1
                tblParentes.ListColumns("NomeParente").DataBodyRange.Cells(totalRows).Value = .Nome
                tblParentes.ListColumns("GrauParentescoParente").DataBodyRange.Cells(totalRows).Value = .GrauParentesco
                tblParentes.ListColumns("DataNascimentoParente").DataBodyRange.Cells(totalRows).Value = .DataNascimento
                tblParentes.ListColumns("EscolaridadeParente").DataBodyRange.Cells(totalRows).Value = .Escolaridade
            End With
        Next index
    End With
    
    With cadastro
        AddRow tblCadastros
        
        totalRows = tblCadastros.DataBodyRange.Rows.count
        
        With .Assistido
            tblCadastros.ListColumns("NomeAssistido").DataBodyRange.Cells(totalRows).Value = .Nome
            tblCadastros.ListColumns("DataNascimentoAssistido").DataBodyRange.Cells(totalRows).Value = .DataNascimento
            tblCadastros.ListColumns("EstadoCivilAssistido").DataBodyRange.Cells(totalRows).Value = .EstadoCivil
            tblCadastros.ListColumns("ProfissaoAssistido").DataBodyRange.Cells(totalRows).Value = .Profissao
            tblCadastros.ListColumns("EscolaridadeAssistido").DataBodyRange.Cells(totalRows).Value = .Escolaridade
            tblCadastros.ListColumns("CPFAssistido").DataBodyRange.Cells(totalRows).Value = .CPF
            tblCadastros.ListColumns("TelefoneAssistido").DataBodyRange.Cells(totalRows).Value = .Telefone
        End With
        
        With .Conjugue
            tblCadastros.ListColumns("NomeConjugue").DataBodyRange.Cells(totalRows).Value = .Nome
            tblCadastros.ListColumns("DataNascimentoConjugue").DataBodyRange.Cells(totalRows).Value = .DataNascimento
            tblCadastros.ListColumns("EstadoCivilConjugue").DataBodyRange.Cells(totalRows).Value = .EstadoCivil
            tblCadastros.ListColumns("ProfissaoConjugue").DataBodyRange.Cells(totalRows).Value = .Profissao
            tblCadastros.ListColumns("EscolaridadeConjugue").DataBodyRange.Cells(totalRows).Value = .Escolaridade
            tblCadastros.ListColumns("CPFConjugue").DataBodyRange.Cells(totalRows).Value = .CPF
            tblCadastros.ListColumns("TelefoneConjugue").DataBodyRange.Cells(totalRows).Value = .Telefone
        End With
        
        With .DemaisInfo
            tblCadastros.ListColumns("DataSindicancia").DataBodyRange.Cells(totalRows).Value = .DataSindicancia
            tblCadastros.ListColumns("NomeVisitador").DataBodyRange.Cells(totalRows).Value = .NomeVisitador
            tblCadastros.ListColumns("ParticipaProgGov?").DataBodyRange.Cells(totalRows).Value = .ParticipaProgramaGov
            tblCadastros.ListColumns("ProgGov").DataBodyRange.Cells(totalRows).Value = .ProgramaGov
            tblCadastros.ListColumns("TipoMoradia").DataBodyRange.Cells(totalRows).Value = .TipoMoradia
            tblCadastros.ListColumns("NPessoas").DataBodyRange.Cells(totalRows).Value = .NPessoasNaCasa
            tblCadastros.ListColumns("RecebeCesta?").DataBodyRange.Cells(totalRows).Value = .RecebeCesta
        End With
        
        With .Endereco
            tblCadastros.ListColumns("Logradouro").DataBodyRange.Cells(totalRows).Value = .Logradouro
            tblCadastros.ListColumns("Numero").DataBodyRange.Cells(totalRows).Value = .NumeroCasa
            tblCadastros.ListColumns("Bairro").DataBodyRange.Cells(totalRows).Value = .Bairro
            tblCadastros.ListColumns("Cidade").DataBodyRange.Cells(totalRows).Value = .Cidade
        End With
    End With
    
    tblCadastros.ListColumns("KeyParentes").DataBodyRange.Cells(totalRows).Value = keyParentesAnterior + 1
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

Attribute VB_Name = "DBMS"
Option Explicit

Private tblCadastros As ListObject
Private tblParentes As ListObject

Public Sub RealizarCadastro()
    Dim index As Integer
    Dim totalRows As Integer
    Dim keyParentesAtual As Integer
    
    InitializeTables
    
    If ParentesInicializado() Then
        With tblParentes
            If Not boolCadastrar Then
                keyParentesAtual = CInt(keyParenteAntesDeletar)
            Else
                keyParentesAtual = .ListColumns("KeyParente").DataBodyRange.Cells(.DataBodyRange.Rows.Count).Value + 1
            End If
            
            For index = 1 To UBound(Cadastro.parentes)
                AddRow tblParentes
                
                totalRows = .DataBodyRange.Rows.Count
                
                With Cadastro.parentes(index)
                    tblParentes.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = keyParentesAtual
                    tblParentes.ListColumns("NomeParente").DataBodyRange.Cells(totalRows).Value = .Nome
                    tblParentes.ListColumns("GrauParentescoParente").DataBodyRange.Cells(totalRows).Value = .GrauParentesco
                    tblParentes.ListColumns("DataNascimentoParente").DataBodyRange.Cells(totalRows).Value = .DataNascimento
                    tblParentes.ListColumns("EscolaridadeParente").DataBodyRange.Cells(totalRows).Value = .Escolaridade
                End With
            Next index
        End With
    End If
    
    With Cadastro
        AddRow tblCadastros
        
        totalRows = tblCadastros.DataBodyRange.Rows.Count
        
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
    
    If keyParentesAtual = 0 Then
        tblCadastros.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = "-"
    Else
        tblCadastros.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = keyParentesAtual
    End If
End Sub

Public Sub PopulateTypeCadastro()
    Dim totalParentes As Integer
    
    InitializeTables
        
    totalParentes = CInt(listRowComparativo.Range(1, tblCadastros.ListColumns("NPessoas").index).Value)

    With Cadastro
        If totalParentes <> 0 Then
            Dim index As Integer
            Dim keyParentes As Integer
            Dim lr As ListRow
            Dim arrListRowParentes() As ListRow
            
            ReDim arrListRowParentes(1 To totalParentes)
            ReDim Cadastro.parentes(1 To totalParentes)
            
            index = 1
            
            keyParentes = CInt(listRowComparativo.Range(1, tblCadastros.ListColumns("KeyParente").index).Value)
            
            For Each lr In tblParentes.ListRows
                If lr.Range(1, tblParentes.ListColumns("KeyParente").index).Value = keyParentes Then
                    Set arrListRowParentes(index) = lr
                    index = index + 1
                End If
            Next lr
            
            For index = 1 To totalParentes
                With arrListRowParentes(index)
                    Cadastro.parentes(index).Nome = .Range(1, tblParentes.ListColumns("NomeParente").index).Value
                    Cadastro.parentes(index).GrauParentesco = .Range(1, tblParentes.ListColumns("GrauParentescoParente").index).Value
                    Cadastro.parentes(index).Escolaridade = .Range(1, tblParentes.ListColumns("EscolaridadeParente").index).Value
                    On Error Resume Next
                    Cadastro.parentes(index).DataNascimento = .Range(1, tblParentes.ListColumns("DataNascimentoParente").index).Value
                    On Error GoTo 0
                End With
            Next index
        End If
        
        With .Assistido
            .Nome = listRowComparativo.Range(1, tblCadastros.ListColumns("NomeAssistido").index).Value
            On Error Resume Next
            .DataNascimento = listRowComparativo.Range(1, tblCadastros.ListColumns("DataNascimentoAssistido").index).Value
            On Error GoTo 0
            .EstadoCivil = listRowComparativo.Range(1, tblCadastros.ListColumns("EstadoCivilAssistido").index).Value
            .Profissao = listRowComparativo.Range(1, tblCadastros.ListColumns("ProfissaoAssistido").index).Value
            .Escolaridade = listRowComparativo.Range(1, tblCadastros.ListColumns("EscolaridadeAssistido").index).Value
            .CPF = listRowComparativo.Range(1, tblCadastros.ListColumns("CPFAssistido").index).Value
            .Telefone = listRowComparativo.Range(1, tblCadastros.ListColumns("TelefoneAssistido").index).Value
        End With
        With .Conjugue
            .Nome = listRowComparativo.Range(1, tblCadastros.ListColumns("NomeConjugue").index).Value
            On Error Resume Next
            .DataNascimento = listRowComparativo.Range(1, tblCadastros.ListColumns("DataNascimentoConjugue").index).Value
            On Error GoTo 0
            .EstadoCivil = listRowComparativo.Range(1, tblCadastros.ListColumns("EstadoCivilConjugue").index).Value
            .Profissao = listRowComparativo.Range(1, tblCadastros.ListColumns("ProfissaoConjugue").index).Value
            .Escolaridade = listRowComparativo.Range(1, tblCadastros.ListColumns("EscolaridadeConjugue").index).Value
            .CPF = listRowComparativo.Range(1, tblCadastros.ListColumns("CPFConjugue").index).Value
            .Telefone = listRowComparativo.Range(1, tblCadastros.ListColumns("TelefoneConjugue").index).Value
        End With
        With .DemaisInfo
            On Error Resume Next
            .DataSindicancia = listRowComparativo.Range(1, tblCadastros.ListColumns("DataSindicancia").index).Value
            On Error GoTo 0
            .NomeVisitador = listRowComparativo.Range(1, tblCadastros.ListColumns("NomeVisitador").index).Value
            .ParticipaProgramaGov = CBool(listRowComparativo.Range(1, tblCadastros.ListColumns("ParticipaProgGov?").index).Value)
            .ProgramaGov = listRowComparativo.Range(1, tblCadastros.ListColumns("ProgGov").index).Value
            .TipoMoradia = listRowComparativo.Range(1, tblCadastros.ListColumns("TipoMoradia").index).Value
            .NPessoasNaCasa = listRowComparativo.Range(1, tblCadastros.ListColumns("NPessoas").index).Value
            .RecebeCesta = CBool(listRowComparativo.Range(1, tblCadastros.ListColumns("RecebeCesta?").index).Value)
        End With
        With .Endereco
            .Logradouro = listRowComparativo.Range(1, tblCadastros.ListColumns("Logradouro").index).Value
            .NumeroCasa = listRowComparativo.Range(1, tblCadastros.ListColumns("Numero").index).Value
            .Bairro = listRowComparativo.Range(1, tblCadastros.ListColumns("Bairro").index).Value
            .Cidade = listRowComparativo.Range(1, tblCadastros.ListColumns("Cidade").index).Value
        End With
    End With
End Sub

Public Sub DeletarCadastro()
    InitializeTables

    Dim numParentes As Integer
    
    numParentes = CInt(listRowComparativo.Range(1, tblCadastros.ListColumns("NPessoas").index).Value)

    keyParenteAntesDeletar = listRowComparativo.Range(1, tblCadastros.ListColumns("KeyParente").index).Value
    listRowComparativo.Delete
    
    If numParentes > 0 Then
        Dim index As Integer
        
        With tblParentes
            For index = .ListRows.Count To 1 Step -1
                If .ListRows(index).Range(1, .ListColumns("KeyParente").index).Value = keyParenteAntesDeletar Then
                    .ListRows(index).Delete
                End If
            Next index
        End With
    End If
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
        totalRows = .DataBodyRange.Rows.Count
        If Err.Number <> 0 Then .ListRows.Add
    End With
    On Error GoTo 0
    On Error Resume Next
    With tblParentes
        totalRows = .DataBodyRange.Rows.Count
        If Err.Number <> 0 Then .ListRows.Add
    End With
    On Error GoTo 0
End Sub

Private Sub AddRow(ByRef tbl As ListObject)
    With tbl.ListRows
        If Application.WorksheetFunction.CountBlank(.item(.Count).Range.Cells) <> .item(.Count).Range.Cells.Count Then
            .Add
        End If
    End With
End Sub

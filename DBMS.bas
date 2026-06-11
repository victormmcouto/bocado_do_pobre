Attribute VB_Name = "DBMS"
Option Explicit

Private tblCadastros As ListObject
Private tblParentes As ListObject
Private tblAcompanhamento As ListObject

Public Sub RealizarCadastro()
    Dim index As Integer
    Dim totalRows As Integer
    Dim keyParentesAtual As Integer
    Dim keyAcompanhamentoAtual As Integer
    
    InitializeTables
    
    If ParentesInicializado() Then
        With tblParentes
            If Not boolCadastrar Then
                keyParentesAtual = CInt(keyParenteAntesDeletar)
            Else
                keyParentesAtual = .ListColumns("KeyParente").DataBodyRange.Cells(.DataBodyRange.Rows.count).Value + 1
            End If
            
            For index = 1 To UBound(Cadastro.parentes)
                AddRow tblParentes
                
                totalRows = .DataBodyRange.Rows.count
                
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
        
        With .conjuge
            tblCadastros.ListColumns("Nomeconjuge").DataBodyRange.Cells(totalRows).Value = .Nome
            tblCadastros.ListColumns("DataNascimentoconjuge").DataBodyRange.Cells(totalRows).Value = .DataNascimento
            tblCadastros.ListColumns("EstadoCivilconjuge").DataBodyRange.Cells(totalRows).Value = .EstadoCivil
            tblCadastros.ListColumns("Profissaoconjuge").DataBodyRange.Cells(totalRows).Value = .Profissao
            tblCadastros.ListColumns("Escolaridadeconjuge").DataBodyRange.Cells(totalRows).Value = .Escolaridade
            tblCadastros.ListColumns("CPFconjuge").DataBodyRange.Cells(totalRows).Value = .CPF
            tblCadastros.ListColumns("Telefoneconjuge").DataBodyRange.Cells(totalRows).Value = .Telefone
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
        
        With .Acompanhamento
            Dim lr As ListRow
            
            With tblAcompanhamento
                If Not boolCadastrar Then
                    keyAcompanhamentoAtual = CInt(keyAcompanhamentoAntesDeletar)
                Else
                    keyAcompanhamentoAtual = .ListColumns("KeyAcompanhamento").DataBodyRange.Cells(.DataBodyRange.Rows.count).Value + 1
                End If
            End With
            
            Set lr = AddRow(tblAcompanhamento)
            
            SetTableCellValue lr, "KeyAcompanhamento", keyAcompanhamentoAtual
            SetTableCellValue lr, "Jan", .Jan
            SetTableCellValue lr, "Fev", .Fev
            SetTableCellValue lr, "Mar", .Mar
            SetTableCellValue lr, "Abr", .Abr
            SetTableCellValue lr, "Mai", .Mai
            SetTableCellValue lr, "Jun", .Jun
            SetTableCellValue lr, "Jul", .Jul
            SetTableCellValue lr, "Ago", .Ago
            SetTableCellValue lr, "Set", .Set
            SetTableCellValue lr, "Out", .Out
            SetTableCellValue lr, "Nov", .Nov
            SetTableCellValue lr, "Dez", .Dez
        End With
    End With
    
    If keyParentesAtual = 0 Then
        tblCadastros.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = "-"
    Else
        tblCadastros.ListColumns("KeyParente").DataBodyRange.Cells(totalRows).Value = keyParentesAtual
    End If
    
    tblCadastros.ListColumns("KeyAcompanhamento").DataBodyRange.Cells(totalRows).Value = keyAcompanhamentoAtual
End Sub

Public Sub PopulateTypeCadastro()
    Dim totalParentes As Integer
    Dim keyAcompanhamento As String
    Dim lr As ListRow
    
    InitializeTables
        
    totalParentes = CInt(listRowComparativo.Range(1, tblCadastros.ListColumns("NPessoas").index).Value)

    With Cadastro
        If totalParentes <> 0 Then
            Dim index As Integer
            Dim keyParentes As Integer
            'Dim lr As ListRow
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
        With .conjuge
            .Nome = listRowComparativo.Range(1, tblCadastros.ListColumns("Nomeconjuge").index).Value
            On Error Resume Next
            .DataNascimento = listRowComparativo.Range(1, tblCadastros.ListColumns("DataNascimentoconjuge").index).Value
            On Error GoTo 0
            .EstadoCivil = listRowComparativo.Range(1, tblCadastros.ListColumns("EstadoCivilconjuge").index).Value
            .Profissao = listRowComparativo.Range(1, tblCadastros.ListColumns("Profissaoconjuge").index).Value
            .Escolaridade = listRowComparativo.Range(1, tblCadastros.ListColumns("Escolaridadeconjuge").index).Value
            .CPF = listRowComparativo.Range(1, tblCadastros.ListColumns("CPFconjuge").index).Value
            .Telefone = listRowComparativo.Range(1, tblCadastros.ListColumns("Telefoneconjuge").index).Value
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
        With .Acompanhamento
            keyAcompanhamento = listRowComparativo.Range(1, tblCadastros.ListColumns("KeyAcompanhamento").index).Value
            For Each lr In tblAcompanhamento.ListRows
                If lr.Range(1, tblAcompanhamento.ListColumns("KeyAcompanhamento").index).Value = keyAcompanhamento Then
                    .Jan = CBool(GetTableCellValue(lr, "Jan"))
                    .Fev = CBool(GetTableCellValue(lr, "Fev"))
                    .Mar = CBool(GetTableCellValue(lr, "Mar"))
                    .Abr = CBool(GetTableCellValue(lr, "Abr"))
                    .Mai = CBool(GetTableCellValue(lr, "Mai"))
                    .Jun = CBool(GetTableCellValue(lr, "Jun"))
                    .Jul = CBool(GetTableCellValue(lr, "Jul"))
                    .Ago = CBool(GetTableCellValue(lr, "Ago"))
                    .Set = CBool(GetTableCellValue(lr, "Set"))
                    .Out = CBool(GetTableCellValue(lr, "Out"))
                    .Nov = CBool(GetTableCellValue(lr, "Nov"))
                    .Dez = CBool(GetTableCellValue(lr, "Dez"))
                End If
            Next lr
        End With
    End With
End Sub

Public Sub DeletarCadastro()
    InitializeTables

    Dim numParentes As Integer
    Dim lr As ListRow
    
    numParentes = CInt(listRowComparativo.Range(1, tblCadastros.ListColumns("NPessoas").index).Value)

    keyParenteAntesDeletar = listRowComparativo.Range(1, tblCadastros.ListColumns("KeyParente").index).Value
    keyAcompanhamentoAntesDeletar = listRowComparativo.Range(1, tblCadastros.ListColumns("KeyAcompanhamento").index).Value
    
    listRowComparativo.Delete
    
    If numParentes > 0 Then
        Dim index As Integer
        
        With tblParentes
            For index = .ListRows.count To 1 Step -1
                If .ListRows(index).Range(1, .ListColumns("KeyParente").index).Value = keyParenteAntesDeletar Then
                    .ListRows(index).Delete
                End If
            Next index
        End With
    End If
    
    With tblAcompanhamento
        For Each lr In .ListRows
            If lr.Range(1, .ListColumns("KeyAcompanhamento").index).Value = keyAcompanhamentoAntesDeletar Then
                lr.Delete
                Exit For
            End If
        Next lr
    End With
End Sub

Private Sub InitializeTables()
    Dim totalRows As Integer
    If tblCadastros Is Nothing Then
        Set tblCadastros = wksCADASTROS.ListObjects(1)
    End If
    If tblParentes Is Nothing Then
        Set tblParentes = wksPARENTES.ListObjects(1)
    End If
    If tblAcompanhamento Is Nothing Then
        Set tblAcompanhamento = wksACOMPANHAMENTO.ListObjects(1)
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
    On Error Resume Next
    With tblAcompanhamento
        totalRows = .DataBodyRange.Rows.count
        If Err.Number <> 0 Then .ListRows.Add
    End With
    On Error GoTo 0
End Sub

Private Function AddRow(ByRef tbl As ListObject) As ListRow
    With tbl.ListRows
        If Application.WorksheetFunction.CountBlank(.item(.count).Range.Cells) <> .item(.count).Range.Cells.count Then
            .Add
        End If
            
        Set AddRow = .item(.count)
    End With
End Function

'Private Function NormalizarDados(ByVal strDado As String, ByVal padrao As String, ByVal locate As String)
'    If strDado = locate Then
'        NormalizarDados = padrao
'    Else
'        NormalizarDados = strDado
'    End If
'End Function

Public Function GetCadastroListRow(ByVal strNomeAssistido As String) As ListRow
    InitializeTables
    
    Dim lr As ListRow
    
    For Each lr In tblCadastros.ListRows
        If GetTableCellValue(lr, "NomeAssistido") = strNomeAssistido Then
            Set GetCadastroListRow = lr
            Exit Function
        End If
    Next lr
    
    Set GetCadastroListRow = Nothing
End Function

Public Function GetAcompanhamentoListRow(ByVal keyAcompanhamento As Integer) As ListRow
    InitializeTables
    
    Dim lr As ListRow
    
    For Each lr In tblAcompanhamento.ListRows
        If CInt(GetTableCellValue(lr, "KeyAcompanhamento")) = keyAcompanhamento Then
            Set GetAcompanhamentoListRow = lr
            Exit Function
        End If
    Next lr
End Function

Public Function GetTableCellValue(ByRef lr As ListRow, ByVal colName As String) As String
    With lr
        GetTableCellValue = lr.Range(1, lr.Parent.ListColumns(colName).index).Value
    End With
End Function

Public Sub SetTableCellValue(ByRef lr As ListRow, ByVal colName As String, val As Variant)
    lr.Range(1, lr.Parent.ListColumns(colName).index).Value = val
End Sub

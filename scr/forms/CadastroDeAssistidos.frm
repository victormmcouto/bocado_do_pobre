VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroDeAssistidos 
   Caption         =   "FICHA CADASTRAL DOS ASSISTIDOS"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12585
   OleObjectBlob   =   "CadastroDeAssistidos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CadastroDeAssistidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private arrCamposObrigatorios() As MSForms.control

Private Sub cbttAddParentes_Click()
    DescritivoParentes.Show
End Sub

Private Sub cbttCadastrar_Click()
    On Error GoTo ErrHandler
    If CamposObrigatoriosPreenchidos(arrCamposObrigatorios) Then
        Dim result As VbMsgBoxResult
        
        If boolCadastrar Then
            result = MsgBox("Deseja seguir com a realização do cadastro?", _
                            vbYesNo, _
                            "Deseja prosseguir?")
            
            If result = vbNo Then Exit Sub
            
            Call RealizarCadastro
            
            result = MsgBox("Cadastro Realizado! Deseja realizar outro cadastro?", _
                            vbInformation + vbYesNo + vbMsgBoxSetForeground, _
                            "Cadastro Realizado!")
            
            If result = vbYes Then
                Call LimparEntradas(frmDadosCadastrais)
                cadastro = CadastroVazio
            Else
                Unload Me
            End If
        Else
            Call DeletarCadastro
            Call RealizarCadastro
            Unload Me
        End If
    End If
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
End Sub

' ============================================================================================================
' ASSISTIDO
' ============================================================================================================

Private Sub txtbNomeAssistido_Change()
    cadastro.Assistido.nome = txtbNomeAssistido.Value
End Sub

Private Sub combProfissaoAssistido_Change()
    cadastro.Assistido.profissao = combProfissaoAssistido.Value
End Sub

Private Sub combEscolaridadeAssistido_Change()
    cadastro.Assistido.escolaridade = combEscolaridadeAssistido.Value
End Sub

Private Sub txtbCPFAssistido_AfterUpdate()
    With txtbCPFAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            cadastro.Assistido.cpf = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbTelefoneAssistido_AfterUpdate()
    With txtbTelefoneAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarFormatacaoNumTel(.Value) Then
            cadastro.Assistido.telefone = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbDataNascimentoAssistido_AfterUpdate()
    With txtbDataNascimentoAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarMaiorDeIdade(.Value) Then
            cadastro.Assistido.dataNascimento = Format(.Value, "dd/mm/yyyy")
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub combEstadoCivilAssistido_Change()
    With combEstadoCivilAssistido
        If Not (.Value Like "*Casado*" Or .Value = "") Then
            Call EnableFrameControls(FrameConjuge, False)
            combEstadoCivilConjuge.Value = ""
        Else
            Call EnableFrameControls(FrameConjuge, True)
            combEstadoCivilConjuge.Value = combEstadoCivilAssistido.Value
            Call EnableFrameControls(frmEstadoCivilConjuge, False)
        End If
    End With
    
    cadastro.Assistido.estadoCivil = combEstadoCivilAssistido.Value
End Sub

' ============================================================================================================
' CÔNJUGUE
' ============================================================================================================

Private Sub txtbNomeconjuge_Change()
    cadastro.Conjuge.nome = txtbNomeConjuge.Value
End Sub

Private Sub combProfissaoconjuge_Change()
    cadastro.Conjuge.profissao = combProfissaoConjuge.Value
End Sub

Private Sub combEstadoCivilconjuge_Change()
    cadastro.Conjuge.estadoCivil = combEstadoCivilConjuge.Value
End Sub

Private Sub combEscolaridadeconjuge_Change()
    cadastro.Conjuge.escolaridade = combEscolaridadeConjuge.Value
End Sub

Private Sub txtbCPFconjuge_AfterUpdate()
    With txtbCPFConjuge
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            cadastro.Conjuge.cpf = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbTelefoneconjuge_AfterUpdate()
    With txtbTelefoneConjuge
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarFormatacaoNumTel(.Value) Then
            cadastro.Conjuge.telefone = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbDataDeNascimentoconjuge_AfterUpdate()
    With txtbDataDeNascimentoConjuge
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarMaiorDeIdade(.Value) Then
            cadastro.Conjuge.dataNascimento = Format(.Value, "dd/mm/yyyy")
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

' ============================================================================================================
' DEMAIS INFORMAÇÕES
' ============================================================================================================

Private Sub optParticipaProgramaGovSIM_Click()
    cadastro.DemaisInfo.ParticipaProgramaGov = optParticipaProgramaGovSIM.Value
    If optParticipaProgramaGovNAO Then
        Call EnableFrameControls(frmProgGov, False)
    Else
        Call EnableFrameControls(frmProgGov, True)
    End If
End Sub

Private Sub optParticipaProgramaGovNAO_Click()
    optParticipaProgramaGovSIM_Click
End Sub

Private Sub combProgramaGov_Change()
    cadastro.DemaisInfo.ProgramaGov = combProgramaGov.Value
End Sub

Private Sub combTipoMoradia_Change()
    cadastro.DemaisInfo.tipoMoradia = combTipoMoradia.Value
End Sub

Private Sub txtbNPessoasNaCasa_Change()
    cadastro.DemaisInfo.NPessoasNaCasa = txtbNPessoasNaCasa.Value
End Sub

Private Sub optRecebeCestaSIM_Click()
    cadastro.DemaisInfo.recebeCesta = optRecebeCestaSIM.Value
End Sub

Private Sub optRecebeCestaNAO_Click()
    optRecebeCestaSIM_Click
End Sub

Private Sub txtbDataSindicancia_Change()
    cadastro.DemaisInfo.dataSindicancia = txtbDataSindicancia.Value
End Sub

Private Sub txtbNomeVisitador_Change()
    cadastro.DemaisInfo.nomeVisitador = txtbNomeVisitador.Value
End Sub

Private Sub SpinButtonNPessoas_Change()
    Dim totalPessoas As Integer
    
    totalPessoas = SpinButtonNPessoas.Value
     
    lblNPessoasNaCasa.Caption = totalPessoas
    
    If totalPessoas > 0 Then
        If Not ParentesInicializado() Then 'Inicializa o array de parentes caso não tenha sido inicializado
            ReDim cadastro.parentes(1 To totalPessoas)
        ElseIf UBound(cadastro.parentes) < totalPessoas Then 'Redimenciona o array de parentes caso o total mude
            ReDim Preserve cadastro.parentes(1 To totalPessoas)
        End If
        
        cbttAddParentes.Enabled = True
    Else
        cbttAddParentes.Enabled = False
        Erase cadastro.parentes
    End If
    
    cadastro.DemaisInfo.NPessoasNaCasa = totalPessoas
End Sub

' ============================================================================================================
' ENDEREÇO
' ============================================================================================================

Private Sub txtbLogradouro_Change()
    cadastro.Endereco.logradouro = txtbLogradouro.Value
End Sub

Private Sub txtbNumeroLogradouro_Change()
    cadastro.Endereco.NumeroCasa = txtbNumeroLogradouro.Value
End Sub

Private Sub txtbBairro_Change()
    cadastro.Endereco.bairro = txtbBairro.Value
End Sub

Private Sub txtbCidade_Change()
    cadastro.Endereco.cidade = txtbCidade.Value
End Sub

Private Sub UserForm_Initialize()
    Call PopulateComboBoxes                                 'Popula as combo box com os valores armazenados nas tabelas de dados
    txtbDataSindicancia.Enabled = False
    CamposObrigatorios
    
    If boolCadastrar Then
        With txtbDataSindicancia
            .Value = Format(Date, "dd/mm/yyyy")
            cadastro.DemaisInfo.dataSindicancia = .Value
        End With
        lblNPessoasNaCasa.Caption = 0
        cbttCadastrar.Caption = "CADASTRAR"
        EnableFrameControls frmProgGov, False
    Else
        cbttCadastrar.Caption = "ATUALIZAR"
        Call PreencherCampos(Me)
    End If
    
    If Not cadastro.Assistido.estadoCivil Like "*Casado*" Then
        EnableFrameControls FrameConjuge, False
    End If
    If cadastro.DemaisInfo.NPessoasNaCasa = 0 Then
        cbttAddParentes.Enabled = False
    Else
        cbttAddParentes.Enabled = True
    End If
End Sub

Private Sub CamposObrigatorios()
    ReDim arrCamposObrigatorios(1 To 5)
    
    Set arrCamposObrigatorios(1) = txtbNomeAssistido
    Set arrCamposObrigatorios(2) = txtbBairro
    Set arrCamposObrigatorios(3) = txtbLogradouro
    Set arrCamposObrigatorios(4) = txtbNumeroLogradouro
    Set arrCamposObrigatorios(5) = txtbTelefoneAssistido
    
    Dim index As Integer
    
    For index = 1 To UBound(arrCamposObrigatorios)
        Call EidatarComoCampoPrioritario(arrCamposObrigatorios(index))
    Next index
End Sub

Private Sub OptBttProgramaGovFedNAO_Click()
    Image1.Picture = LoadPicture(Application.GetOpenFilename(, , "Selecione a declaração"))
End Sub

Public Sub PopulateComboBoxes()
    Dim tblProfissoes As ListObject
    Dim tblEstadosCivis As ListObject
    Dim tblEscolaridades As ListObject
    Dim tblProgramaGov As ListObject
    Dim tblTipoMoradia As ListObject
    
    With ThisWorkbook
        Set tblProfissoes = wksPROFISSOES.ListObjects(1)
        Set tblEstadosCivis = wksESTADOS_CIVIS.ListObjects(1)
        Set tblEscolaridades = wksESCOLARIDADES.ListObjects(1)
        Set tblProgramaGov = wksPROGRAMAS_GOV.ListObjects(1)
        Set tblTipoMoradia = wksTIPO_MORADIA.ListObjects(1)
    End With
    
    Populate tblProfissoes.DataBodyRange, combProfissaoAssistido
    Populate tblProfissoes.DataBodyRange, combProfissaoConjuge
    
    Populate tblEstadosCivis.DataBodyRange, combEstadoCivilAssistido
    Populate tblEstadosCivis.DataBodyRange, combEstadoCivilConjuge
    
    Populate tblEscolaridades.DataBodyRange, combEscolaridadeAssistido
    Populate tblEscolaridades.DataBodyRange, combEscolaridadeConjuge
    
    Populate tblProgramaGov.DataBodyRange, combProgramaGov
    
    Populate tblTipoMoradia.DataBodyRange, combTipoMoradia
End Sub

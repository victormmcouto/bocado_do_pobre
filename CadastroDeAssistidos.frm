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
        
        result = MsgBox("Deseja seguir com a realizaçã do cadastro?", _
                        vbYesNo, _
                        "Deseja prosseguir?")
        
        If result = vbNo Then Exit Sub
        
        Call RealizarCadastro
        
        result = MsgBox("Cadastro Realizado! Deseja realizar outro cadastro?", _
                        vbInformation + vbYesNo + vbMsgBoxSetForeground, _
                        "Cadastro Realizado!")
        
        If result = vbYes Then
            Call LimparEntradas(frmDadosCadastrais)
            Cadastro = CadastroVazio
        Else
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
    Cadastro.Assistido.Nome = txtbNomeAssistido.Value
End Sub

Private Sub combProfissaoAssistido_Change()
    Cadastro.Assistido.Profissao = combProfissaoAssistido.Value
End Sub

Private Sub combEscolaridadeAssistido_Change()
    Cadastro.Assistido.Escolaridade = combEscolaridadeAssistido.Value
End Sub

Private Sub txtbCPFAssistido_AfterUpdate()
    With txtbCPFAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            Cadastro.Assistido.CPF = .Value
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
            Cadastro.Assistido.Telefone = .Value
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
            Cadastro.Assistido.DataNascimento = Format(.Value, "dd/mm/yyyy")
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
            Call EnableFrameControls(FrameConjugue, False)
            combEstadoCivilConjugue.Value = ""
        Else
            Call EnableFrameControls(FrameConjugue, True)
            combEstadoCivilConjugue.Value = combEstadoCivilAssistido.Value
            Call EnableFrameControls(frmEstadoCivilConjugue, False)
        End If
    End With
    
    Cadastro.Assistido.EstadoCivil = combEstadoCivilAssistido.Value
End Sub

' ============================================================================================================
' CÔNJUGUE
' ============================================================================================================

Private Sub txtbNomeConjugue_Change()
    Cadastro.Conjugue.Nome = txtbNomeConjugue.Value
End Sub

Private Sub combProfissaoConjugue_Change()
    Cadastro.Conjugue.Profissao = combProfissaoConjugue.Value
End Sub

Private Sub combEstadoCivilConjugue_Change()
    Cadastro.Conjugue.EstadoCivil = combEstadoCivilConjugue.Value
End Sub

Private Sub combEscolaridadeConjugue_Change()
    Cadastro.Conjugue.Escolaridade = combEscolaridadeConjugue.Value
End Sub

Private Sub txtbCPFConjugue_AfterUpdate()
    With txtbCPFConjugue
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            Cadastro.Conjugue.CPF = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbTelefoneConjugue_AfterUpdate()
    With txtbTelefoneConjugue
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarFormatacaoNumTel(.Value) Then
            Cadastro.Conjugue.Telefone = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbDataDeNascimentoConjugue_AfterUpdate()
    With txtbDataDeNascimentoConjugue
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarMaiorDeIdade(.Value) Then
            Cadastro.Conjugue.DataNascimento = Format(.Value, "dd/mm/yyyy")
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
    Cadastro.DemaisInfo.ParticipaProgramaGov = optParticipaProgramaGovSIM.Value
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
    Cadastro.DemaisInfo.ProgramaGov = combProgramaGov.Value
End Sub

Private Sub combTipoMoradia_Change()
    Cadastro.DemaisInfo.TipoMoradia = combTipoMoradia.Value
End Sub

Private Sub txtbNPessoasNaCasa_Change()
    Cadastro.DemaisInfo.NPessoasNaCasa = txtbNPessoasNaCasa.Value
End Sub

Private Sub optRecebeCestaSIM_Click()
    Cadastro.DemaisInfo.RecebeCesta = optRecebeCestaSIM.Value
End Sub

Private Sub optRecebeCestaNAO_Click()
    optRecebeCestaSIM_Click
End Sub

Private Sub txtbDataSindicancia_Change()
    Cadastro.DemaisInfo.DataSindicancia = txtbDataSindicancia.Value
End Sub

Private Sub txtbNomeVisitador_Change()
    Cadastro.DemaisInfo.NomeVisitador = txtbNomeVisitador.Value
End Sub

Private Sub SpinButtonNPessoas_Change()
    Dim totalPessoas As Integer
    
    totalPessoas = SpinButtonNPessoas.Value
     
    lblNPessoasNaCasa.Caption = totalPessoas
    
    If totalPessoas > 0 Then
        If Not ParentesInicializado() Then 'Inicializa o array de parentes caso não tenha sido inicializado
            ReDim Cadastro.parentes(1 To totalPessoas)
        ElseIf UBound(Cadastro.parentes) < totalPessoas Then 'Redimenciona o array de parentes caso o total mude
            ReDim Preserve Cadastro.parentes(1 To totalPessoas)
        End If
        
        cbttAddParentes.Enabled = True
    Else
        cbttAddParentes.Enabled = False
        Erase Cadastro.parentes
    End If
    
    Cadastro.DemaisInfo.NPessoasNaCasa = totalPessoas
End Sub

' ============================================================================================================
' ENDEREÇO
' ============================================================================================================

Private Sub txtbLogradouro_Change()
    Cadastro.Endereco.Logradouro = txtbLogradouro.Value
End Sub

Private Sub txtbNumeroLogradouro_Change()
    Cadastro.Endereco.NumeroCasa = txtbNumeroLogradouro.Value
End Sub

Private Sub txtbBairro_Change()
    Cadastro.Endereco.Bairro = txtbBairro.Value
End Sub

Private Sub txtbCidade_Change()
    Cadastro.Endereco.Cidade = txtbCidade.Value
End Sub

Private Sub UserForm_Initialize()
    Call PopulateComboBoxes                                 'Popula as combo box com os valores armazenados nas tabelas de dados
    txtbDataSindicancia.Enabled = False
    CamposObrigatorios
    
    If boolCadastrar Then
        With txtbDataSindicancia
            .Value = Format(Date, "dd/mm/yyyy")
            Cadastro.DemaisInfo.DataSindicancia = .Value
        End With
        lblNPessoasNaCasa.Caption = 0
    Else
        Call PreencherCampos(Me)
    End If
    
    If Not Cadastro.Assistido.EstadoCivil Like "*Casado*" Then
        EnableFrameControls FrameConjugue, False
    End If
    If Cadastro.DemaisInfo.NPessoasNaCasa = 0 Then
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
    Populate tblProfissoes.DataBodyRange, combProfissaoConjugue
    
    Populate tblEstadosCivis.DataBodyRange, combEstadoCivilAssistido
    Populate tblEstadosCivis.DataBodyRange, combEstadoCivilConjugue
    
    Populate tblEscolaridades.DataBodyRange, combEscolaridadeAssistido
    Populate tblEscolaridades.DataBodyRange, combEscolaridadeConjugue
    
    Populate tblProgramaGov.DataBodyRange, combProgramaGov
    
    Populate tblTipoMoradia.DataBodyRange, combTipoMoradia
End Sub

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
            
            repoCadastro.Repository_save cdst
            
            result = MsgBox("Cadastro Realizado! Deseja realizar outro cadastro?", _
                            vbInformation + vbYesNo + vbMsgBoxSetForeground, _
                            "Cadastro Realizado!")
            
            If result = vbYes Then
                Call LimparEntradas(frmDadosCadastrais)
                Set cdst = New Cadastro
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
    cdst.getAssistido.setNome = txtbNomeAssistido.Value
End Sub

Private Sub combProfissaoAssistido_Change()
    cdst.getAssistido.setProfissao = combProfissaoAssistido.Value
End Sub

Private Sub combEscolaridadeAssistido_Change()
    cdst.getAssistido.setEscolaridade = enumEscolaridade.Enums_getNum(combEscolaridadeAssistido.Value)
End Sub

Private Sub txtbCPFAssistido_AfterUpdate()
    With txtbCPFAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            cdst.getAssistido.setCpf = .Value
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
            cdst.getAssistido.setTelefone = .Value
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
            cdst.getAssistido.setDataNascimento = Format(.Value, "dd/mm/yyyy")
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub combEstadoCivilAssistido_Change()
    With combEstadoCivilAssistido
        If Not (.Value Like "*Casado*") Then
            Call EnableFrameControls(FrameConjuge, False)
            combEstadoCivilConjuge.Value = ""
        Else
            Call EnableFrameControls(FrameConjuge, True)
            combEstadoCivilConjuge.Value = combEstadoCivilAssistido.Value
            Call EnableFrameControls(frmEstadoCivilConjuge, False)
        End If
    End With
    
    cdst.getAssistido.setEstadoCivil = enumEstadoCivil.Enums_getNum(combEstadoCivilAssistido.Value)
End Sub

' ============================================================================================================
' CÔNJUGUE
' ============================================================================================================

Private Sub txtbNomeconjuge_Change()
    cdst.getAssistido.getConjuge.setNome = txtbNomeConjuge.Value
End Sub

Private Sub combProfissaoconjuge_Change()
    cdst.getAssistido.getConjuge.setProfissao = combProfissaoConjuge.Value
End Sub

Private Sub combEstadoCivilconjuge_Change()
    cdst.getAssistido.getConjuge.setEstadoCivil = enumEstadoCivil.Enums_getNum(combEstadoCivilConjuge.Value)
End Sub

Private Sub combEscolaridadeconjuge_Change()
    cdst.getAssistido.getConjuge.setEscolaridade = enumEscolaridade.Enums_getNum(combEscolaridadeConjuge.Value)
End Sub

Private Sub txtbCPFconjuge_AfterUpdate()
    With txtbCPFConjuge
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            cdst.getAssistido.getConjuge.setCpf = .Value
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
            cdst.getAssistido.getConjuge.setTelefone = .Value
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
            cdst.getAssistido.getConjuge.setDataNascimento = Format(.Value, "dd/mm/yyyy")
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
    cdst.setParticipaProgGov = optParticipaProgramaGovSIM.Value
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
    cdst.setProgGov = combProgramaGov.Value
End Sub

Private Sub combTipoMoradia_Change()
    cdst.setTipoMoradia = enumTipoMoradia.Enums_getNum(combTipoMoradia.Value)
End Sub

'Private Sub txtbNPessoasNaCasa_Change()
'    cdst.setNDependentes = CInt(txtbNPessoasNaCasa.Value)
'End Sub

Private Sub optRecebeCestaSIM_Click()
    cdst.setRecebeCesta = optRecebeCestaSIM.Value
End Sub

Private Sub optRecebeCestaNAO_Click()
    optRecebeCestaSIM_Click
End Sub

Private Sub txtbDataSindicancia_Change()
    cdst.setDataSindicancia = CDate(txtbDataSindicancia.Value)
End Sub

Private Sub txtbNomeVisitador_Change()
    cdst.setNomeVisitador = txtbNomeVisitador.Value
End Sub

Private Sub SpinButtonNPessoas_Change()
    Dim totalPessoas As Integer
    
    totalPessoas = SpinButtonNPessoas.Value
     
    lblNPessoasNaCasa.Caption = totalPessoas
    
    If totalPessoas = 0 Then
        cbttAddParentes.Enabled = False
        cdst.getAssistido.RemoveDependentes cdst.getAssistido.getTotalDependentes
    Else
        cbttAddParentes.Enabled = True
        
        If cdst.getAssistido.getTotalDependentes < totalPessoas Then
            cdst.getAssistido.AddDependentes 1
        ElseIf cdst.getAssistido.getTotalDependentes > totalPessoas Then
            cdst.getAssistido.RemoveDependentes 1
        End If
    End If
    
    'cdst.setNDependentes = totalPessoas
End Sub

' ============================================================================================================
' ENDEREÇO
' ============================================================================================================

Private Sub txtbLogradouro_Change()
    cdst.setLogradouro = txtbLogradouro.Value
End Sub

Private Sub txtbNumeroLogradouro_Change()
    cdst.setNumero = txtbNumeroLogradouro.Value
End Sub

Private Sub txtbBairro_Change()
    cdst.setBairro = txtbBairro.Value
End Sub

Private Sub txtbCidade_Change()
    cdst.setCidade = txtbCidade.Value
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
    Dim profissoes As Variant
    
    profissoes = appendArrays(ValoresUnicos(repoAssistido.Repository_getTbl.getListColumn("profissao").DataBodyRange), _
                              ValoresUnicos(repoConjuge.Repository_getTbl.getListColumn("profissao").DataBodyRange))

    
    combProfissaoAssistido.List = profissoes
    combProfissaoConjuge.List = profissoes
    
    combEstadoCivilAssistido.List = enumEstadoCivil.Enums_getArray
    combEstadoCivilConjuge.List = enumEstadoCivil.Enums_getArray
    
    combEscolaridadeAssistido.List = enumEscolaridade.Enums_getArray
    combEscolaridadeConjuge.List = enumEscolaridade.Enums_getArray
    
    combTipoMoradia.List = enumTipoMoradia.Enums_getArray
    
    combProgramaGov.List = ValoresUnicos(repoCadastro.Repository_getTbl.getListColumn("ProgGov").DataBodyRange)
End Sub

Private Sub UserForm_Initialize()
    Call PopulateComboBoxes                                 'Popula as combo box com os valores armazenados nas tabelas de dados
    txtbDataSindicancia.Enabled = False
    CamposObrigatorios
    
    If boolCadastrar Then
        With txtbDataSindicancia
            .Value = Format(Date, "dd/mm/yyyy")
            'cdst.setDataSindicancia = .Value
        End With
        lblNPessoasNaCasa.Caption = 0
        cbttCadastrar.Caption = "CADASTRAR"
        EnableFrameControls frmProgGov, False
    Else
        cbttCadastrar.Caption = "ATUALIZAR"
        Call PreencherCampos(Me)
    End If
    
    If Not cdst.getAssistido Is Nothing Then
        If cdst.getAssistido.getEstadoCivil <> enumEstadoCivil.CASADO_A Then
            EnableFrameControls FrameConjuge, False
        End If
    End If
    If cdst.getAssistido.getTotalDependentes = 0 Then
        cbttAddParentes.Enabled = False
    Else
        cbttAddParentes.Enabled = True
    End If
End Sub

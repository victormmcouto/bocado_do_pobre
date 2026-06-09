VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroDeAssistidos 
   Caption         =   "FICHA CADASTRAL DOS ASSISTIDOS"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
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
    Call AdicionarParentes(CInt(txtbNPessoasNaCasa.Value))
End Sub

Private Sub cbttCadastrar_Click()
    On Error GoTo ErrHandler
    If CamposObrigatoriosPreenchidos(arrCamposObrigatorios) Then
        Dim result As VbMsgBoxResult
        
        'Call RealizarCadastro
        
        result = MsgBox("Cadastro Realizado! Deseja realizar outro cadastro?", _
                        vbInformation + vbYesNo + vbMsgBoxSetForeground, _
                        "Cadastro Realizado!")
        
        If result = vbYes Then
            Call LimparEntradas(Me)
        Else
            Unload Me
        End If
    End If
    
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
End Sub

Private Sub combEstadoCivilAssistido_Change()
    With combEstadoCivilAssistido
        If Not .Value Like "*Casado*" Or .Value = "" Then
            Call EnableFrameControls(FrameConjugue, False)
        Else
            Call EnableFrameControls(FrameConjugue, True)
        End If
    End With
End Sub

Private Sub SpinButtonNPessoas_Change()
    txtbNPessoasNaCasa.Value = SpinButtonNPessoas.Value
End Sub

Private Sub txtbNomeAssistido_Change()
    cadastro.Assistido.Nome = txtbNomeAssistido.Value
End Sub

Private Sub txtbDataNascimentoAssistido_AfterUpdate()
    With txtbDataNascimentoAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarMaiorDeIdade(.Value) Then
            cadastro.Assistido.DataNascimento = Format(.Value, "dd/mm/yyyy")
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
            cadastro.Conjugue.DataNascimento = Format(.Value, "dd/mm/yyyy")
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
            cadastro.Assistido.Telefone = .Value
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
            cadastro.Assistido.Conjugue.Telefone = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbCPFAssistido_AfterUpdate()
    With txtbCPFAssistido
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            cadastro.Assistido.CPF = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbCPFConjugue_AfterUpdate()
    With txtbCPFConjugue
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarCPF(.Value) Then
            cadastro.Assistido.Conjugue.CPF = .Value
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub UserForm_Initialize()
    Call PopulateComboBoxes 'Popula as combo box com os valores armazenados nas tabelas de dados
    With txtbDataSindicancia
        .Value = Format(Date, "dd/mm/yyyy")
        .Enabled = False
    End With
    txtbNPessoasNaCasa.Value = 1
    EnableFrameControls FrameConjugue, False
    
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
    
    Populate tblProfissoes, combProfissaoAssistido
    Populate tblProfissoes, combProfissaoConjugue
    
    Populate tblEstadosCivis, combEstadoCivilAssistido
    Populate tblEstadosCivis, combEstadoCivilConjugue
    
    Populate tblEscolaridades, combEscolaridadeAssistido
    Populate tblEscolaridades, combEscolaridadeConjugue
    
    Populate tblProgramaGov, combProgramaGov
    
    Populate tblTipoMoradia, combTipoMoradia
End Sub

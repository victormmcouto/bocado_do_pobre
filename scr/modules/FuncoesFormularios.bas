Attribute VB_Name = "FuncoesFormularios"
Option Explicit

'Public cadastro As cadastro

'Public Type Conjuge
'    nome As String
'    dataNascimento As Date
'    profissao As String
'    estadoCivil As String
'    escolaridade As String
'    telefone As String
'    cpf As String
'End Type
'
'Public Type parente
'    nome As String
'    dataNascimento As Date
'    escolaridade As String
'    GrauParentesco As String
'End Type
'
'Public Type Assistido
'    nome As String
'    dataNascimento As Date
'    estadoCivil As String
'    profissao As String
'    escolaridade As String
'    telefone As String
'    cpf As String
'End Type
'
'Public Type Endereco
'    logradouro As String
'    NumeroCasa As String
'    bairro As String
'    cidade As String
'End Type
'
'Public Type DemaisInfo
'    ParticipaProgramaGov As Boolean
'    ProgramaGov As String
'    tipoMoradia As String
'    NPessoasNaCasa As Integer
'    recebeCesta As Boolean
'    dataSindicancia As Date
'    nomeVisitador As String
'End Type
'
'Type AcompanhamentoMes
'    Jan As Boolean
'    Fev As Boolean
'    Mar As Boolean
'    Abr As Boolean
'    Mai As Boolean
'    Jun As Boolean
'    Jul As Boolean
'    Ago As Boolean
'    Set As Boolean
'    Out As Boolean
'    Nov As Boolean
'    Dez As Boolean
'End Type
'
'Public Type cadastro
'    Assistido As Assistido
'    Conjuge As Conjuge
'    parentes() As parente
'    Endereco As Endereco
'    DemaisInfo As DemaisInfo
'    acompanhamento As AcompanhamentoMes
'End Type
'
'Public Enum Errors1
'    errFormatoDataInvalida = vbObjectError + 1
'    errFormatoNumeroTelefoneInvalido = vbObjectError + 2
'    errValidacaoCPF = vbObjectError + 3
'    errCamposObrigatorios = vbObjectError + 4
'End Enum
'
'Public total As Integer

Public Function ParentesInicializado() As Boolean
    On Error Resume Next

    Dim lb As Long
    lb = LBound(Cadastro.parentes)

    ParentesInicializado = (Err.Number = 0)

    On Error GoTo 0
End Function

Public Sub Populate(ByVal data As Variant, ByRef comb As ComboBox)
    Dim index As Integer
    
    For index = 0 To UBound(data, 1)
        comb.AddItem data(index)
        ReDim Preserve arrControl(1 To index + 1)
        arrControl(index) = data(index)
        For indexControl = 0 To UBound(arrControl)
    Next index
End Sub

Public Function ValidarDataCompleta(ByVal strDate As String) As Boolean
    Dim regexDate As New RegExp
    
    regexDate.Pattern = "^\d{2}/\d{2}/\d{4}$"
    
    If Not (IsDate(Trim(strDate)) And regexDate.test(Trim(strDate))) Then
        Err.Raise errFormatoDataInvalida, "Foratação Errada!", "Formato de data errado! Formato válido: (dd/mm/aaaa)"
    End If
    
    ValidarDataCompleta = True
End Function

Public Function ValidarMaiorDeIdade(ByVal strBirthDate As String) As Boolean
    If ValidarDataCompleta(strBirthDate) Then
        If DateDiff("m", CDate(strBirthDate), Date) / 12 <= 18 Then
            Err.Raise errFormatoDataInvalida, "Idade Mínima não Atingida!", "O referendo deve ter idade igual ou supeiror a 18 anos!"
        End If
    End If
    
    ValidarMaiorDeIdade = True
End Function

Public Function ValidarFormatacaoNumTel(ByVal strNumTel As String) As Boolean
    Dim regexNumTel As New RegExp
    
    regexNumTel.Pattern = "^\(?\d{2}\)?\s?9?\s?\d{4}\s?-?\s?\d{4}$"
    
    
    If Not regexNumTel.test(Trim(strNumTel)) Then
        Err.Raise errFormatoNumeroTelefoneInvalido, _
                  "Formato telefônico errado!", _
                  "Formato de número telefônico errado! Formatos válidos: DDD X XXXX-XXXX ou DDD XXXX-XXXX"
    End If
    
    ValidarFormatacaoNumTel = True
End Function

Public Function ValidarCPF(ByVal strCPF As String) As Boolean
    Dim regexCPF As New RegExp
    Dim index As Integer
    
    regexCPF.Pattern = "^((\d{11})|(\d{3}\.){2}\d{3}-\d{2})$"
    
    If Not regexCPF.test(Trim(strCPF)) Then
        Err.Raise errValidacaoCPF, "Formato de CPF Inválido!", _
                  "A formatação do CPF não está correta! Formatos válidos: XXX.XXX.XXX-XX ou XXXXXXXXXXX"
    End If
    If Not ValidarDigitosCPF(Trim(strCPF)) Then
        Err.Raise errValidacaoCPF, "CPF Inválido!", _
                  "O valor inserido para o CPF não é válido!"
    End If
    
    ValidarCPF = True
End Function

Private Function ValidarDigitosCPF(ByVal strCPF) As Boolean
    Dim i As Long
    Dim Soma As Long
    Dim DV1 As Long
    Dim DV2 As Long
    Dim Resto As Long
    Dim primeiroCaracter As String
    
    strCPF = Trim(Replace(Replace(strCPF, ".", ""), "-", ""))
    
    primeiroCaracter = Mid$(strCPF, 1, 1)
    
    For i = 1 To Len(strCPF)
        If Mid$(strCPF, i, 1) <> primeiroCaracter Then
            Exit For
        Else
            If i = Len(strCPF) Then Exit Function
        End If
    Next i
    
    For i = 1 To 9
        Soma = Soma + val(Mid$(strCPF, i, 1)) * (11 - i)
    Next i

    Resto = Soma Mod 11

    If Resto < 2 Then
        DV1 = 0
    Else
        DV1 = 11 - Resto
    End If

    Soma = 0

    For i = 1 To 10
        Soma = Soma + val(Mid$(strCPF, i, 1)) * (12 - i)
    Next i

    Resto = Soma Mod 11

    If Resto < 2 Then
        DV2 = 0
    Else
        DV2 = 11 - Resto
    End If
    
    ValidarDigitosCPF = (DV1 = val(Mid$(strCPF, 10, 1))) And _
                        (DV2 = val(Mid$(strCPF, 11, 1)))
End Function

Public Sub EnableFrameControls(ByRef frm As MSForms.frame, boolEnable As Boolean)
    Dim ctrl As MSForms.control
    
    frm.Enabled = boolEnable
    
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "Frame" Then
            EnableFrameControls ctrl, boolEnable
            ctrl.Enabled = boolEnable
        Else
            ctrl.Enabled = boolEnable
        End If
    Next ctrl
End Sub

Public Sub EidatarComoCampoPrioritario(ByRef ctrl As MSForms.control)
    With ctrl
        .BorderColor = &HFF&
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
    End With
End Sub

Public Function CamposObrigatoriosPreenchidos(ByRef arrCampos() As MSForms.control)
    Dim index As Integer
    
    For index = 1 To UBound(arrCampos)
        If arrCampos(index).Value = "" Then
            Err.Raise errCamposObrigatorios, "Campos obirgatórios faltantes!", _
                      "Há campos obrigatórios sem preenchimento! É impossível realizar um cadastro sem esse mínimo de informações!"
        End If
    Next index
    
    CamposObrigatoriosPreenchidos = True
End Function

Public Sub LimparEntradas(ByRef frm As MSForms.frame)
    Dim ctrl As MSForms.control
    
    For Each ctrl In frm.Controls
        Debug.Print ctrl.Name
        If TypeName(ctrl) = "Frame" Then
            LimparEntradas ctrl
        ElseIf TypeName(ctrl) = "Label" Then
            ctrl.Caption = 0
        ElseIf TypeName(ctrl) = "OptionButton" Or TypeName(ctrl) = "CheckBox" Then
            ctrl.Value = False
        ElseIf ctrl.Name = "txtbDataSindicancia" Then
            ctrl.Value = Format(Date, "dd/mm/yyyy")
        ElseIf TypeName(ctrl) = "SpinButton" Then
            ctrl.Value = 0
        ElseIf Not TypeName(ctrl) = "CommandButton" Then
            ctrl.Value = ""
        End If
    Next ctrl
End Sub

Public Function CadastroVazio() As Cadastro
    Dim newCadastro As Cadastro
    CadastroVazio = newCadastro
End Function

Public Sub PreencherCampos(ByRef form As CadastroDeAssistidos)
    With form
        .txtbNomeAssistido.Value = Cadastro.Assistido.nome
        .txtbDataNascimentoAssistido.Value = Cadastro.Assistido.dataNascimento
        .combEstadoCivilAssistido.Value = Cadastro.Assistido.estadoCivil
        .combProfissaoAssistido.Value = Cadastro.Assistido.profissao
        .combEscolaridadeAssistido.Value = Cadastro.Assistido.escolaridade
        .txtbCPFAssistido.Value = Cadastro.Assistido.cpf
        .txtbTelefoneAssistido.Value = Cadastro.Assistido.telefone
    
        .txtbNomeConjuge.Value = Cadastro.Conjuge.nome
        .txtbDataDeNascimentoConjuge.Value = Cadastro.Conjuge.dataNascimento
        .combEstadoCivilConjuge.Value = Cadastro.Conjuge.estadoCivil
        .combProfissaoConjuge.Value = Cadastro.Conjuge.profissao
        .combEscolaridadeConjuge.Value = Cadastro.Conjuge.escolaridade
        .txtbCPFConjuge.Value = Cadastro.Conjuge.cpf
        .txtbTelefoneConjuge.Value = Cadastro.Conjuge.telefone
        
        .optParticipaProgramaGovNAO.Value = Not Cadastro.DemaisInfo.ParticipaProgramaGov
        .optParticipaProgramaGovSIM.Value = Cadastro.DemaisInfo.ParticipaProgramaGov
        .combProgramaGov.Value = Cadastro.DemaisInfo.ProgramaGov
        .combTipoMoradia.Value = Cadastro.DemaisInfo.tipoMoradia
        .lblNPessoasNaCasa.Caption = Cadastro.DemaisInfo.NPessoasNaCasa
        .SpinButtonNPessoas.Value = Cadastro.DemaisInfo.NPessoasNaCasa
        .optRecebeCestaSIM.Value = Cadastro.DemaisInfo.recebeCesta
        .optRecebeCestaNAO.Value = Not Cadastro.DemaisInfo.recebeCesta
        
        .txtbDataSindicancia.Value = Cadastro.DemaisInfo.dataSindicancia
        .txtbNomeVisitador.Value = Cadastro.DemaisInfo.nomeVisitador
        
        .txtbLogradouro.Value = Cadastro.Endereco.logradouro
        .txtbNumeroLogradouro.Value = Cadastro.Endereco.NumeroCasa
        .txtbBairro.Value = Cadastro.Endereco.bairro
        .txtbCidade.Value = Cadastro.Endereco.cidade
    End With
End Sub

Public Function Capitalize(str As String) As String
    Capitalize = UCase(Mid(str, 1, 1)) & LCase(Mid(str, 2, Len(str)))
End Function

Function ValoresUnicos(rng As Range) As Variant
    Dim dict As Object
    Dim cel As Range

    Set dict = CreateObject("Scripting.Dictionary")

    For Each cel In rng.Cells
        If Not dict.Exists(cel.Value) And cel.Value <> "" Then
            dict.Add cel.Value, Empty
        End If
    Next cel

    ValoresUnicos = dict.Keys
End Function

Function appendArrays(ParamArray arrs() As Variant) As Variant
    Dim arrTemp() As Variant
    Dim indexArrs As Long
    Dim indexArr As Long
    Dim pos As Long
    
    pos = -1
    
    For indexArrs = LBound(arrs) To UBound(arrs)
        For indexArr = LBound(arrs(indexArrs)) To UBound(arrs(indexArrs))
            pos = pos + 1
            ReDim Preserve arrTemp(0 To pos)
            arrTemp(pos) = arrs(indexArrs)(indexArr)
        Next indexArr
    Next indexArrs
    
    appendArrays = arrTemp
End Function

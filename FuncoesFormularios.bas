Attribute VB_Name = "FuncoesFormularios"
Option Explicit

Public cadastro As cadastro

Public Type Conjugue
    Nome As String
    DataNascimento As Date
    Profissao As String
    EstadoCivil As String
    Escolaridade As String
    Telefone As String
    CPF As String
End Type

Public Type parente
    Nome As String
    DataNascimento As Date
    Escolaridade As String
    GrauParentesco As String
End Type

Public Type Assistido
    Nome As String
    DataNascimento As Date
    EstadoCivil As String
    Profissao As String
    Escolaridade As String
    Telefone As String
    CPF As String
    Conjugue As Conjugue
End Type

Public Type Endereco
    Logradouro As String
    NumeroCasa As String
    Bairro As String
    Cidade As String
    PathAutorizacao As String
End Type

Public Type DemaisInfo
    ParticipaProgramaGov As Boolean
    ProgramaGov As String
    TipoMoradia As String
    NPessoasNaCasa As Integer
    RecebeCesta As Boolean
    DataSindicancia As Date
    NomeVisitador As String
End Type

Public Type cadastro
    Assistido As Assistido
    Conjugue As Conjugue
    parentes() As parente
    Endereco As Endereco
    DemaisInfo As DemaisInfo
End Type

Public Enum errors
    errFormatoDataInvalida = vbObjectError + 1
    errFormatoNumeroTelefoneInvalido = vbObjectError + 2
    errValidacaoCPF = vbObjectError + 3
End Enum

Public total As Integer

Public Sub call_(total_ As Integer)
    total = total_
    DescritivoParentes.Show
End Sub

Public Sub Main()
    Dim formCadastro As New CadastroDeAssistidos
    
    formCadastro.Show
    
End Sub

Public Function ParentesInicializado(parentes() As parente) As Boolean
    On Error Resume Next

    Dim lb As Long
    lb = LBound(parentes)

    ParentesInicializado = (Err.Number = 0)

    On Error GoTo 0
End Function

Public Sub Populate(ByRef tbl As ListObject, ByRef comb As ComboBox)
    Dim item As Range
    
    For Each item In tbl.DataBodyRange.Cells
        comb.AddItem item.Value
    Next item
End Sub

Public Function ValidarDataCompleta(ByVal strDate As String) As Boolean
    Dim regexDate As New RegExp
    
    regexDate.Pattern = "^\d{2}/\d{2}/\d{4}$"
    
    If Not (IsDate(Trim(strDate)) And regexDate.Test(Trim(strDate))) Then
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
    
    If Not regexNumTel.Test(Trim(strNumTel)) Then
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
    
    If Not regexCPF.Test(Trim(strCPF)) Then
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
        Soma = Soma + Val(Mid$(strCPF, i, 1)) * (11 - i)
    Next i

    Resto = Soma Mod 11

    If Resto < 2 Then
        DV1 = 0
    Else
        DV1 = 11 - Resto
    End If

    Soma = 0

    For i = 1 To 10
        Soma = Soma + Val(Mid$(strCPF, i, 1)) * (12 - i)
    Next i

    Resto = Soma Mod 11

    If Resto < 2 Then
        DV2 = 0
    Else
        DV2 = 11 - Resto
    End If
    
    ValidarDigitosCPF = (DV1 = Val(Mid$(strCPF, 10, 1))) And _
                        (DV2 = Val(Mid$(strCPF, 11, 1)))
End Function

Sub teste()
    ValidarCPF "118.782.536-03"
    
End Sub

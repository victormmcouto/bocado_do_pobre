Attribute VB_Name = "FuncoesFormularios"
Option Explicit

Public cadastro As cadastro

Public Type Conjugue
    Nome As String
    DataNascimento As Date
    Profissao As String
    EstadoCivil As String
    Escolaridade As String
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

    ParentesInicializado = (Err.Number <> 0)

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
    
    regexDate.Pattern = "[0-9]{2}/[0-9]{2}/[0-9]{4}"
    
    If Not (IsDate(strDate) And regexDate.Test(strDate)) Then
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

Sub teste()
    On Error Resume Next
    Debug.Print ValidarData("13/11")
    If Err.Number <> 0 Then Debug.Print Err.Description
    On Error GoTo 0
    On Error Resume Next
    Debug.Print ValidarData("2000")
    If Err.Number <> 0 Then Debug.Print Err.Description
    On Error GoTo 0
    On Error Resume Next
    Debug.Print ValidarData("32/11/2001")
    If Err.Number <> 0 Then Debug.Print Err.Description
    On Error GoTo 0
    On Error Resume Next
    Debug.Print ValidarData("12/04/2015")
    If Err.Number <> 0 Then Debug.Print Err.Description
    On Error GoTo 0
    On Error Resume Next
    Debug.Print ValidarData("12/04/2007")
    If Err.Number <> 0 Then Debug.Print Err.Description
    On Error GoTo 0
End Sub

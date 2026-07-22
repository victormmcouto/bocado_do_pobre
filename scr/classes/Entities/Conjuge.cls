VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Conjuge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private me_nome As String
Private me_dataNascimento As Date
Private me_escolaridade As Integer
Private me_estadoCivil As Integer
Private me_profissao As String
Private me_cpf As String
Private me_telefone As String
Private me_key As Long

Private Sub Class_Initialize()
    me_nome = vbNullString
    me_dataNascimento = 0
    me_escolaridade = 0
    me_estadoCivil = 0
    me_profissao = vbNullString
    me_cpf = vbNullString
    me_telefone = vbNullString
    me_key = 0
End Sub

Public Function create(ByVal nome As String, _
                       ByVal dataNascimento As Date, _
                       ByVal escolaridade As Integer, _
                       ByVal estadoCivil As Integer, _
                       ByVal profissao As String, _
                       ByVal cpf As String, _
                       ByVal telefone As String) As Conjuge
    
    me_nome = nome
    me_dataNascimento = dataNascimento
    me_escolaridade = escolaridade
    me_estadoCivil = estadoCivil
    me_profissao = profissao
    me_cpf = cpf
    me_telefone = telefone
    
    Set create = Me
End Function

Public Function CreateFromArray(ByVal arr As Variant) As Conjuge
    me_key = arr(1, 1)
    me_nome = arr(1, 2)
    me_dataNascimento = arr(1, 3)
    me_escolaridade = arr(1, 4)
    me_estadoCivil = arr(1, 5)
    me_profissao = arr(1, 6)
    me_cpf = arr(1, 7)
    me_telefone = arr(1, 8)
    
    Set CreateFromArray = Me
End Function

Public Property Get getKey() As Long
    getKey = me_key
End Property

Public Property Let setKey(ByVal key As Long)
    me_key = key
End Property

Public Property Get getNome() As String
    getNome = me_nome
End Property

Public Property Let setNome(ByVal nome As String)
    me_nome = nome
End Property

Public Property Get getDataNascimento() As Date
    getDataNascimento = me_dataNascimento
End Property

Public Property Let setDataNascimento(ByVal data As Date)
    me_dataNascimento = data
End Property

Public Property Get getEscolaridade() As Integer
    getEscolaridade = me_escolaridade
End Property

Public Property Let setEscolaridade(ByVal escolaridade As Integer)
    me_escolaridade = escolaridade
End Property

Public Property Get getEstadoCivil() As Integer
    getEstadoCivil = me_estadoCivil
End Property

Public Property Let setEstadoCivil(ByVal estadoCivil As Integer)
    me_estadoCivil = estadoCivil
End Property

Public Property Get getProfissao() As String
    getProfissao = me_profissao
End Property

Public Property Let setProfissao(ByVal profissao As String)
    me_profissao = profissao
End Property

Public Property Get getCpf() As String
    getCpf = me_cpf
End Property

Public Property Let setCpf(ByVal cpf As String)
    me_cpf = cpf
End Property

Public Property Get getTelefone() As String
    getTelefone = me_telefone
End Property

Public Property Let setTelefone(ByVal telefone As String)
    me_telefone = telefone
End Property

Public Function Equals(ByRef outro As Conjuge) As Boolean
    If outro Is Nothing Then Exit Function
    
    If Me.getKey = 0 Or outro.getKey = 0 Then
        Equals = (Me Is outro)
        Exit Function
    End If
    
    Equals = (Me.getKey = outro.getKey)
End Function

Public Function toArray() As Variant
   toArray = Array(me_key, _
                   me_nome, _
                   me_dataNascimento, _
                   me_escolaridade, _
                   me_estadoCivil, _
                   me_profissao, _
                   me_cpf, _
                   me_telefone)
End Function

Public Function toString(Optional ByVal showExtra As Boolean = True) As String
    Dim escolaridade As New enumEscolaridade
    Dim estadoCivil As New enumEstadoCivil
    
    toString = "Key             = " & me_key & ", " & vbNewLine & _
               "Nome            = " & me_nome & ", " & vbNewLine & _
               "Data Nascimento = " & me_dataNascimento & ", " & vbNewLine & _
               "Escolaridade    = " & escolaridade.Enums_getNome(me_escolaridade) & ", " & vbNewLine & _
               "Estado Civil    = " & estadoCivil.Enums_getNome(me_estadoCivil) & ", " & vbNewLine & _
               "Profissao       = " & me_profissao & ", " & vbNewLine & _
               "CPF             = " & me_cpf & ", " & vbNewLine & _
               "Telefone        = " & me_telefone & vbNewLine
    
    Set escolaridade = Nothing
    Set estadoCivil = Nothing
End Function



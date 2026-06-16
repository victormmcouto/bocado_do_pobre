VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assistido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private me_nome As String
Private me_dataNascimento As Date
Private me_escolaridade As String
Private me_estadoCivil As String
Private me_profissao As String
Private me_cpf As String
Private me_telefone As String
Private me_key As Integer

Private me_dependentes As New Collection

Public Function Create(ByVal nome As String, _
                       ByVal dataNascimento As Date, _
                       ByVal escolaridade As String, _
                       ByVal estadoCivil As String, _
                       ByVal profissao As String, _
                       ByVal cpf As String, _
                       ByVal telefone As String) As assistido
    
    me_nome = nome
    me_dataNascimento = dataNascimento
    me_escolaridade = escolaridade
    me_estadoCivil = estadoCivil
    me_profissao = profissao
    me_cpf = cpf
    me_telefone = telefone
    
    Set Create = Me
End Function

Public Property Get getKey() As Integer
    getKey = me_key
End Property

Public Property Let setKey(ByVal key As Integer)
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

Public Property Get getEscolaridade() As String
    getEscolaridade = me_escolaridade
End Property

Public Property Let setEscolaridade(ByVal escolaridade As String)
    me_escolaridade = escolaridade
End Property

Public Property Get getEstadoCivil() As String
    getEstadoCivil = me_estadoCivil
End Property

Public Property Let setEstadoCivil(ByVal estadoCivil As String)
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

Public Property Get getDependentes() As Collection
    Set getDependentes = me_dependentes
End Property

Public Function toArray() As Variant
    Dim enum_parentesco As New EnumParentesco
    Dim enum_escolaridade As New EnumEscolaridade
    
   toArray = Array(me_key, _
                   me_nome, _
                   me_dataNascimento, _
                   enum_escolaridade.EnumInterface_getNum(me_escolaridade), _
                   me_estadoCivil, _
                   me_profissao, _
                   me_cpf, _
                   me_telefone)
End Function


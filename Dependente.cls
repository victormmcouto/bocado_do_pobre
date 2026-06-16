VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dependente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private me_nome As String
Private me_dataNascimento As Date
Private me_escolaridade As String
Private me_parentesco As String

Public Function Create(ByVal nome As String, _
                       ByVal dataNascimento As Date, _
                       ByVal escolaridade As String, _
                       ByVal parentesco As String) As Dependente
    
    me_nome = nome
    me_dataNascimento = dataNascimento
    me_escolaridade = escolaridade
    me_parentesco = parentesco
    
    Set Create = Me
End Function

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

Public Property Get getParentesco() As String
    getParentesco = me_parentesco
End Property

Public Property Let setParentesco(ByVal parentesco As String)
    me_parentesco = parentesco
End Property

Public Function toArray() As Variant
    Set toArray = Array(me_nome, me_dataNascimento, me_escolaridade, me_parentesco)
End Function



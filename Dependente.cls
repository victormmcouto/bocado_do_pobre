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
Private me_escolaridade As Integer
Private me_parentesco As Integer
Private me_assistido As assistido
Private me_key As Integer

Public Function Create(ByVal nome As String, _
                       ByVal dataNascimento As Date, _
                       ByVal escolaridade As Integer, _
                       ByVal parentesco As Integer, _
                       ByRef astd As assistido) As dependente
    
    me_nome = nome
    me_dataNascimento = dataNascimento
    me_escolaridade = escolaridade
    me_parentesco = parentesco
    Set me_assistido = astd
    
    Set Create = Me
End Function

Public Function CreateFromArray(ByVal arr As Variant) As dependente
    me_key = arr(1, 1)
    me_nome = arr(1, 2)
    me_dataNascimento = arr(1, 3)
    me_escolaridade = arr(1, 4)
    me_parentesco = arr(1, 5)
    
    Set CreateFromArray = Me
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

Public Property Get getEscolaridade() As Integer
    getEscolaridade = me_escolaridade
End Property

Public Property Let setEscolaridade(ByVal escolaridade As Integer)
    me_escolaridade = escolaridade
End Property

Public Property Get getParentesco() As Integer
    getParentesco = me_parentesco
End Property

Public Property Let setParentesco(ByVal parentesco As Integer)
    me_parentesco = parentesco
End Property

Public Property Get getAssistido() As assistido
    getAssistido = me_assistido
End Property

Public Property Set setAssistido(ByVal astd As assistido)
    Set me_assistido = astd
End Property

Public Function toArray() As Variant
    Dim enum_parentesco As New EnumParentesco
    Dim enum_escolaridade As New EnumEscolaridade
    
    toArray = Array(me_key, _
                    me_nome, _
                    me_dataNascimento, _
                    me_escolaridade, _
                    me_parentesco, _
                    me_assistido.getKey)
End Function

Public Function toString(Optional ByVal showExtra As Boolean = True) As String
    Dim escolaridade As New EnumEscolaridade
    Dim parentesco As New EnumParentesco
    
    toString = "Key             = " & me_key & ", " & vbNewLine & _
               "Nome            = " & me_nome & ", " & vbNewLine & _
               "Data Nascimento = " & me_dataNascimento & ", " & vbNewLine & _
               "Escolaridade    = " & escolaridade.Enums_getNome(me_escolaridade) & ", " & vbNewLine & _
               "Parentesco      = " & parentesco.Enums_getNome(me_parentesco) & vbNewLine

    If showExtra Then
        toString = toString & vbNewLine & "ASSISTIDO: " & vbNewLine & me_assistido.toString(False)
    End If
End Function



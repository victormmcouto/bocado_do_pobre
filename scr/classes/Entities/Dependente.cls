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
Private me_keyAssistido As Long
Private me_key As Long

Private Sub Class_Initialize()
    me_nome = vbNullString
    me_dataNascimento = 0
    me_escolaridade = 0
    me_parentesco = 0
    me_keyAssistido = 0
    me_key = 0
End Sub

Public Function create(ByVal nome As String, _
                       ByVal dataNascimento As Date, _
                       ByVal escolaridade As Integer, _
                       ByVal parentesco As Integer) As Dependente
    
    me_nome = nome
    me_dataNascimento = dataNascimento
    me_escolaridade = escolaridade
    me_parentesco = parentesco
    
    Set create = Me
End Function

Public Function CreateFromArray(ByVal arr As Variant) As Dependente
    me_key = arr(1, 1)
    me_nome = arr(1, 2)
    me_dataNascimento = arr(1, 3)
    me_escolaridade = arr(1, 4)
    me_parentesco = arr(1, 5)
    me_keyAssistido = arr(1, 6)
    
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

Public Property Get getParentesco() As Integer
    getParentesco = me_parentesco
End Property

Public Property Let setParentesco(ByVal parentesco As Integer)
    me_parentesco = parentesco
End Property

Public Property Get getKeyAssistido() As Long
    getKeyAssistido = me_keyAssistido
End Property

Public Property Let setKeyAssistido(ByVal key As Long)
    me_keyAssistido = key
End Property

Public Function clone() As Dependente
    Dim novo As New Dependente
    
    novo.CreateFromArray Me.toArray
    
    Set clone = novo
End Function

Public Function Equals(ByRef outro As Dependente) As Boolean
    If outro Is Nothing Then Exit Function
    
    If Me.getKey = 0 Or outro.getKey = 0 Then
        Equals = (Me Is outro)
        Exit Function
    End If
    
    Equals = (Me.getKey = outro.getKey)
End Function

Public Function toArray() As Variant
    Dim keyAstd As Long
       
    toArray = Array(me_key, _
                    me_nome, _
                    me_dataNascimento, _
                    me_escolaridade, _
                    me_parentesco, _
                    me_keyAssistido)
End Function

Public Function toString(Optional ByVal showExtra As Boolean = True) As String
    Dim escolaridade As New enumEscolaridade
    Dim parentesco As New enumParentesco
    
    toString = "Key             = " & me_key & ", " & vbNewLine & _
               "Nome            = " & me_nome & ", " & vbNewLine & _
               "Data Nascimento = " & me_dataNascimento & ", " & vbNewLine & _
               "Escolaridade    = " & escolaridade.Enums_getNome(me_escolaridade) & ", " & vbNewLine & _
               "Parentesco      = " & parentesco.Enums_getNome(me_parentesco) & ", " & vbNewLine & _
               "Key Assistido   = " & me_keyAssistido & vbNewLine
End Function



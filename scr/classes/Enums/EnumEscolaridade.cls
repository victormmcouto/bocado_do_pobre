VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumEscolaridade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Enums

Const me_ANALFABETO As Integer = 1
Const me_FUNDAMENTAL_INCOMPLETO As Integer = 2
Const me_FUNDAMENTAL_COMPLETO As Integer = 3
Const me_MEDIO_INCONPLETO As Integer = 4
Const me_MEDIO_COMPLETO As Integer = 5
Const me_CURSO_TECNICO As Integer = 6
Const me_SUPERIOR_INCOMPLETO As Integer = 7
Const me_SUPERIOR_COMPLETO As Integer = 8
Const me_ESPECIALIZACAO As Integer = 9
Const me_MESTRADO As Integer = 10
Const me_DOUTORADO As Integer = 11
Const me_POS_DOUTORADO As Integer = 12

Private arr(1 To 12) As Variant

Private Sub Class_Initialize()
    arr(1) = "Analfabeto"
    arr(2) = "Ensino Fundamental Incompleto"
    arr(3) = "Ensino Fundamental Completo"
    arr(4) = "Ensino Médio Incompleto"
    arr(5) = "Ensino Médio Completo"
    arr(6) = "Curso Técnico"
    arr(7) = "Ensino Superior Incompleto"
    arr(8) = "Ensino Superior Completo"
    arr(9) = "Especialização"
    arr(10) = "Mestrado"
    arr(11) = "Doutorado"
    arr(12) = "Pós-Doutorado"
End Sub

Public Property Get ANALFABETO() As Integer
    ANALFABETO = me_ANALFABETO
End Property
Public Property Get FUNDAMENTAL_INCOMPLETO() As Integer
    FUNDAMENTAL_INCOMPLETO = me_FUNDAMENTAL_INCOMPLETO
End Property
Public Property Get FUNDAMENTAL_COMPLETO() As Integer
    FUNDAMENTAL_COMPLETO = me_FUNDAMENTAL_COMPLETO
End Property
Public Property Get MEDIO_INCONPLETO() As Integer
    MEDIO_INCONPLETO = me_MEDIO_COMPLETO
End Property
Public Property Get MEDIO_COMPLETO() As Integer
    MEDIO_COMPLETO = me_MEDIO_COMPLETO
End Property
Public Property Get CURSO_TECNICO() As Integer
    CURSO_TECNICO = me_CURSO_TECNICO
End Property
Public Property Get SUPERIOR_INCOMPLETO() As Integer
    SUPERIOR_INCOMPLETO = me_SUPERIOR_INCOMPLETO
End Property
Public Property Get SUPERIOR_COMPLETO() As Integer
    SUPERIOR_COMPLETO = me_SUPERIOR_COMPLETO
End Property
Public Property Get ESPECIALIZACAO() As Integer
    ESPECIALIZACAO = me_ESPECIALIZACAO
End Property
Public Property Get MESTRADO() As Integer
    MESTRADO = me_MESTRADO
End Property
Public Property Get DOUTORADO() As Integer
    DOUTORADO = me_DOUTORADO
End Property
Public Property Get POS_DOUTORADO() As Integer
    POS_DOUTORADO = me_POS_DOUTORADO
End Property

Public Function Enums_getNome(ByVal num As Integer) As String
    Enums_getNome = arr(num)
End Function

Public Function Enums_getNum(ByVal nome As String) As Integer
    Dim index As Integer
    
    For index = 1 To UBound(arr)
        If arr(index) = nome Then Enums_getNum = index: Exit Function
    Next index
End Function

Public Function Enums_getArray() As Variant
    Enums_getArray = arr
End Function


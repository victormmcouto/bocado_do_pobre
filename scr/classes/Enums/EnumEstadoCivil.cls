VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumEstadoCivil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Enums

Const me_SOLTEIRO_A As Integer = 1
Const me_CASADO_A As Integer = 2
Const me_SEPARADO_JUDI As Integer = 3
Const me_DIVORCIADO_A As Integer = 4
Const me_VIUVO_A As Integer = 5

Private arr(1 To 5) As Variant

Private Sub Class_Initialize()
    arr(1) = "Solteiro(a)"
    arr(2) = "Casado(a)"
    arr(3) = "Separado(a) Judicialmente"
    arr(4) = "Divorciado(a)"
    arr(5) = "Viúvo(a)"
End Sub

Public Property Get SOLTEIRO_A() As Integer
    SOLTEIRO_A = me_SOLTEIRO_A
End Property
Public Property Get CASADO_A() As Integer
    CASADO_A = me_CASADO_A
End Property
Public Property Get SEPARADO_JUDI() As Integer
    SEPARADO_JUDI = me_SEPARADO_JUDI
End Property
Public Property Get DIVORCIADO_A() As Integer
    DIVORCIADO_A = me_DIVORCIADO_A
End Property
Public Property Get VIUVO_A() As Integer
    VIUVO_A = me_VIUVO_A
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



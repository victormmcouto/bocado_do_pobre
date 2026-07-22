VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumTipoMoradia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Enums

Const me_NULO As Integer = 0
Const me_PROPRIA As Integer = 1
Const me_ALUGADA As Integer = 2
Const me_CEDIDA As Integer = 3

Private arr(0 To 3) As Variant

Private Sub Class_Initialize()
    arr(0) = vbNullString
    arr(1) = "Própria"
    arr(2) = "Alugada"
    arr(3) = "Cedida"
End Sub

Public Property Get PROPRIA() As Integer
    PROPRIA = me_PROPRIA
End Property
Public Property Get ALUGADA() As Integer
    ALUGADA = me_ALUGADA
End Property
Public Property Get CEDIDA() As Integer
    CEDIDA = me_CEDIDA
End Property
Public Property Get NULO() As Integer
    NULO = me_NULO
End Property

Public Function Enums_getNome(ByVal num As Integer) As String
    Enums_getNome = arr(num)
End Function

Public Function Enums_getNum(ByVal nome As String) As Integer
    Dim index As Integer
    
    For index = 0 To UBound(arr)
        If arr(index) = nome Then Enums_getNum = index: Exit Function
    Next index
End Function

Public Function Enums_getArray() As Variant
    Enums_getArray = arr
End Function




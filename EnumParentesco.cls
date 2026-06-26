VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumParentesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Enums

Const me_FILHO_A As Integer = 1
Const me_NETO_A As Integer = 2
Const me_BISNETO_A As Integer = 3
Const me_PAI_MAE As Integer = 4
Const me_AVO As Integer = 5
Const me_BISAVO As Integer = 6
Const me_IRMAO_A As Integer = 7
Const me_TIO_A As Integer = 8
Const me_SOBRINHO_A As Integer = 9
Const me_PRIMO_A As Integer = 10
Const me_ENTEADO_A As Integer = 11
Const me_CUNHADO_A As Integer = 12
Const me_OUTRO As Integer = 13

Private arr(1 To 13) As Variant

Private Sub Class_Initialize()
        arr(1) = "Filho(a)"
        arr(2) = "Neto(a)"
        arr(3) = "Bisneto(a)"
        arr(4) = "Pai/Mãe"
        arr(5) = "Avô/Avó"
        arr(6) = "Bisavô/Bisavó"
        arr(7) = "Irmão/Irmã"
        arr(8) = "Tio/Tia"
        arr(9) = "Sobrinho/Sobrinha"
        arr(10) = "Primo/Prima"
        arr(11) = "Enteado/Enteada"
        arr(12) = "Cunhado/Cunhada"
        arr(13) = "Outro"
End Sub

Public Property Get FILHO_A() As Integer
    FILHO_A = me_FILHO_A
End Property
Public Property Get NETO_A() As Integer
    NETO_A = me_NETO_A
End Property
Public Property Get BISNETO_A() As Integer
    BISNETO_A = me_BISNETO_A
End Property
Public Property Get PAI_MAE() As Integer
    PAI_MAE = me_PAI_MAE
End Property
Public Property Get AVO() As Integer
    AVO = me_AVO
End Property
Public Property Get BISAVO() As Integer
    BISAVO = me_BISAVO
End Property
Public Property Get IRMAO_A() As Integer
    IRMAO_A = me_IRMAO_A
End Property
Public Property Get TIO_A() As Integer
    TIO_A = me_TIO_A
End Property
Public Property Get SOBRINHO_A() As Integer
    SOBRINHO_A = me_SOBRINHO_A
End Property
Public Property Get PRIMO_A() As Integer
    PRIMO_A = me_PRIMO_A
End Property
Public Property Get ENTEADO_A() As Integer
    ENTEADO_A = me_ENTEADO_A
End Property
Public Property Get CUNHADO_A() As Integer
    CUNHADO_A = me_CUNHADO_A
End Property
Public Property Get outro() As Integer
    outro = me_OUTRO
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

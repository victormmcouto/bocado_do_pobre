VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Acompanhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private me_key As Long
Private me_keyCadastro As Long
Private me_ano As Integer
Private me_meses(1 To 12) As Boolean

Public Function create(ByVal meses As Variant, _
                       ByVal ano As Integer) As Acompanhamento
    
    Dim index As Integer
    
    For index = 1 To UBound(meses)
        me_meses(index) = meses(index)
    Next index
    
    me_ano = ano
    
    Set create = Me
End Function

Public Function CreateFromArray(ByVal arr As Variant) As Acompanhamento
    me_key = arr(1, 1)
    me_ano = arr(1, 2)
    me_keyCadastro = arr(1, 3)
    me_meses(1) = arr(1, 4)
    me_meses(2) = arr(1, 5)
    me_meses(3) = arr(1, 6)
    me_meses(4) = arr(1, 7)
    me_meses(5) = arr(1, 8)
    me_meses(6) = arr(1, 9)
    me_meses(7) = arr(1, 10)
    me_meses(8) = arr(1, 11)
    me_meses(9) = arr(1, 12)
    me_meses(10) = arr(1, 13)
    me_meses(11) = arr(1, 14)
    me_meses(12) = arr(1, 15)
    
    Set CreateFromArray = Me
End Function

Public Property Get getKey() As Long
    getKey = me_key
End Property

Public Property Let setKey(ByVal key As Long)
    me_key = key
End Property

Public Property Get getMeses() As Variant
    getMeses = me_meses
End Property

Public Property Get getAno() As Integer
    getAno = me_ano
End Property

Public Property Let setAno(ByVal ano As Long)
    setAno = ano
End Property

Public Property Get getKeyCadastro() As Long
    getKeyCadastro = me_keyCadastro
End Property

Public Property Let setKeyCadastro(ByVal keyCadastro As Long)
    me_keyCadastro = keyCadastro
End Property

Public Sub marcarEntrega(ByVal mesNum As Integer)
    me_meses(mesNum) = True
End Sub

Public Function foiEntregue(ByVal mesNum As Integer) As Boolean
    foiEntregue = me_meses(mesNum)
End Function

Public Function Equals(ByRef outro As Acompanhamento) As Boolean
    If outro Is Nothing Then Exit Function
    
    If Me.getKey = 0 Or outro.getKey = 0 Then
        Equals = (Me Is outro)
        Exit Function
    End If
    
    Equals = (Me.getKey = outro.getKey)
End Function

Public Function toArray() As Variant
    Dim keyConjg As Long
    
    toArray = Array(me_key, _
                    me_ano, _
                    me_keyCadastro, _
                    me_meses(1), _
                    me_meses(2), _
                    me_meses(3), _
                    me_meses(4), _
                    me_meses(5), _
                    me_meses(6), _
                    me_meses(7), _
                    me_meses(8), _
                    me_meses(9), _
                    me_meses(10), _
                    me_meses(11), _
                    me_meses(12))
End Function

Public Function toString(Optional ByVal showExtra As Boolean = True) As String
    toString = "Key             = " & me_key & ", " & vbNewLine & _
               "Ano             = " & me_ano & ", " & vbNewLine & _
               "Key Cadastro    = " & me_keyCadastro & ", " & vbNewLine & _
               "JANEIRO         = " & me_meses(1) & ", " & vbNewLine & _
               "FEVEREIRO       = " & me_meses(2) & ", " & vbNewLine & _
               "MARÇO           = " & me_meses(3) & ", " & vbNewLine & _
               "ABRIL           = " & me_meses(4) & ", " & vbNewLine & _
               "MAIO            = " & me_meses(5) & ", " & vbNewLine & _
               "JUNHO           = " & me_meses(6) & ", " & vbNewLine & _
               "JULHO           = " & me_meses(7) & ", " & vbNewLine & _
               "AGOSTO          = " & me_meses(8) & ", " & vbNewLine & _
               "SETEMBRO        = " & me_meses(9) & ", " & vbNewLine & _
               "OUTUBRO         = " & me_meses(10) & ", " & vbNewLine & _
               "NOVEMBRO        = " & me_meses(11) & ", " & vbNewLine & _
               "DEZEMBRO        = " & me_meses(12)
End Function


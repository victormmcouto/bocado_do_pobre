VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private me_key As Long
Private me_assistido As Assistido
Private me_logradouro As String
Private me_numero As String
Private me_bairro As String
Private me_cidade As String
Private me_dataSindicancia As Date
Private me_nomeVisitador As String
Private me_participaProgGov As Boolean
Private me_progGov As String
Private me_tipoMoradia As Integer
Private me_nDependentes As Integer
Private me_recebeCesta As Boolean

Private me_acompanhamentos As Collection

Public Function create(ByRef astd As Assistido, _
                       ByVal logradouro As String, _
                       ByVal numero As String, _
                       ByVal bairro As String, _
                       ByVal cidade As String, _
                       ByVal dataSindicancia As Date, _
                       ByVal nomeVisitador As String, _
                       ByVal participaProgGov As Boolean, _
                       ByVal progGov As String, _
                       ByVal tipoMoradia As Integer, _
                       ByVal nDependentes As Integer, _
                       ByVal recebeCesta As Boolean) As cadastro
    
    Set me_assistido = astd
    me_logradouro = logradouro
    me_numero = numero
    me_bairro = bairro
    me_cidade = cidade
    me_dataSindicancia = dataSindicancia
    me_nomeVisitador = nomeVisitador
    me_participaProgGov = participaProgGov
    me_progGov = progGov
    me_tipoMoradia = tipoMoradia
    me_nDependentes = nDependentes
    me_recebeCesta = recebeCesta
    
    Set create = Me
End Function

Public Function CreateFromArray(ByVal arr As Variant) As cadastro
    me_key = arr(1, 1)
    me_logradouro = arr(1, 3)
    me_numero = arr(1, 4)
    me_bairro = arr(1, 5)
    me_cidade = arr(1, 6)
    me_dataSindicancia = arr(1, 7)
    me_nomeVisitador = arr(1, 8)
    me_participaProgGov = arr(1, 9)
    me_progGov = arr(1, 10)
    me_tipoMoradia = arr(1, 11)
    me_nDependentes = arr(1, 12)
    me_recebeCesta = arr(1, 13)
    
    Set CreateFromArray = Me
End Function

Public Property Get getAssistido() As Assistido
    Set getAssistido = me_assistido
End Property

Public Property Set setAssistido(ByVal astd As Assistido)
    Set me_assistido = astd
End Property

Public Property Set setAcompanhamentos(ByRef acompanhamentos As Collection)
    Set me_acompanhamentos = acompanhamentos
End Property

Public Sub AddAcompanhamento(ByRef acompanhamento As acompanhamento)
    me_acompanhamentos.Add acompanhamento, acompanhamento.getAno
End Sub

Public Function getAcompanhamento(ByVal ano As Integer) As acompanhamento
    Set getAcompanhamento = me_acompanhamentos(CStr(ano))
End Function

Public Function toArray() As Variant
    toArray = Array(me_key, _
                    me_assistido.getKey, _
                    me_logradouro, _
                    me_numero, _
                    me_bairro, _
                    me_cidade, _
                    me_dataSindicancia, _
                    me_nomeVisitador, _
                    me_participaProgGov, _
                    me_progGov, _
                    me_tipoMoradia, _
                    me_nDependentes, _
                    me_recebeCesta)
End Function

Public Property Let setKey(ByVal key As Long)
    me_key = key
End Property

Public Property Get getKey() As Long
    getKey = me_key
End Property

Public Function Equals(ByRef outro As cadastro) As Boolean
    If outro Is Nothing Then Exit Function
    
    If Me.getKey = 0 Or outro.getKey = 0 Then
        Equals = (Me Is outro)
        Exit Function
    End If
    
    Equals = (Me.getKey = outro.getKey)
End Function

Public Function toString(Optional ByVal showExtra As Boolean = True) As String
    Dim tipoMoradia As New EnumTipoMoradia
    
    toString = "DADOS CADASTRAIS:" & vbNewLine
    
    toString = toString & _
               "Key                               = " & me_key & "," & vbNewLine & _
               "Key Assistido                     = " & getAssistido.getKey & "," & vbNewLine & _
               "Logradouro                        = " & me_logradouro & "," & vbNewLine & _
               "Número                            = " & me_numero & "," & vbNewLine & _
               "Bairro                            = " & me_bairro & "," & vbNewLine & _
               "Cidade                            = " & me_cidade & "," & vbNewLine & _
               "Data Sindicância                  = " & me_dataSindicancia & "," & vbNewLine & _
               "Nome Visitador                    = " & me_nomeVisitador & "," & vbNewLine & _
               "Participa de Programa do Governo? = " & me_participaProgGov & "," & vbNewLine & _
               "Programa do Governo               = " & me_progGov & "," & vbNewLine & _
               "Tipo de Moradia                   = " & tipoMoradia.Enums_getNome(me_tipoMoradia) & "," & vbNewLine & _
               "Nº dependentes                    = " & me_nDependentes & "," & vbNewLine & _
               "Recebe cesta básica?              = " & me_recebeCesta & vbNewLine & vbNewLine
    
    If showExtra Then
        toString = toString & me_assistido.toString(True) & vbNewLine & vbNewLine
        
        Dim acmpmnt As acompanhamento
        
        For Each acmpmnt In me_acompanhamentos
            toString = toString & acmpmnt.toString(True) & vbNewLine & vbNewLine
        Next acmpmnt
    End If
End Function

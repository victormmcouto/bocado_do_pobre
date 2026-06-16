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

Private Enum Parentesco
    FILHO_A = 1
    NETO_A = 2
    BISNETO_A = 3
    PAI_MAE = 4
    AVO = 5
    BISAVO = 6
    IRMAO_A = 7
    TIO_A = 8
    SOBRINHO_A = 9
    PRIMO_A = 10
    ENTEADO_A = 11
    CUNHADO_A = 12
    OUTRO = 13
End Enum

Public Property Get getParentesco() As Parentesco
    getParentesco = Parentesco
End Property

Public Function getNomeParentesco(ByVal num As Integer) As String
    Select Case num
        Case FILHO_A: getNomeParentesco = "Filho(a)"
        Case NETO_A: getNomeParentesco = "Neto(a)"
        Case BISNETO_A: getNomeParentesco = "Bisneto(a)"
        Case PAI_MAE: getNomeParentesco = "Pai/Mãe"
        Case AVO: getNomeParentesco = "Avô/Avó"
        Case BISAVO: getNomeParentesco = "Bisavô/Bisavó"
        Case IRMAO_A: getNomeParentesco = "Irmão/Irmã"
        Case TIO_A: getNomeParentesco = "Tio/Tia"
        Case SOBRINHO_A: getNomeParentesco = "Sobrinho/Sobrinha"
        Case PRIMO_A: getNomeParentesco = "Primo/Prima"
        Case ENTEADO_A: getNomeParentesco = "Enteado/Enteada"
        Case CUNHADO_A: getNomeParentesco = "Cunhado/Cunhada"
        Case OUTRO: getNomeParentesco = "Outro"
    End Select
End Function

Public Function getNumParenteco(ByVal nome As String) As Integer
    Select Case nome
        Case "Filho(a)": getNumParenteco = FILHO_A
        Case "Neto(a)": getNumParenteco = NETO_A
        Case "Bisneto(a)": getNumParenteco = BISNETO_A
        Case "Pai/Mãe": getNumParenteco = PAI_MAE
        Case "Avô/Avó": getNumParenteco = AVO
        Case "Bisavô/Bisavó": getNumParenteco = BISAVO
        Case "Irmão/Irmã": getNumParenteco = IRMAO_A
        Case "Tio/Tia": getNumParenteco = TIO_A
        Case "Sobrinho/Sobrinha": getNumParenteco = SOBRINHO_A
        Case "Primo/Prima": getNumParenteco = PRIMO_A
        Case "Enteado/Enteada": getNumParenteco = ENTEADO_A
        Case "Cunhado/Cunhada": getNumParenteco = CUNHADO_A
        Case "Outro": getNumParenteco = OUTRO
    End Select
End Function

Public Function getdictParentesco() As Dictionary
    With getdictParentesco
        .Add getNomeParentesco(1), 1
        .Add getNomeParentesco(2), 2
        .Add getNomeParentesco(3), 3
        .Add getNomeParentesco(4), 4
        .Add getNomeParentesco(5), 5
        .Add getNomeParentesco(6), 6
        .Add getNomeParentesco(7), 7
        .Add getNomeParentesco(8), 8
        .Add getNomeParentesco(9), 9
        .Add getNomeParentesco(10), 10
        .Add getNomeParentesco(11), 11
        .Add getNomeParentesco(12), 12
        .Add getNomeParentesco(13), 13
    End With
End Function

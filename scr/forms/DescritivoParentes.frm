VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DescritivoParentes 
   Caption         =   "DESCRITIVO DE PARENTESCO"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   OleObjectBlob   =   "DescritivoParentes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DescritivoParentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SpinButtonParente_Change()
    Me.Frame1.Caption = "Parente " & SpinButtonParente.Value
    
    With cdst.getAssistido.getDependente(SpinButtonParente.Value)
        txtbNomeParente.Value = .getNome
        combEscolaridadeParente.Value = enumEscolaridade.Enums_getNome(.getEscolaridade)
        combGrauParentescoParente.Value = enumParentesco.Enums_getNome(.getParentesco)
        If .getDataNascimento = 0 Then
            txtbDataNascimentoParente.Value = ""
        Else
            txtbDataNascimentoParente.Value = .getDataNascimento
        End If
    End With
End Sub

Private Sub txtbDataNascimentoParente_AfterUpdate()
    With txtbDataNascimentoParente
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarDataCompleta(.Value) Then
            cdst.getAssistido.getDependente(SpinButtonParente.Value).setDataNascimento = Format(.Value, "dd/mm/yyyy")
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbNomeParente_Change()
    cdst.getAssistido.getDependente(SpinButtonParente.Value).setNome = txtbNomeParente.Value
End Sub

Private Sub combEscolaridadeParente_Change()
    cdst.getAssistido.getDependente(SpinButtonParente.Value).setEscolaridade = enumEscolaridade.Enums_getNum(combEscolaridadeParente.Value)
End Sub

Private Sub combGrauParentescoParente_Change()
    cdst.getAssistido.getDependente(SpinButtonParente.Value).setParentesco = enumParentesco.Enums_getNum(combGrauParentescoParente.Value)
End Sub

Private Sub UserForm_Initialize()
    Dim totalPessoas As Integer
    
    SpinButtonParente.Value = 1
    SpinButtonParente.Min = 1
    SpinButtonParente.Max = cdst.getAssistido.getTotalDependentes
    
    combEscolaridadeParente.List = enumEscolaridade.Enums_getArray
    combGrauParentescoParente.List = enumParentesco.Enums_getArray
End Sub


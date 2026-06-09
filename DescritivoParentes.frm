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
    
    With Cadastro.parentes(SpinButtonParente.Value)
        txtbNomeParente.Value = .Nome
        combEscolaridadeParente.Value = .Escolaridade
        combGrauParentescoParente.Value = .GrauParentesco
        If .DataNascimento = 0 Then
            txtbDataNascimentoParente.Value = ""
        Else
            txtbDataNascimentoParente.Value = .DataNascimento
        End If
    End With
End Sub

Private Sub txtbDataNascimentoParente_AfterUpdate()
    With txtbDataNascimentoParente
        If .Value = "" Then Exit Sub
        On Error GoTo ErrHandler
        
        If ValidarMaiorDeIdade(.Value) Then
            Cadastro.parentes(SpinButtonParente.Value).DataNascimento = Format(.Value, "dd/mm/yyyy")
        End If
        
        Exit Sub
ErrHandler:
        MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground, Err.Source
        .Value = ""
    End With
End Sub

Private Sub txtbNomeParente_Change()
    Cadastro.parentes(SpinButtonParente.Value).Nome = txtbNomeParente.Value
End Sub

Private Sub combEscolaridadeParente_Change()
    Cadastro.parentes(SpinButtonParente.Value).Escolaridade = combEscolaridadeParente.Value
End Sub

Private Sub combGrauParentescoParente_Change()
    Cadastro.parentes(SpinButtonParente.Value).GrauParentesco = combGrauParentescoParente.Value
End Sub

Private Sub UserForm_Initialize()
    Dim totalPessoas As Integer
    
    SpinButtonParente.Value = 1
    SpinButtonParente.Min = 1
    SpinButtonParente.Max = Cadastro.DemaisInfo.NPessoasNaCasa
    
    Call PopulateComboBoxes
End Sub

Private Sub PopulateComboBoxes()
    Dim tblEscolaridadeParente As ListObject
    Dim tblGrausParentesco As ListObject
    
    With ThisWorkbook
        Set tblEscolaridadeParente = wksESCOLARIDADES.ListObjects(1)
        Set tblGrausParentesco = wksGRAUS_PARENTESCO.ListObjects(1)
    End With
    
    Populate tblEscolaridadeParente.DataBodyRange, combEscolaridadeParente
    
    Populate tblGrausParentesco.DataBodyRange, combGrauParentescoParente
End Sub


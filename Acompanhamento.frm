VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Acompanhamento 
   Caption         =   "ACOMPANHAMENTO DE ASSISTIDOS"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "Acompanhamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Acompanhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IgnorarEventos As Boolean
Private dicCkbMeses As New Dictionary
Private keyAcompanhamento As Integer

Private Sub ckbJan_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbJan
End Sub

Private Sub ckbFev_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbFev
End Sub

Private Sub ckbMar_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbMar
End Sub

Private Sub ckbAbr_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbAbr
End Sub

Private Sub ckbMai_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbMai
End Sub

Private Sub ckbJun_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbJun
End Sub

Private Sub ckbJul_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbJul
End Sub

Private Sub ckbAgo_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbAgo
End Sub

Private Sub ckbSet_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbSet
End Sub

Private Sub ckbOut_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbOut
End Sub

Private Sub ckbNov_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbNov
End Sub

Private Sub ckbDez_Click()
    If IgnorarEventos Then Exit Sub
    ValidateCheckBox ckbDez
End Sub

Private Sub combNomeAssistido_Change()
    If combNomeAssistido.Text = "" Then Exit Sub
    
    frmCestasMes.Enabled = True
    
    Cadastro = CadastroVazio
    
    Set listRowComparativo = GetCadastroListRow(combNomeAssistido.Text)
            
    If Not listRowComparativo Is Nothing Then
        keyAcompanhamento = CInt(GetTableCellValue(listRowComparativo, "KeyAcompanhamento"))
        
        PopulateTypeCadastro
        
        txtbTelefone.Value = Cadastro.Assistido.Telefone
        txtbEndereco.Value = Cadastro.Endereco.Logradouro & ", " & Cadastro.Endereco.NumeroCasa
        txtbBairro.Value = Cadastro.Endereco.Bairro
        
        ChangeValueMesesTo , True
    Else
        txtbTelefone.Value = "------"
        txtbEndereco.Value = "------"
        txtbBairro.Value = "------"
        
        ChangeValueMesesTo False
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim tblCadastros As ListObject
    Dim control As MSForms.control
    
    Set tblCadastros = wksCADASTROS.ListObjects(1)
    
    Populate tblCadastros.ListColumns("NomeAssistido").DataBodyRange, combNomeAssistido
    
    txtbTelefone.Value = "------"
    txtbEndereco.Value = "------"
    txtbBairro.Value = "------"
    
    'Realça o mês atual, bloqueia edição dos demais
    For Each control In frmCestasMes.Controls
        With control
            If TypeName(control) = "CheckBox" Then
                If SiglaMesParaNum(LCase(.Caption)) < Month(Date) Then
                    .Locked = True
                ElseIf SiglaMesParaNum(LCase(.Caption)) > Month(Date) Then
                    .Enabled = False
                Else
                    .ForeColor = &HFF&
                End If
                
                dicCkbMeses.Add .Caption, control
            End If
        End With
    Next control
    
    frmCestasMes.Enabled = False
End Sub

Private Sub ChangeValueMesesTo(Optional ByVal state As Boolean, Optional ByVal typeMeses As Boolean = False)
    IgnorarEventos = True
    If typeMeses Then
        With Cadastro.Acompanhamento
            ckbJan.Value = .Jan
            ckbFev.Value = .Fev
            ckbMar.Value = .Mar
            ckbAbr.Value = .Abr
            ckbMai.Value = .Mai
            ckbJun.Value = .Jun
            ckbJul.Value = .Jul
            ckbAgo.Value = .Ago
            ckbSet.Value = .Set
            ckbOut.Value = .Out
            ckbNov.Value = .Nov
            ckbDez.Value = .Dez
        End With
    Else
        Dim control As MSForms.control
        For Each control In frmCestasMes.Controls
            With control
                .Value = state
            End With
        Next control
    End If
    IgnorarEventos = False
End Sub

Private Function ValidateCheckBox(ckb As MSForms.CheckBox) As Boolean
    Dim result As VbMsgBoxResult
    
    IgnorarEventos = True
    
    If ckb.Value Then
        result = MsgBox("Deseja mesmo marcar a cesta do mês como RECEBIDA?" & vbNewLine & _
                        "Marcar como RECEBIDA fará com que não seja possível alterar o status posteriormente!", _
                        vbInformation + vbYesNo, _
                        "Importante!")
        
        If result = vbYes Then
            IgnorarEventos = False
            SetTableCellValue GetAcompanhamentoListRow(keyAcompanhamento), Capitalize(ckb.Caption), ckb.Value
            Exit Function
        Else
            ckb.Value = False
        End If
    Else
        ckb.Value = True
                
        MsgBox "Não há possibilidade de alterar o valor já marcado do mês de " & ckb.Caption & "!", _
               vbInformation + vbOKOnly, _
               "Mês já marcado!"
    End If
    
    IgnorarEventos = False
End Function

Private Sub ChangeCheckBoxMesState(ByVal strNomeMes As String, ByVal state As Boolean)
    With Cadastro.Acompanhamento
        Select Case strNomeMes
            Case "jan": .Jan = state
            Case "fev": .Fev = state
            Case "mar": .Mar = state
            Case "abr": .Abr = state
            Case "mai": .Mai = state
            Case "jun": .Jun = state
            Case "jul": .Jul = state
            Case "ago": .Ago = state
            Case "set": .Set = state
            Case "out": .Out = state
            Case "nov": .Nov = state
            Case "dez": .Dez = state
        End Select
    End With
End Sub

Private Function SiglaMesParaNum(ByVal strSigla As String) As Integer
    Select Case strSigla
        Case "jan": SiglaMesParaNum = 1
        Case "fev": SiglaMesParaNum = 2
        Case "mar": SiglaMesParaNum = 3
        Case "abr": SiglaMesParaNum = 4
        Case "mai": SiglaMesParaNum = 5
        Case "jun": SiglaMesParaNum = 6
        Case "jul": SiglaMesParaNum = 7
        Case "ago": SiglaMesParaNum = 8
        Case "set": SiglaMesParaNum = 9
        Case "out": SiglaMesParaNum = 10
        Case "nov": SiglaMesParaNum = 11
        Case "dez": SiglaMesParaNum = 12
    End Select
End Function

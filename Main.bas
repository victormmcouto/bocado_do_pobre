Attribute VB_Name = "Main"
Option Explicit

'Public cadastroComparativo As Cadastro
Public listRowComparativo As ListRow
Public boolCadastrar As Boolean
'Public nomeAssistido As String
Public keyParenteAntesDeletar As String

Public Sub Cadastrar(ByVal control As IRibbonControl)
    Dim formCadastro As New CadastroDeAssistidos
    
    Cadastro = CadastroVazio
    
    boolCadastrar = True
    
    formCadastro.Show
End Sub

Public Sub Atualizar(ByVal control As IRibbonControl)
    Dim formProcurarCadastro As New ProcurarCadastro
    Dim formCadastro As New CadastroDeAssistidos
    
    formProcurarCadastro.Show
        
    If Not listRowComparativo Is Nothing Then
        boolCadastrar = False
        
        Cadastro = CadastroVazio
        
        If Not listRowComparativo Is Nothing Then
            PopulateTypeCadastro
            formCadastro.Show
        End If
    End If
End Sub

Public Sub Deletar(ByVal control As IRibbonControl)
    Dim formProcurarCadastro As New ProcurarCadastro
    Dim result As VbMsgBoxResult
    Dim nomeAssistido As String
    
    formProcurarCadastro.Show
    
    If Not listRowComparativo Is Nothing Then
        nomeAssistido = listRowComparativo.Range(1, listRowComparativo.Parent.ListColumns("NomeAssistido").index).Value
        
        result = MsgBox("Deseja mesmo prosseguir com a exclusão de cadastro do(a) " & nomeAssistido & "?" & _
                        vbNewLine & "Todos os dados cadastrados serão excluídos!", _
                        vbCritical + vbYesNo, _
                        "ATENÇÃO!!")
        
        If result = vbYes Then
            Call DeletarCadastro
            MsgBox "Cadastro no nome do(a) " & nomeAssistido & " deletado completamente!", _
                   vbInformation + vbOKOnly, _
                   "Deleção Concluída!"
        Else
            MsgBox "Cadastro no nome de " & nomeAssistido & " preservado!", _
                   vbInformation + vbOKOnly, _
                   "Nada mudou!"
        End If
    End If
End Sub

Public Sub BuscarCadastro(ByVal control As IRibbonControl)
    Dim formProcurarCadastro As New ProcurarCadastro
    Dim nomeAssistido As String
    
    formProcurarCadastro.Show
    
    If Not listRowComparativo Is Nothing Then
        nomeAssistido = listRowComparativo.Range(1, listRowComparativo.Parent.ListColumns("NomeAssistido").index).Value
        
        wksFICHA_CADASTRAL.Range("NomeAssistido").Value = nomeAssistido
        wksFICHA_CADASTRAL.PrintPreview
    End If
End Sub


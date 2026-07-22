Attribute VB_Name = "Main"
Option Explicit

'Public cadastroComparativo As Cadastro
Public listRowComparativo As ListRow
Public boolCadastrar As Boolean
'Public nomeAssistido As String
Public keyParenteAntesDeletar As String
Public keyAcompanhamentoAntesDeletar As String

Public repoCadastro As RepositoryCadastro
Public repoAssistido As RepositoryAssistido
Public repoConjuge As RepositoryConjuge
Public repoDependente As RepositoryDependente
Public repoAcompanhamento As RepositoryAcompanhamento

Public enumEstadoCivil As New enumEstadoCivil
Public enumEscolaridade As New enumEscolaridade
Public enumParentesco As New enumParentesco
Public enumTipoMoradia As New enumTipoMoradia

Public cdst As Cadastro

Public Sub Cadastrar(ByVal control As IRibbonControl)
    Dim formCadastro As New CadastroDeAssistidos
    
    Set cdst = New Cadastro
    
    Call startRepos
    
    boolCadastrar = True
    
    formCadastro.Show
    
    Call exitRepos
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
'        Dim tempWkb As Workbook
'        Set tempWkb = Workbooks.Add()
'        wksFICHA_CADASTRAL.Copy tempWkb.Worksheets(1)
'        tempWkb.Worksheets(1).Visible = True
'        tempWkb.Worksheets(1).PrintPreview
'        tempWkb.Close False
        wksFICHA_CADASTRAL.PrintPreview
    End If
End Sub

Public Sub AcompanhamentoMensal(ByVal control As IRibbonControl)
    Dim formAcompanhamento As New acompanhamento
    
    formAcompanhamento.Show
    
    Unload formAcompanhamento
End Sub

Private Sub startRepos()
    Set repoCadastro = New RepositoryCadastro
    Set repoAssistido = New RepositoryAssistido
    Set repoConjuge = New RepositoryConjuge
    Set repoDependente = New RepositoryDependente
    Set repoAcompanhamento = New RepositoryAcompanhamento
    
    Set repoCadastro.setRepoAssistiddo = repoAssistido
    Set repoCadastro.setRepoAcompanhamento = repoAcompanhamento
    
    Set repoAssistido.setRepoConjuge = repoConjuge
    Set repoAssistido.setRepoDependente = repoDependente
    
    Set repoConjuge.setRepoAssistido = repoAssistido
    
    Set repoDependente.setRepoAssistido = repoAssistido
End Sub

Private Sub exitRepos()
    Set repoCadastro = Nothing
    Set repoAssistido = Nothing
    Set repoConjuge = Nothing
    Set repoDependente = Nothing
    Set repoAcompanhamento = Nothing
End Sub



Attribute VB_Name = "Main"
Option Explicit

'Public cadastroComparativo As Cadastro
Public listRowComparativo As ListRow
Public boolCadastrar As Boolean

Public Sub Cadastrar()
    Dim formCadastro As New CadastroDeAssistidos
    
    boolCadastrar = True
    
    formCadastro.Show
End Sub

Public Sub Atualizar()
    Dim formProcurarCadastro As New ProcurarCadastro
    Dim formCadastro As New CadastroDeAssistidos
    
    boolCadastrar = False
    
    formProcurarCadastro.Show
    
    Cadastro = CadastroVazio
    
    If Not listRowComparativo Is Nothing Then
        PopulateTypeCadastro
        formCadastro.Show
    End If
End Sub

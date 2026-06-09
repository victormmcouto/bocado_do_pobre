Attribute VB_Name = "Main"
Option Explicit

'Public cadastroComparativo As Cadastro
Public listRowComparativo As ListRow

Public Sub Cadastrar()
    Dim formCadastro As New CadastroDeAssistidos
    
    formCadastro.Show
End Sub

Public Sub Atualizar()
    Dim formProcurarCadastro As New ProcurarCadastro
    
    formProcurarCadastro.Show
    
    Cadastro = CadastroVazio
    
    If Not listRowComparativo Is Nothing Then
        PopulateTypeCadastro
    End If
End Sub

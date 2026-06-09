VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProcurarCadastro 
   Caption         =   "PROCURAR CADASTRO"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5850
   OleObjectBlob   =   "ProcurarCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProcurarCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rngNomesAssistidos As Range
Private tblCadastros As ListObject

Private Sub cbttProcurar_Click()
    Dim lr As ListRow
    Dim bolEcontrado As Boolean
    Dim nomeAProcurar As String
    Dim nomeAtual As String
    
    nomeAProcurar = Cadastro.Assistido.Nome
    
    For Each lr In tblCadastros.ListRows
        nomeAtual = lr.Range(1, tblCadastros.ListColumns("NomeAssistido").index).Value
        
        If nomeAtual = nomeAProcurar Then
            Set listRowComparativo = lr
            bolEcontrado = True
        End If
    Next lr
    
    If Not bolEcontrado Then
        MsgBox "Não encontrado"
        Set listRowComparativo = Nothing
    Else
        Call PopulateTypeCadastro
    End If
End Sub

Private Sub combNomeAssistido_change()
    Cadastro.Assistido.Nome = combNomeAssistido.Value
End Sub

Public Sub UserForm_Initialize()
    Set tblCadastros = wksCADASTROS.ListObjects(1)
    Set rngNomesAssistidos = tblCadastros.ListColumns("NomeAssistido").DataBodyRange
    
    Call Populate(rngNomesAssistidos, combNomeAssistido)
End Sub

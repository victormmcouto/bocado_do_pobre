VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroDeAssistidos 
   Caption         =   "FICHA CADASTRAL DOS ASSISTIDOS"
   ClientHeight    =   8640.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12450
   OleObjectBlob   =   "CadastroDeAssistidos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CadastroDeAssistidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EspecificacaoProgGovFed_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    EspecificacaoProgGovFed.DropDown
End Sub

Private Sub UserForm_Initialize()
    With Me.ComboBox9
        For i = 1 To 20
            .AddItem i
        Next i
    End With
End Sub

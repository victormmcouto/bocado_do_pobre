VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroDeAssistidos 
   Caption         =   "FICHA CADASTRAL DOS ASSISTIDOS"
   ClientHeight    =   7770
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
Private Sub CommandButton1_Click()
    total = CInt(ComboBox9.Value)
    call_ total
End Sub

Private Sub EspecificacaoProgGovFed_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    EspecificacaoProgGovFed.DropDown
End Sub

Private Sub UserForm_Initialize()
    Call PopulateComboBoxes 'Popula as combo box com os valores armazenados nas tabelas de dados

    imAutorizacao.Picture = LoadPicture(vbNullString)
End Sub

Private Sub OptBttProgramaGovFedNAO_Click()
    Image1.Picture = LoadPicture(Application.GetOpenFilename(, , "Selecione a declaração"))
End Sub

Public Sub PopulateComboBoxes()
    Dim tblProfissoes As ListObject
    Dim tblEstadosCivis As ListObject
    Dim tblEscolaridades As ListObject
    Dim tblProgramaGov As ListObject
    Dim arrTbls As Variant
    Dim arrCombDadosPessoais As Variant
    
    With ThisWorkbook
        Set tblProfissoes = wksPROFISSOES.ListObjects(1)
        Set tblEstadosCivis = wksESTADOS_CIVIS.ListObjects(1)
        Set tblEscolaridades = wksESCOLARIDADES.ListObjects(1)
        Set tblProgramaGov = wksPROGRAMAS_GOV.ListObjects(1)
    End With
    
    Populate tblProfissoes, combProfissaoAssistido
    Populate tblProfissoes, combProfissaoConjugue
    
    Populate tblEstadosCivis, combEstadoCivilAssistido
    Populate tblEstadosCivis, combEstadoCivilConjugue
    
    Populate tblEscolaridades, combEscolaridadeAssistido
    Populate tblEscolaridades, combEscolaridadeConjugue
    
    Populate tblProgramaGov, combEspecificacaoProgGov
End Sub

Private Sub Populate(ByRef tbl As ListObject, ByRef comb As ComboBox)
    Dim item As Range
    
    For Each item In tbl.DataBodyRange.Cells
        comb.AddItem item.Value
    Next item
End Sub

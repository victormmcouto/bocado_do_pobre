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

Private count As Integer

Private Sub cbttAdd_Click()
    With cadastro.parentes(CInt(SpinButtonParente))
        .Nome = txtbNomeParente.Value
        .Escolaridade = combEscolaridadeParente.Value
        .GrauParentesco = combGrauParentescoParente.Value
        .DataNascimento = txtbDataNascimentoParente.Value
    End With
End Sub

Private Sub SpinButtonParente_Change()
    Me.Frame1.Caption = "Parente " & SpinButtonParente.Value
    
    With cadastro.parentes(CInt(SpinButtonParente))
        txtbNomeParente.Value = .Nome
        combEscolaridadeParente.Value = .Escolaridade
        combGrauParentescoParente.Value = .GrauParentesco
        txtbDataNascimentoParente.Value = .DataNascimento
    End With
End Sub

Private Sub UserForm_Initialize()
    If ParentesInicializado(cadastro.parentes) Then ReDim cadastro.parentes(1 To total)
    
    count = 1
    SpinButtonParente.Value = count
    SpinButtonParente.Min = count
    SpinButtonParente.Max = total
    
    Call PopulateComboBoxes
End Sub

Private Sub PopulateComboBoxes()
    Dim tblEscolaridadeParente As ListObject
    Dim tblGrausParentesco As ListObject
    
    With ThisWorkbook
        Set tblEscolaridadeParente = wksESTADOS_CIVIS.ListObjects(1)
        Set tblGrausParentesco = wksGRAUS_PARENTESCO.ListObjects(1)
    End With
    
    Populate tblEscolaridadeParente, combEscolaridadeParente
    
    Populate tblGrausParentesco, combGrauParentescoParente
End Sub


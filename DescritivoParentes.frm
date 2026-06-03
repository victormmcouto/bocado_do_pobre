VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DescritivoParentes 
   Caption         =   "DESCRITIVO DE PARENTESCO"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
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

Private Sub CommandButton1_Click()
    If count = total Then count = 1 Else count = count + 1
    Me.Frame1.Caption = "Parente " & count
End Sub

Private Sub CommandButton2_Click()
    If count <= 1 Then count = total Else count = count - 1
    Me.Frame1.Caption = "Parente " & count
End Sub

Private Sub SpinButton1_Change()
    Me.Frame1.Caption = "Parente " & SpinButton1.Value
End Sub

Private Sub UserForm_Initialize()
    count = 1
    SpinButton1.Value = count
    SpinButton1.Min = count
    SpinButton1.Max = total
    Me.Frame1.Caption = "Parente " & SpinButton1.Value
End Sub

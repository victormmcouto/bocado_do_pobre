VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Repository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Create(ByVal tbl As ListObject) As Repository
End Function

Public Function factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
End Function

Public Function getTbl() As TableManagement
End Function

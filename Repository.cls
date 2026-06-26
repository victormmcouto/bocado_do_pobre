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

Public Function factory(ByVal key As Integer, Optional ByVal complete As Boolean = True) As Object
End Function

Public Function getTbl() As TableManagement
End Function

Public Function save(ByRef obj As Object, Optional ByVal autoIncrement As Boolean = True) As Boolean
End Function

Public Function readById(ByVal key As Integer) As Object
End Function

Public Function updateById(ByVal key As Integer, ByRef obj As Object) As Boolean
End Function

Public Function deleteById(ByVal key As Integer) As Boolean
End Function

Public Function deleteFkey(ByVal fk As Integer, ByVal fkName As String) As Boolean
End Function


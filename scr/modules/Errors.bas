Attribute VB_Name = "Errors"
Option Explicit

Public Enum Errs
    errInvalidKey = vbObjectError + 1
    errFormatoNumeroTelefoneInvalido = vbObjectError + 2
    errFormatoDataInvalida = vbObjectError + 3
    errValidacaoCPF = vbObjectError + 4
    errCamposObrigatorios = vbObjectError + 5
End Enum

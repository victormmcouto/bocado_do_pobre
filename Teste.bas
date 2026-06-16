Attribute VB_Name = "Teste"
Option Explicit

'Teste da classe Dependente e da Persistência de dados para ADD
Sub test1()
    Debug.Print "TESTE 1"
    Dim dpnt As New dependente
    Dim tblmgmt As New TableManagement
    Dim escolaridade As New EnumEscolaridade
    Dim parentesco As New EnumParentesco
    Dim estadoCivil As New EnumEstadoCivil
    
    Dim astd As New assistido
    
    Set dpnt = dpnt.Create("Allan", _
                           CDate("13/04/1997"), _
                           escolaridade.MEDIO_COMPLETO, _
                           parentesco.FILHO_A, _
                           astd.Create("João", _
                                       CDate("19/08/2001"), _
                                       escolaridade.DOUTORADO, _
                                       estadoCivil.DIVORCIADO_A, _
                                       "Joalheiro", _
                                       "11234253501", _
                                       "314523424"))
    astd.setKey = 5
                           
    Set tblmgmt = tblmgmt.Create(wksDEPENDENTE.ListObjects(1))
    
    tblmgmt.InsertDataRow dpnt
End Sub

'Teste EnumParentesco
Sub test2()
    Debug.Print "TESTE 2"
    Dim parentesco As New EnumParentesco
    Debug.Print parentesco.Enums_getNome(3)
    Debug.Print parentesco.Enums_getNum("Avô/Avó")
    Debug.Print parentesco.CUNHADO_A
End Sub

'Teste EnumEscolaridade
Sub test3()
    Debug.Print "TESTE 3"
    Dim escolaridade As New EnumEscolaridade
    Debug.Print escolaridade.Enums_getNome(3)
    Debug.Print escolaridade.Enums_getNum("Doutorado")
    Debug.Print escolaridade.FUNDAMENTAL_COMPLETO
End Sub

'Teste Assistido com Dependente
Sub test4()
    Debug.Print "TESTE 4"
    Dim astd As New assistido
    Dim dpnt As New dependente
    Dim escolaridade As New EnumEscolaridade
    Dim parentesco As New EnumParentesco
    
    Dim tblmgmt1 As New TableManagement
    Dim tblmgmt2 As New TableManagement
        
    Set astd = astd.Create("Gustavo", _
                           CDate("06/07/1967"), _
                           escolaridade.MEDIO_INCONPLETO, _
                           "Solteiro", _
                           "Marceneiro", _
                           "11897630803", _
                           "31956735208")
    
    With astd.getDependentes
        .Add dpnt.Create("Allan", _
                         CDate("02/03/1999"), _
                         escolaridade.FUNDAMENTAL_INCOMPLETO, _
                         parentesco.BISAVO, _
                         astd)
    End With
    
    Set tblmgmt1 = tblmgmt1.Create(wksASSISTIDO.ListObjects(1))
    Set tblmgmt2 = tblmgmt2.Create(wksDEPENDENTE.ListObjects(1))
    
    tblmgmt1.InsertDataRow astd
    tblmgmt2.InsertDataRow dpnt
    
    tblmgmt1.InsertDataRow astd
    tblmgmt2.InsertDataRow dpnt
End Sub

'Teste RepositoryDependente (Criação e objeto via tabela)
Sub test5()
    Debug.Print "TESTE 5"
    Dim dpnt As New dependente
    Dim repoDpnt As New RepositoryDependente
    
    
    Set repoDpnt = repoDpnt.Repository_Create(wksDEPENDENTE.ListObjects(1))
    Set dpnt = repoDpnt.Repository_factory(9)
    
    With dpnt
        Debug.Print .getKey
        Debug.Print .getNome
        Debug.Print .getDataNascimento
        Debug.Print .getEscolaridade
        Debug.Print .getParentesco
    End With
End Sub

'Teste Repositories (Criação de Objetos Assistido e Dependente via Tbela)
Sub test6()
    Debug.Print "TEST 6"
    Dim dpnt As New dependente
    Dim astd As New assistido
    Dim repoDpnt As New RepositoryDependente
    Dim repoAstd As New RepositoryAssistido
    
    Set repoDpnt = repoDpnt.Repository_Create(wksDEPENDENTE.ListObjects(1))
    Set repoAstd = repoAstd.Repository_Create(wksASSISTIDO.ListObjects(1))
    
    Set dpnt = repoDpnt.Repository_factory(25)
    Set astd = repoAstd.Repository_factory(1)
    
    Debug.Print dpnt.toString
    Debug.Print vbNewLine
    Debug.Print astd.toString
End Sub

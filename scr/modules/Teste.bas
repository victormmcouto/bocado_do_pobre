Attribute VB_Name = "Teste"
Option Explicit

Private repoCadastro As RepositoryCadastro
Private repoAssistido As RepositoryAssistido
Private repoConjuge As RepositoryConjuge
Private repoDependente As RepositoryDependente

'Teste da classe Dependente e da Persistência de dados para ADD
Sub test1()
    Debug.Print "TESTE 1"
    Dim dpnt As New Dependente
    Dim tblmgmt As New TableManagement
    Dim escolaridade As New EnumEscolaridade
    Dim parentesco As New EnumParentesco
    Dim estadoCivil As New EnumEstadoCivil
    
    Dim astd As New Assistido
    
    Set dpnt = dpnt.create("Allan", _
                           CDate("13/04/1997"), _
                           escolaridade.MEDIO_COMPLETO, _
                           parentesco.FILHO_A, _
                           astd.create("João", _
                                       CDate("19/08/2001"), _
                                       escolaridade.DOUTORADO, _
                                       estadoCivil.DIVORCIADO_A, _
                                       "Joalheiro", _
                                       "11234253501", _
                                       "314523424"))
    astd.setKey = 5
                           
    Set tblmgmt = tblmgmt.create(wksDEPENDENTE.ListObjects(1))
    
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
    Dim astd As New Assistido
    Dim dpnt As New Dependente
    Dim escolaridade As New EnumEscolaridade
    Dim parentesco As New EnumParentesco
    
    Dim tblmgmt1 As New TableManagement
    Dim tblmgmt2 As New TableManagement
        
    Set astd = astd.create("Gustavo", _
                           CDate("06/07/1967"), _
                           escolaridade.MEDIO_INCONPLETO, _
                           "Solteiro", _
                           "Marceneiro", _
                           "11897630803", _
                           "31956735208")
    
    With astd.getDependentes
        .Add dpnt.create("Allan", _
                         CDate("02/03/1999"), _
                         escolaridade.FUNDAMENTAL_INCOMPLETO, _
                         parentesco.BISAVO, _
                         astd)
    End With
    
    Set tblmgmt1 = tblmgmt1.create(wksASSISTIDO.ListObjects(1))
    Set tblmgmt2 = tblmgmt2.create(wksDEPENDENTE.ListObjects(1))
    
    tblmgmt1.InsertDataRow astd
    tblmgmt2.InsertDataRow dpnt
    
    tblmgmt1.InsertDataRow astd
    tblmgmt2.InsertDataRow dpnt
End Sub

'Teste RepositoryDependente (Criação e objeto via tabela)
Sub test5()
    Debug.Print "TESTE 5"
    Dim dpnt As New Dependente
    Dim repoDpnt As New RepositoryDependente
    
    
    Set repoDpnt = repoDpnt.Repository_create(wksDEPENDENTE.ListObjects(1))
    Set dpnt = repoDpnt.Repository_readByKey(9)
    
    With dpnt
        Debug.Print .getKey
        Debug.Print .getNome
        Debug.Print .getDataNascimento
        Debug.Print .getEscolaridade
        Debug.Print .getParentesco
    End With
End Sub

'Teste Repositories (Criação de Objetos Assistido e Dependente via Tabela)
Sub test6()
    Debug.Print "TEST 6"
    Dim dpnt As New Dependente
    Dim astd As New Assistido
    Dim repoDpnt As New RepositoryDependente
    Dim repoAstd As New RepositoryAssistido
    
    Set repoDpnt = repoDpnt.Repository_create(wksDEPENDENTE.ListObjects(1))
    Set repoAstd = repoAstd.Repository_create(wksASSISTIDO.ListObjects(1))
    
    Set dpnt = repoDpnt.Repository_readByKey(25, True)
    Set astd = repoAstd.Repository_readByKey(2, True)
    
    Debug.Print dpnt.toString
    Debug.Print vbNewLine
    Debug.Print astd.toString
End Sub

'Teste Conjugue
Sub test7()
    Debug.Print "TEST 7"
    Dim cnjg As New Conjuge
    Dim cnjg1 As New Conjuge
    Dim repoCnjg As New RepositoryConjuge
    Dim escolaridade As New EnumEscolaridade
    Dim estadoCivil As New EnumEstadoCivil
        
    Set repoCnjg = repoCnjg.Repository_create(wksCONJUGE.ListObjects(1))
    
    Set cnjg = repoCnjg.Repository_readByKey(2)
    
    Debug.Print (cnjg.toString(False))
    
    Set cnjg1 = cnjg1.create("Alessandra", _
                             CDate("18/08/2010"), _
                             escolaridade.CURSO_TECNICO, _
                             estadoCivil.CASADO_A, _
                             "Vidraceira", _
                             "45609812378", _
                             "31962430936")
    
    repoCnjg.Repository_getTbl.InsertDataRow cnjg1
End Sub

'Teste de Cadastro
Sub tes8()
    Debug.Print "TESTE 8"
    Dim cdst As cadastro
    
    startRepos
    
    Set cdst = repoCadastro.Repository_readByKey(2)
    
    Debug.Print cdst.toString(False)
    
    exitRepos
End Sub

Sub teste9()
    Debug.Print "TESTE 9 - Teste de salvamento de informações de Assistido"
    Dim repoAssistido As New RepositoryAssistido
    Dim astd As New Assistido
    
    Set astd = createAssistido
    
    astd.AddDependente createDependente
    
    Debug.Print repoAssistido.Repository_save(astd)
End Sub

Sub test10()
    Debug.Print "TESTE 10 - Teste de salvamento de informações de Cadastro"
    
    Dim repoCadastro As New RepositoryCadastro
    Dim cdtr As New cadastro
    
    Set cdtr = createCadastro
    
    Debug.Print repoCadastro.Repository_save(cdtr)
End Sub

Sub test11()
    On Error GoTo ErrHandler
    
    Debug.Print "TESTE 11 - Teste de deleção de Dependente por key"
    
    Dim repoDependente As New RepositoryDependente
    
    Debug.Print repoDependente.Repository_deleteById(13)
    'Espera erro
    Debug.Print repoDependente.Repository_deleteById(13)
    
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
End Sub

Sub test12()
    On Error GoTo ErrHandler
    
    Debug.Print "TESTE 12 - Teste de deleção de conjuge por key com preservação de integridade referencial"
    
    startRepos
    
    Debug.Print repoConjuge.Repository_deleteById(5)
    'Espera erro
    Debug.Print repoConjuge.Repository_deleteById(5)

fim:
    exitRepos
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
    GoTo fim
End Sub

Sub test13()
    On Error GoTo ErrHandler
    
    Debug.Print "TESTE 13 - Teste de deleção de Assistido por key com preservação de integridade referencial"
    
    startRepos
    
    Debug.Print repoAssistido.Repository_deleteById(14)
    'Espera erro
    Debug.Print repoAssistido.Repository_deleteById(14)

fim:
    exitRepos
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
    GoTo fim
End Sub

Sub test14()
    On Error GoTo ErrHandler
    
    Debug.Print "TESTE 14 - Teste de deleção de Cadastro por key com preservação de integridade referencial"
    
    startRepos
    
    Debug.Print repoCadastro.Repository_deleteById(2)
    'Espera erro
    Debug.Print repoCadastro.Repository_deleteById(2)

fim:
    exitRepos
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
    GoTo fim
End Sub

Sub test15()
    On Error GoTo ErrHandler
    
    startRepos
    
    Dim newDpnt As Dependente
    
    Debug.Print "TESTE 15 - Teste de atualização de dependente"
    
    Set newDpnt = createDependente
    repoDependente.Repository_updateById 18, newDpnt
fim:
    exitRepos
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
    GoTo fim
End Sub

Sub test16()
    On Error GoTo ErrHandler
    
    startRepos
    
    Dim newCnjg As Conjuge
    
    Debug.Print "TESTE 16 - Teste de atualização de Conjuge"
    
    Set newCnjg = createConjuge
    repoConjuge.Repository_updateById 3, newCnjg
fim:
    exitRepos
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
    GoTo fim
End Sub

Sub test17()
    On Error GoTo ErrHandler
    
    startRepos
    
    Dim newAstd As Assistido
    
    Debug.Print "TESTE 17 - Teste de atualização de Assistido"
    
    Set newAstd = createAssistido
    Dim dpnt2 As Dependente
    Dim dpnt3 As Dependente
    Dim dpnt4 As Dependente
    Dim dpnt5 As Dependente
    
    newAstd.getConjuge.setKey = 4
    
    Set dpnt2 = createDependente
    Set dpnt3 = createDependente
    Set dpnt4 = createDependente
    Set dpnt5 = createDependente
    
    dpnt2.setKey = 2
    dpnt3.setKey = 3
    dpnt4.setKey = 4
    
    newAstd.getDependente(1).setKey = 1 'de 1 à 4 devem apenas atualziar
    newAstd.AddDependente dpnt2
    newAstd.AddDependente dpnt3
    newAstd.AddDependente dpnt4
    newAstd.AddDependente dpnt5 'Sem chave (Dever ser adicionado como chave nova)
    
    newAstd.setKey = 8
    repoAssistido.Repository_updateById 8, newAstd
fim:
    exitRepos
    Exit Sub
ErrHandler:
    MsgBox Err.Source, vbCritical, Err.Description
    GoTo fim
End Sub

Public Sub test18()
    Debug.Print "TESTE 18 de acompanhamento"
    
    Dim acmpmnt As New Acompanhamento
    
    acmpmnt.create Array(True, _
                         True, _
                         False, _
                         True, _
                         False, _
                         True, _
                         True, _
                         True, _
                         True, _
                         True, _
                         False), _
                   2026
    
    Debug.Print acmpmnt.toString
End Sub

Public Sub test19()
    Debug.Print "TESTE 19 de rpository acompanhamento readByKey" & vbNewLine
    
    Dim Reposiotiryacmpmnt As New RepositoryAcompanhamento
    
    Debug.Print Reposiotiryacmpmnt.Repository_readByKey(CLng(2026 & 6)).toString
End Sub

Public Sub test20()
    Debug.Print "TESTE 20 de rpository acompanhamento save" & vbNewLine
    
    Dim Reposiotiryacmpmnt As New RepositoryAcompanhamento
    Dim acmpmnt As New Acompanhamento
    Dim arr(1 To 12) As Boolean
    
    Set acmpmnt = acmpmnt.create(arr, 2027)
    
    acmpmnt.setKeyCadastro = 99
    
    Reposiotiryacmpmnt.Repository_save acmpmnt, False
End Sub

Public Sub test21()
    Debug.Print "TESTE 21 de rpository acompanhamento delete" & vbNewLine
    
    Dim Reposiotiryacmpmnt As New RepositoryAcompanhamento
    
    Reposiotiryacmpmnt.Repository_deleteById 20268
End Sub

Public Sub test22()
    Debug.Print "TESTE 22 de rpository acompanhamento update" & vbNewLine
    
    Dim Reposiotiryacmpmnt As New RepositoryAcompanhamento
    Dim acmpmnt As New Acompanhamento
     
    Set acmpmnt = Reposiotiryacmpmnt.Repository_readByKey(20267)
    acmpmnt.marcarEntrega 6
    
    Reposiotiryacmpmnt.Repository_updateById acmpmnt.getKey, acmpmnt
End Sub

Public Function createAssistido() As Assistido
    Dim astd As New Assistido
    Dim escolaridade As New EnumEscolaridade
    Dim estadoCivil As New EnumEstadoCivil
    
    Set astd = astd.create("Januário Augusto", _
                           CDate("23/06/1996"), _
                           escolaridade.CURSO_TECNICO, _
                           estadoCivil.CASADO_A, _
                           "Marceneiro", _
                           "11878432903", _
                           "31963741296", _
                           createConjuge)
    
    astd.AddDependente createDependente
    
    Set createAssistido = astd
End Function

Public Function createDependente() As Dependente
    Dim dpnt As New Dependente
    Dim parentesco As New EnumParentesco
    Dim escolaridade As New EnumEscolaridade
    
    Set createDependente = dpnt.create("Augusto de Lima", _
                                       CDate("20/09/2007"), _
                                       escolaridade.CURSO_TECNICO, _
                                       parentesco.NETO_A)
End Function

Public Function createConjuge() As Conjuge
    Dim cnjg As New Conjuge
    Dim escolaridade As New EnumEscolaridade
    Dim estadoCivil As New EnumEstadoCivil
    
    Set createConjuge = cnjg.create("Débora Augusta", _
                                    CDate("03/04/1998"), _
                                    escolaridade.ESPECIALIZACAO, _
                                    estadoCivil.CASADO_A, _
                                    "Analista de BI", _
                                    "18635609821", _
                                    "31963894234")
End Function

Public Function createCadastro() As cadastro
    Dim cdtr As New cadastro
    Dim tipoMoradia As New EnumTipoMoradia
    Dim astd As New Assistido
    
    Set astd = createAssistido
    
    Set createCadastro = cdtr.create(astd, _
                                     "Rua dos Perdizes", _
                                     "467A", _
                                     "Cidade Bela", _
                                     "Curujás", _
                                     Date, _
                                     "Cirlei", _
                                     True, _
                                     "MCMV", _
                                     tipoMoradia.CEDIDA, _
                                     astd.getTotalDependentes, _
                                     False)
End Function

Public Sub startRepos()
    Set repoCadastro = New RepositoryCadastro
    Set repoAssistido = New RepositoryAssistido
    Set repoConjuge = New RepositoryConjuge
    Set repoDependente = New RepositoryDependente
    
    Set repoCadastro.setRepoAssistiddo = repoAssistido
    
    Set repoAssistido.setRepoConjuge = repoConjuge
    Set repoAssistido.setRepoDependente = repoDependente
    
    Set repoConjuge.setRepoAssistido = repoAssistido
    
    Set repoDependente.setRepoAssistido = repoAssistido
End Sub

Public Sub exitRepos()
    Set repoCadastro = Nothing
    Set repoAssistido = Nothing
    Set repoConjuge = Nothing
    Set repoDependente = Nothing
End Sub


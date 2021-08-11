VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cContribuinte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BuscarContribuinte(IM As String, Optional RetNomeContrib As String, Optional RetTipoLogrContr As String, _
                                                Optional RetLogrContr As String, Optional RetNumeroContr As String, Optional RetCompContri As String, _
                                                Optional RetBairroContr As String, Optional RetCepContr As String, Optional RetMunicContr As String, _
                                                Optional RetUF As String) As Boolean
    Dim sql As String, Rs As VSRecordset
    sql = "Select tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
            " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF from Tab_Contribuinte where tci_im = '" & IM & "'"
    If Bdados.AbreTabela(sql, Rs) Then
        RetNomeContrib = Rs(0)
        RetTipoLogrContr = Rs(1)
        RetLogrContr = Rs(2)
        RetNumeroContr = Rs(3)
        RetCompContri = Rs(4)
        RetBairroContr = Rs(5)
        RetCepContr = Rs(6)
        RetMunicContr = Rs(7)
        RetUF = Rs(8)
        BuscarContribuinte = True
    End If
End Function

Public Function VerificaSitCadastral(IM As String) As String
    Dim sql As String
    Dim Rs As VSRecordset
    sql = "Select tci_tsc_cod_sit_cad from Tab_Contribuinte where tci_im= '" & IM & "'"
        If Bdados.AbreTabela(sql, Rs) Then
            VerificaSitCadastral = Rs!tci_tsc_cod_sit_cad - 1
        End If
End Function

Public Function PreencherCboSitCad(Combo As Object) As Boolean
    Dim sql As String
    sql = "Select tsc_nome, tsc_cod_sit_cad  From Tab_sit_cadastral order by tsc_nome"
    Combo.Preencher Bdados, sql, 0
End Function

Public Function GravarContribuinte(IM As String, NomeContrib As String, TipoLogrContrib As String, NomeLogrContrib As String, _
                                        NumeroContrib As String, CompContrib As String, BairroContrib As String, CepContrib As String, _
                                        MunicContrib As String, UFContrib As String) As Boolean
    Dim Campos As String, Valores As String
    Campos = "tci_im, tci_nome , tci_logradouro, " _
                    & "tci_nome_logradouro, tci_numero, tci_complemento," _
                    & "tci_bairro, tci_cep , tci_cidade, tci_UF, tci_data_cadastro," _
                    & "tci_tsc_cod_sit_cad,tci_tus_cod_usuario,tci_tipo_contribuinte"
    Valores = Bdados.PreparaValor(IM, NomeContrib, TipoLogrContrib, _
                    NomeLogrContrib, NumeroContrib, CompContrib, BairroContrib, CepContrib, MunicContrib, UFContrib, _
                    Bdados.Converte(Date, TCDataHora), 1, Aplicacoes.Usuario, 0)
    If Bdados.InsereDados("Tab_Contribuinte", Valores, Campos) Then
        GravarContribuinte = True
    End If
End Function

Public Function PreencherGrd(Grid As Object, NomeContrib As String) As Boolean
    Dim sql As String
        sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & NomeContrib & "%' or tci_nome like '% " & NomeContrib & "%'"
        sql = sql & " and tci_tsc_cod_sit_cad =1"
        
        If Grid.Preencher(Bdados, sql) Then PreencherGrd = True
End Function

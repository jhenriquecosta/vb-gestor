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
Public Function BuscarContribuinte(Im As String, Optional RetNomeContrib As String, Optional RetTipoLogrContr As String, _
                                                Optional RetLogrContr As String, Optional RetNumeroContr As String, Optional RetCompContri As String, _
                                                Optional RetBairroContr As String, Optional RetCepContr As String, Optional RetMunicContr As String, _
                                                Optional RetUF As String, Optional RetDocumento As String, Optional RetCodAtividade As Long, Optional InscricaoReal As String) As Boolean
    Dim Sql As String, rs As VSRecordset
    Sql = "Select tci_im,tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
            " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF, tci_cgc_cpf, TCI_TAE_CAE from Tab_Contribuinte where 1 = 1 " & _
            " " 'and tci_tsc_cod_sit_cad=1"
        If Im <> "" Then
            Sql = Sql & " AND tci_im = '" & Im & "'"
        End If
    If Bdados.AbreTabela(Sql, rs) Then
        If Not IsMissing(InscricaoReal) Then InscricaoReal = "" & rs!tci_im
        RetNomeContrib = "" & rs!tci_nome
        RetTipoLogrContr = "" & rs!tci_logradouro
        RetLogrContr = "" & rs!tci_nome_logradouro
        RetNumeroContr = "" & rs!tci_numero
        RetCompContri = "" & rs!tci_complemento
        RetBairroContr = "" & rs!tci_bairro
        RetCepContr = "" & rs!tci_cep
        RetMunicContr = "" & rs!tci_cidade
        RetDocumento = "" & rs!tci_cgc_cpf
        RetUF = "" & rs!tci_UF
        RetDocumento = "" & rs!tci_cgc_cpf
        If Not IsNull(rs!TCI_TAE_CAE) Then
            RetCodAtividade = rs!TCI_TAE_CAE
        End If
        BuscarContribuinte = True
    End If
End Function


Public Function PreencherGrid(Grid As Object, Nome As String) As Boolean
    Dim Sql As String
    If Nome <> "" Then
        Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf AS Cpf_Cnpj from Tab_Contribuinte where tci_nome like '" & Nome & "%' or tci_nome like '% " & Nome & "%'"
                Sql = Sql & " and tci_tsc_cod_sit_cad =1"
        If Trim(Nome) = "" Then
            Avisa "Informe o critério da busca."
            Exit Function
        End If
    End If
    If Grid.Preencher(Bdados, Sql) Then PreencherGrid = True
    
End Function

Public Function NomeAtividade(CodAtividade As Long) As String
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = " SELECT Tae_nome, tae_cae from tab_atividade_economica"
    Sql = Sql & " WHERE tae_cae = " & CodAtividade
    If Bdados.AbreTabela(Sql, rs) Then
        NomeAtividade = rs(0).Value
    Else
        NomeAtividade = ""
    End If
    Bdados.FechaTabela rs
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsContribuinte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_InsMunicipal As String

Public Property Get InsMunicipal() As String
    InsMunicipal = m_InsMunicipal
End Property

Public Property Let InsMunicipal(ByVal value As String)
    m_InsMunicipal = value
End Property

Public Function VerificaTEMImovel(InsMunicipal As String) As Boolean
    Dim Sql As String                   'VERIFICA SE TEM IMOVEL CADASTRADO PRA INXCRIÇÃO MUNICIPAL
    Dim Rs As VSRecordset
    Sql = "Select tim_ic from Tab_Imovel where tim_tci_im = '" & InsMunicipal & "'"
    If Bdados.AbreTabela(Sql, Rs) Then VerificaTEMImovel = True
End Function

Public Function VerificaTEMDebito(InsMunicipal As String) As Boolean
    Dim Sql As String                   'VERIFICA SE TEM debito PRA INSCRIÇÃO MUNICIPAL
    Dim Rs As VSRecordset
    Sql = "Select * FROM tab_geracao_tributo WHERE tgt_im = '" & InsMunicipal & "' and tgt_data_pagamento is null"
    If Bdados.AbreTabela(Sql, Rs) Then VerificaTEMDebito = True
End Function

Public Function VerificaTEMDAMPago(InsMunicipal As String) As Boolean
    Dim Sql As String                   'VERIFICA SE TEM IMOVEL CADASTRADO PRA INXCRIÇÃO MUNICIPAL
    Dim Rs As VSRecordset
    Sql = "SELECT tdr_tgt_cod_pagamento FROM TAB_DARM_RECEBIDO WHERE TDR_INSCRICAO = '" & InsMunicipal & "'"
    If Bdados.AbreTabela(Sql, Rs) Then VerificaTEMDAMPago = True
End Function

Public Function Inscricao(s_Inscricao As String) As TipoInsc
    Inscricao = IIf(Len(Trim(Edita.TiraPic(Edita.TiraPic(Edita.TiraPic(s_Inscricao, "."), "-"), "/"))) = 10, InscContrib, InscImovel)
End Function

Public Function BuscaAtividadeContribuinte(Contribuinte As String) As String
    
    Dim Sql As String
    Sql = "Select tci_tae_cae from tab_contribuinte where tci_im = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql) Then
        If Not IsNull(Bdados.Tabela(0)) Then
            BuscaAtividadeContribuinte = Bdados.Tabela(0)
        Else
'            Util.Avisa "Contribuinte sem atividade."
        End If
    End If
    
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub PreencherCboRua(Combo As Object, Optional TipoLogr As String)
    Dim Sql As String
    Sql = "select distinct tci_nome_logradouro from tab_contribuinte" & _
            " where tci_logradouro = '" & TipoLogr & "'  order by tci_nome_logradouro "
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboTipoLogr(Combo As Object)
    Dim Sql As String
    Sql = "Select ttl_nome From Tab_Tipo_Logr order by ttl_nome asc"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherPonto(Combo As Object)
    Dim Sql As String
    Sql = "SELECT tpr_descricao FROM TAB_PONTO_RECEPCAO ORDER BY 1"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboBairro(Combo As Object)
    Dim Sql As String
    Sql = "SELECT TBA_NOME FROM TAB_BAIRRO ORDER BY TBA_NOME"
    Combo.Preencher Bdados, Sql
End Sub

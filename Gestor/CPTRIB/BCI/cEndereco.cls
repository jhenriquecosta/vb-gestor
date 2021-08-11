VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Call Edita.AtualizaCombo(Bdados, cboLogr, "Select DISTINCT(tlg_nome) From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio)
'    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr")
'    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select DISTINCT(tba_nome) From Tab_Bairro where tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio)

Public Function PreencherComboTipoLogr(Combo As Object) As Boolean
    Dim sql As String
    sql = "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr"
    If Combo.Preencher(Bdados, sql) Then PreencherComboTipoLogr = True
End Function

Public Function PreencherComboLogr(Combo As Object) As Boolean
    Dim sql As String
    sql = "Select DISTINCT(tlg_nome) From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    If Combo.Preencher(Bdados, sql) Then PreencherComboLogr = True
End Function

Public Function PreencherComboBairro(Combo As Object) As Boolean
    Dim sql As String
    sql = "Select DISTINCT(tba_nome) From Tab_Bairro where tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    If Combo.Preencher(Bdados, sql) Then PreencherComboBairro = True
End Function

Public Function BuscaBairro(Bairro As String) As String
    BuscaBairro = Nvl("" & BuscaCodigo("SELECT TBA_COD_BAIRRO FROM TAB_BAIRRO WHERE TBA_TMU_COD_MUNICIPIO=" & Aplicacoes.Codigo_Municipio & " AND TBA_NOME='" & Bairro & "'"), 1)
End Function

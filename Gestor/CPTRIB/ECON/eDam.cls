VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eDam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BuscaDam(Documento As Double, Optional RetCodReceita, Optional RetIm, Optional RetParcela, Optional RetIC, Optional RetPeriodo, Optional RetDtVenc, Optional RetImposto, Optional RetTaxas, Optional RetJuros, Optional RetMulta) As Boolean
    Dim sql As String
    Dim rs As VSRecordset
    sql = "Select * from tab_geracao_tributo where tgt_cod_pagamento=" & Documento
    If Bdados.AbreTabela(sql, rs) Then
        RetCodReceita = "" & rs!tgt_tip_cod_imposto
        RetIm = "" & rs!tgt_im
        RetParcela = "" & rs!tgt_parcela
        RetIC = "" & rs!tgt_tim_ic
        RetPeriodo = "" & rs!tgt_periodo
        RetDtVenc = "" & rs!tgt_data_vencimento
        RetImposto = "" & rs!tgt_valor_tributo
        RetTaxas = "" & rs!tgt_taxa_expediente
        RetJuros = "" & rs!tgt_valor_juros
        RetMulta = "" & rs!tgt_Valor_multa
        BuscaDam = True
    Else
        sql = "Select * from tab_geracao_tributo_parcela where tgt_cod_pagamento=" & Documento
        If Bdados.AbreTabela(sql, rs) Then
            RetCodReceita = "" & rs!tgt_tip_cod_imposto
            RetIm = "" & rs!tgt_im
            RetParcela = "" & rs!tgt_parcela
            RetIC = "" & rs!tgt_tim_ic
            RetPeriodo = "" & rs!tgt_periodo
            RetDtVenc = "" & rs!tgt_data_vencimento
            RetImposto = "" & rs!tgt_valor_tributo
            RetTaxas = "" & rs!tgt_taxa_expediente
            RetJuros = "" & rs!tgt_valor_juros
            RetMulta = "" & rs!tgt_Valor_multa
            BuscaDam = True
        End If
    End If
End Function

Public Sub PreencherCboTributo(Combo As Object)
    Dim sql As String
    sql = "select tip_sigla_imposto + '" & " - " & "' + tip_cod_imposto + '" & " - " & "' + tip_nome_imposto, tip_cod_imposto, tip_sigla_imposto, tip_nome_imposto" & _
            " from tab_imposto order by tip_sigla_imposto"
    Combo.Preencher Bdados, sql
End Sub

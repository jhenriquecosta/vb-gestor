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
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select * from tab_geracao_tributo where tgt_cod_pagamento=" & Documento
    If Bdados.AbreTabela(Sql, Rs) Then
        RetCodReceita = "" & Rs!tgt_tip_cod_imposto
        RetIm = "" & Rs!tgt_im
        RetParcela = "" & Rs!tgt_parcela
        RetIC = "" & Rs!tgt_tim_ic
        RetPeriodo = "" & Rs!tgt_periodo
        RetDtVenc = "" & Rs!tgt_data_vencimento
        RetImposto = "" & Rs!tgt_valor_tributo
        RetTaxas = "" & Rs!tgt_taxa_expediente
        RetJuros = "" & Rs!tgt_valor_juros
        RetMulta = "" & Rs!tgt_Valor_multa
        BuscaDam = True
    Else
        Sql = "Select * from tab_geracao_tributo_parcela where tgt_cod_pagamento=" & Documento
        If Bdados.AbreTabela(Sql, Rs) Then
            RetCodReceita = "" & Rs!tgt_tip_cod_imposto
            RetIm = "" & Rs!tgt_im
            RetParcela = "" & Rs!tgt_parcela
            RetIC = "" & Rs!tgt_tim_ic
            RetPeriodo = "" & Rs!tgt_periodo
            RetDtVenc = "" & Rs!tgt_data_vencimento
            RetImposto = "" & Rs!tgt_valor_tributo
            RetTaxas = "" & Rs!tgt_taxa_expediente
            RetJuros = "" & Rs!tgt_valor_juros
            RetMulta = "" & Rs!tgt_Valor_multa
            BuscaDam = True
        End If
    End If
End Function

Public Sub PreencherCboTributo(Combo As Object)
    Dim Sql As String
    Sql = "select tip_sigla_imposto + '" & " - " & "' + tip_cod_imposto + '" & " - " & "' + tip_nome_imposto, tip_cod_imposto, tip_sigla_imposto, tip_nome_imposto" & _
            " from tab_imposto order by tip_sigla_imposto"
    Combo.Preencher Bdados, Sql
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCertidao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enuTipoCertidao
    tipCND = 1
    tipCPND = 2
    tipCPD = 3
    tipEspecial = 4
End Enum

Private Function verificarCND(Emissao As String, Optional CodImposto As String, Optional Im As String, Optional Ic As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As String
'Expedicao: NAO for constatado a existencia de creditos NAO vencidos em curso de cobranca executiva
            'que NAO tenha sido efetivada a penhora
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT tgt_cod_pagamento AS DAM, " & _
                " tip_sigla_imposto AS Imposto, " & _
                " tgt_periodo as Periodo, " & _
                " tgt_im AS IM, " & _
                " tgt_tim_ic AS IC, " & _
                " tgt_Valor_tributo as Valor, " & _
                " tgt_data_vencimento AS Vencimento, " & _
            " FROM TAB_GERACAO_TRIBUTO, TAB_IMPOSTO " & _
            " WHERE tgt_cod_pagamento NOT IN (" & _
                                            " SELECT tdr_tgt_cod_pagamento " & _
                                            " FROM TAB_DARM_RECEBIDO " & _
                                            " WHERE tgt_cod_pagamento=tdr_tgt_cod_pagamento" & _
                                            " ) " & _
                " AND tgt_tip_cod_imposto NOT IN('" & Const_Extrato & "', '" & Const_Notificacao & "')" & _
                " AND tgt_data_vencimento<=" & Bdados.FormataValorCampo(Emissao) & _
                " AND tip_cod_imposto=tgt_tip_cod_imposto"
    
    If Trim$(CodImposto) <> "" Then
        Sql = Sql & " AND tgt_tip_cod_imposto='" & CodImposto & "'"
    End If
    If Trim$(Im) <> "" Then
        Sql = Sql & " AND tgt_im='" & Im & "'"
    End If
    If Trim$(Ic) <> "" Then
        Sql = Sql & " AND tgt_tim_ic='" & Ic & "'"
    End If
    If Trim$(PeriodoInicial) <> "" Then
        Sql = Sql & " AND tgt_periodo>=" & PeriodoInicial
    End If
    If Trim$(PeriodoFinal) <> "" Then
        Sql = Sql & " AND tgt_periodo<=" & PeriodoFinal
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        verificarCND = Sql
    Else
        verificarCND = ""
    End If
    Bdados.FechaTabela Rs
End Function

Private Function verificarCPND(Emissao As String, Optional CodImposto As String, Optional Im As String, Optional Ic As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As String
'Expedicao: for constatado a existencia de creditos NAO vencidos em curso de cobranca executiva
            'que tenha sido efetivada a penhora
End Function

Private Function verificarCPD(Emissao As String, Optional CodImposto As String, Optional Im As String, Optional Ic As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As String
'Expedicao: for constatado a existencia de creditos vencidos em curso de cobranca executiva
            'que NAO tenha sido efetivada a penhora
End Function

Private Function verificarEspecial(Emissao As String, Optional CodImposto As String, Optional Im As String, Optional Ic As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As String
'Expedicao: emitida quando é constado o cumprimento das obrigacoes tributarias apesar da remessa
            'do banco ainda nao ter chegado (o funcionario assume o risco)
End Function

Public Function verificarCertidao(Tipo As enuTipoCertidao, Emissao As String, Optional CodImposto As String, Optional Im As String, Optional Ic As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String) As Boolean
    Select Case Tipo
        Case tipCND
            verificarCND Emissao, CodImposto, Im, Ic, PeriodoInicial, PeriodoFinal
        Case tipCPD
        Case tipCPND
        Case tipEspecial
    End Select
End Function

Public Function dataValidade(Tipo As enuTipoCertidao, Emissao As String) As String
    Dim Validade As Integer
    
    If Emissao <> "" Then
        If IsDate(Emissao) Then
            Select Case Tipo
                Case tipCND: Validade = 60
                Case tipCPND: Validade = 30
                Case tipCPD: Validade = 90
                Case tipEspecial: Validade = 0
            End Select
            dataValidade = DateAdd("d", Validade, Emissao)
        End If
    End If
End Function

Public Function recuperarTexto(Tipo As enuTipoCertidao) As String
    Dim Sql As String, parametro As String
    Dim Rs As VSRecordset
    
    Select Case Tipo
        Case tipCND: parametro = "CARTA OMISSOS"
        Case tipCPND: parametro = ""
        Case tipCPD: parametro = ""
        Case tipEspecial: parametro = ""
    End Select
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = '" & parametro & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        recuperarTexto = "" & Rs!TPT_TEXTO
    End If
    Bdados.FechaTabela Rs
End Function

Public Function exibirCertidoes(Tipo As enuTipoCertidao, Optional Imposto As String, Optional Im As String, Optional EmissaoInicio As String, Optional EmissaoFim As String, Optional NumInicio As String, Optional NumFim As String) As String
    Select Case Tipo
        Case tipCND: exibirCertidoes = exibirCNDs(Imposto, Im, EmissaoInicio, EmissaoFim, NumInicio, NumFim)
        Case tipCPND: exibirCertidoes = ""
        Case tipCPD: exibirCertidoes = ""
        Case tipEspecial: exibirCertidoes = ""
    End Select
End Function
Private Function exibirCNDs(Optional Imposto As String, Optional Im As String, Optional EmissaoInicio As String, Optional EmissaoFim As String, Optional NumInicio As String, Optional NumFim As String) As String
    Dim Sql As String
    
    Sql = "SELECT " & _
                " TNT_COD_NOTIFICACAO AS Notificação, " & _
                " TNT_DT_EMISSAO as Emissao, " & _
                " TNT_INSCRICAO as IM, tci_nome AS Contribuinte, " & _
                FuncaoReal("TNT_VALOR_NOTIFICACAO") & " as Valor" & _
            " FROM TAB_NOTIFICACAO, TAB_PAGAMENTO_NOTIFICACAO, TAB_CONTRIBUINTE " & _
            " WHERE TNT_COD_NOTIFICACAO=TPN_TNO_COD_NOTIFICACAO AND TNT_INSCRICAO=TCI_IM"
    If Trim$(Imposto) <> "" Then Sql = Sql & " and TPN_TIP_COD_IMPOSTO='" & Imposto & "'"
    If Trim$(Im) <> "" Then Sql = Sql & " and TNT_INSCRICAO='" & Im & "'"
    If Trim$(EmissaoInicio) <> "" Then Sql = Sql & " and TNT_DT_EMISSAO>=" & Bdados.FormataValorCampo(EmissaoInicio)
    If Trim$(EmissaoFim) <> "" Then Sql = Sql & " and TNT_DT_EMISSAO<=" & Bdados.FormataValorCampo(EmissaoInicio)
    If Trim$(NumInicio) <> "" Then Sql = Sql & " and TNT_COD_NOTIFICACAO>=" & NumInicio & "'"
    If Trim$(NumFim) <> "" Then Sql = Sql & " and TNT_COD_NOTIFICACAO<=" & NumFim & "'"
    Sql = Sql & " ORDER BY TNT_VENCIMENTO"
    
    exibirCNDs = Sql
End Function

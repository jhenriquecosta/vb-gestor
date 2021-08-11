VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransfImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public TipoTransferencia As String
Public AliquotaProprio As Double
Public AliquotaFinanciado As Double
Public DestinoLote As String
Public OcupacaoLote As String
Private Lote As String
Private ContribuinteCedente As String
Public Sub Gravar(NumeroDocumento As String, DataEmissao As String, IMAdquirente As String, CPFAdquirente As String, _
                    NomeAdquirente As String, EnderecoAdquirente As String, EnderecoImovel As String, NomeCedente As String, _
                    EnderecoCedente As String, Optional ValorAvista As Double, Optional ValorFinanciado As Double, _
                    Optional ValorLoteInformado As Double, Optional ValorLoteCalculdado As Double, _
                    Optional ValorItbi As Double, Optional Especie As String, Optional Observacao As String, _
                    Optional AliquotaProprio As Double, Optional AliquotaFinanciado As Double, Optional Processo As String)
    Dim Valores As String
    Dim Campos As String
    
    Campos = "TTI_TIM_IC,TTI_TOC_COD_OBRIGACAO, TTI_DATA_GERACAO,TTI_CEDENTE_NOME,TTI_CEDENTE_ENDERECO, " & _
        "TTI_TCI_IM_ADQUIRENTE,TTI_VALOR_AVISTA,TTI_VALOR_FINANCIADO,TTI_ESPECIE,TTI_VALOR_LOTE_CALCULADO," & _
        "TTI_VALOR_LOTE_INFORMADO,TTI_VALOR_OBRIGACAO,TTI_OBSERVACAO,TTI_ALIQUOTA_PROPRIO,TTI_ALIQUOTA_FINANCIADO," & _
        "TTI_TCI_CPF_ADQUIRENTE,TTI_TCI_NOME_ADQUIRENTE,TTI_TCI_ENDERECO_ADQUIRENTE,TTI_TCI_ENDERECO_IMOVEL,TTI_PROCESSO"
    Valores = Bdados.PreparaValor(Bdados.Converte(Lote, tctexto), NumeroDocumento, _
        Bdados.Converte(DataEmissao, TCDataHora), NomeCedente, EnderecoCedente, IMAdquirente, _
        Bdados.Converte(ValorAvista, TCDuplo), Bdados.Converte(ValorFinanciado, TCDuplo), Especie, _
        ValorLoteCalculdado, ValorLoteInformado, ValorItbi, Observacao, AliquotaProprio, AliquotaFinanciado, _
        CPFAdquirente, NomeAdquirente, EnderecoAdquirente, EnderecoImovel, Processo)
    Bdados.GravaDados "TAB_TRANSFERENCIA_IMOVEL", Valores, Campos, "TTI_TOC_COD_OBRIGACAO =" & NumeroDocumento
End Sub

Public Sub BuscaDetalhesTransferencia(Imovel As String, Optional ValorVenalInformado As Double)
    Dim Sql  As String
    Dim rs As VSRecordset
    Dim Ocupacao As Integer
    Dim Destino As Integer
    Dim Iptu As New VSIptu
    Dim ValorLote As Double
    
    Lote = Imovel
    AliquotaFinanciado = 0
    AliquotaProprio = 0
    OcupacaoLote = ""
    DestinoLote = ""
    Imovel = IIf(Trim(Imovel) = "", "FUTURO", Imovel)
    Sql = "Select tdi_tgc_cod_grupo,tco_descricao_componente,tdi_tco_cod_componente,tdi_tim_ic_unidade from tab_detalhe_imovel,tab_componente_avancado where "
    Sql = Sql & " tdi_tim_ic = '" & Imovel & "' and tdi_tco_cod_componente=tco_cod_componente and tdi_tgc_cod_grupo = tco_grupo  and  (tdi_tgc_cod_grupo = 1 or tdi_tgc_cod_grupo = 11) "
    Sql = Sql & "  order by  tdi_tim_ic_unidade,tdi_tgc_cod_grupo "
    If Bdados.AbreTabela(Sql, rs) Then
        Do
            If rs!tdi_tgc_cod_grupo = 1 Then
                Ocupacao = IIf(rs!tdi_tco_cod_componente <> 5, 1, rs!tdi_tco_cod_componente)
                OcupacaoLote = rs!tco_descricao_componente
            ElseIf rs!tdi_tgc_cod_grupo = 11 Then
                Destino = rs!tdi_tco_cod_componente
                DestinoLote = rs!tco_descricao_componente
                Exit Do
            End If
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    
    
    Sql = "Select TIM_TCI_IM,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC from tab_imovel where tim_ic  ='" & Imovel & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        ValorLote = IIf(ValorVenalInformado > 0, ValorVenalInformado, Nvl("" & rs!TIM_VALOR_TERRENO, 0) + Nvl("" & rs!TIM_VALOR_EDIFIC, 0))
        Contribuinte = rs!tim_tci_im
    Else
        ValorLote = ValorVenalInformado
    End If
    Ocupacao = IIf(Ocupacao = 0, 99, Ocupacao)
    Destino = IIf(Destino = 0, 99, Destino)
    Sql = "Select TTT_COD_ALIQUOTA,TTT_ALIQUOTA_PROPRIO,TTT_ALIQUOTA_FINANCIADO,TTT_LIMITE_INFERIOR,TTT_LIMITE_SUPERIOR from Tab_Tipo_Transferencia_Imovel " & _
    " where TTT_TCO_COD_OCUPACAO =  " & Ocupacao
    Sql = Sql & " and TTT_TCO_COD_DESTINACAO = " & Destino
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            If Nvl("" & rs!TTT_LIMITE_SUPERIOR, 0) = 0 Then 'ULTIMA FAIXA
                If ValorLote >= Nvl("" & rs!TTT_LIMITE_INFERIOR, 0) Then
                    AliquotaProprio = Nvl("" & rs!TTT_ALIQUOTA_PROPRIO, 0)
                    AliquotaFinanciado = Nvl("" & rs!TTT_ALIQUOTA_FINANCIADO, 0)
                    TipoTransferencia = Nvl("" & rs!TTT_COD_ALIQUOTA, 0)
                    Exit Do
                End If
            ElseIf ValorLote >= rs!TTT_LIMITE_INFERIOR And ValorLote <= rs!TTT_LIMITE_SUPERIOR Then
                AliquotaProprio = Nvl("" & rs!TTT_ALIQUOTA_PROPRIO, 0)
                AliquotaFinanciado = Nvl("" & rs!TTT_ALIQUOTA_FINANCIADO, 0)
                TipoTransferencia = Nvl("" & rs!TTT_COD_ALIQUOTA, 0)
                Exit Do
            End If
            rs.MoveNext
        Loop While Not rs.EOF
    End If
End Sub

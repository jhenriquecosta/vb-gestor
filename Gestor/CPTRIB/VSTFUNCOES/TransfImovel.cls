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
Dim bcpCV As String
Public Sub Gravar(NumeroDocumento As String, DataEmissao As String, IMAdquirente As String, CPFAdquirente As String, _
                    NomeAdquirente As String, EnderecoAdquirente As String, EnderecoImovel As String, NomeCedente As String, _
                    EnderecoCedente As String, Optional ValorAvista As Double, Optional ValorFinanciado As Double, _
                    Optional ValorLoteInformado As Double, Optional ValorLoteCalculdado As Double, _
                    Optional ValorItbi As Double, Optional Especie As String, Optional Observacao As String, _
                    Optional AliquotaProprio As Double, Optional AliquotaFinanciado As Double, Optional Processo As String)
    Dim Valores As String
    Dim Campos As String
    
    If IsNull(Lote) Then
        bcpCV = Format(Now, "YYYY") & Right(IMAdquirente, 2) & NumeroDocumento
    Else
        bcpCV = Format(Now, "YYYY") & Right(IMAdquirente, 2) & Lote
    End If
    Campos = "TTI_TIM_IC,TTI_TOC_COD_OBRIGACAO, TTI_DATA_GERACAO,TTI_CEDENTE_NOME,TTI_CEDENTE_ENDERECO, " & _
        "TTI_TCI_IM_ADQUIRENTE,TTI_VALOR_AVISTA,TTI_VALOR_FINANCIADO,TTI_ESPECIE,TTI_VALOR_LOTE_CALCULADO," & _
        "TTI_VALOR_LOTE_INFORMADO,TTI_VALOR_OBRIGACAO,TTI_OBSERVACAO,TTI_ALIQUOTA_PROPRIO,TTI_ALIQUOTA_FINANCIADO," & _
        "TTI_TCI_CPF_ADQUIRENTE,TTI_TCI_NOME_ADQUIRENTE,TTI_TCI_ENDERECO_ADQUIRENTE,TTI_TCI_ENDERECO_IMOVEL,TTI_PROCESSO,TTI_COD_VERIFICADOR"
    Valores = Bdados.PreparaValor(Bdados.Converte(Lote, tctexto), NumeroDocumento, _
        Bdados.Converte(DataEmissao, TCDataHora), NomeCedente, EnderecoCedente, IMAdquirente, _
        Bdados.Converte(ValorAvista, TCDuplo), Bdados.Converte(ValorFinanciado, TCDuplo), Especie, _
        ValorLoteCalculdado, ValorLoteInformado, ValorItbi, Observacao, AliquotaProprio, AliquotaFinanciado, _
        CPFAdquirente, NomeAdquirente, EnderecoAdquirente, EnderecoImovel, Processo, bcpCV)
    Bdados.GravaDados "TAB_TRANSFERENCIA_IMOVEL", Valores, Campos, "TTI_TOC_COD_OBRIGACAO =" & NumeroDocumento
End Sub
Public Sub GravarBCP(NumeroDocumento As String, DataEmissao As String, IMAdquirente As String, CPFAdquirente As String, _
                    NomeAdquirente As String, EnderecoAdquirente As String, EnderecoImovel As String, NomeCedente As String, _
                    EnderecoCedente As String, Optional ValorAvista As Double, Optional ValorFinanciado As Double, _
                    Optional ValorLoteInformado As Double, Optional ValorLoteCalculdado As Double, _
                    Optional ValorItbi As Double, Optional Especie As String, Optional Observacao As String, _
                    Optional AliquotaProprio As Double, Optional AliquotaFinanciado As Double, Optional Processo As String, Optional AliquotaFGTS As Double, Optional ValorFGTS As Double)
    Dim Valores As String
    Dim Campos As String
    
    If IsNull(Lote) Then
        bcpCV = Format(Now, "YYYY") & Right(IMAdquirente, 2) & NumeroDocumento
    Else
        bcpCV = Format(Now, "YYYY") & Right(IMAdquirente, 2) & Lote
    End If
    Campos = "TTI_TIM_IC,TTI_TOC_COD_OBRIGACAO, TTI_DATA_GERACAO,TTI_CEDENTE_NOME,TTI_CEDENTE_ENDERECO, " & _
        "TTI_TCI_IM_ADQUIRENTE,TTI_VALOR_AVISTA,TTI_VALOR_FINANCIADO,TTI_ESPECIE,TTI_VALOR_LOTE_CALCULADO," & _
        "TTI_VALOR_LOTE_INFORMADO,TTI_VALOR_OBRIGACAO,TTI_OBSERVACAO,TTI_ALIQUOTA_PROPRIO,TTI_ALIQUOTA_FINANCIADO," & _
        "TTI_TCI_CPF_ADQUIRENTE,TTI_TCI_NOME_ADQUIRENTE,TTI_TCI_ENDERECO_ADQUIRENTE,TTI_TCI_ENDERECO_IMOVEL,TTI_PROCESSO,TTI_COD_VERIFICADOR,TTI_ALIQUOTA_FGTS,TTI_VALOR_FGTS"
    Valores = Bdados.PreparaValor(Bdados.Converte(Lote, tctexto), NumeroDocumento, _
        Bdados.Converte(DataEmissao, TCDataHora), NomeCedente, EnderecoCedente, IMAdquirente, _
        Bdados.Converte(ValorAvista, TCDuplo), Bdados.Converte(ValorFinanciado, TCDuplo), Especie, _
        ValorLoteCalculdado, ValorLoteInformado, ValorItbi, Observacao, AliquotaProprio, AliquotaFinanciado, _
        CPFAdquirente, NomeAdquirente, EnderecoAdquirente, EnderecoImovel, Processo, bcpCV, AliquotaFGTS, ValorFGTS)
    Bdados.GravaDados "TAB_TRANSFERENCIA_IMOVEL", Valores, Campos, "TTI_TOC_COD_OBRIGACAO =" & NumeroDocumento
End Sub
Public Sub BuscaDetalhesTransferencia(Imovel As String, Optional ValorVenalInformado As Double)
    Dim Sql  As String
    Dim Rs As VSRecordset
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
    If Bdados.AbreTabela(Sql, Rs) Then
        Do
            If Rs!tdi_tgc_cod_grupo = 1 Then
                Ocupacao = IIf(Rs!tdi_tco_cod_componente <> 5, 1, Rs!tdi_tco_cod_componente)
                OcupacaoLote = Rs!tco_descricao_componente
            ElseIf Rs!tdi_tgc_cod_grupo = 11 Then
                Destino = Rs!tdi_tco_cod_componente
                DestinoLote = Rs!tco_descricao_componente
                Exit Do
            End If
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
    
    
    Sql = "Select TIM_TCI_IM,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC from tab_imovel where tim_ic  ='" & Imovel & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        ValorLote = IIf(ValorVenalInformado > 0, ValorVenalInformado, Nvl("" & Rs!TIM_VALOR_TERRENO, 0) + Nvl("" & Rs!TIM_VALOR_EDIFIC, 0))
        Contribuinte = Rs!tim_tci_im
    Else
        ValorLote = ValorVenalInformado
    End If
    Ocupacao = IIf(Ocupacao = 0, 99, Ocupacao)
    Destino = IIf(Destino = 0, 99, Destino)
    Sql = "Select TTT_COD_ALIQUOTA,TTT_ALIQUOTA_PROPRIO,TTT_ALIQUOTA_FINANCIADO,TTT_LIMITE_INFERIOR,TTT_LIMITE_SUPERIOR from Tab_Tipo_Transferencia_Imovel " & _
    " where TTT_TCO_COD_OCUPACAO =  " & Ocupacao
    Sql = Sql & " and TTT_TCO_COD_DESTINACAO = " & Destino
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            If Nvl("" & Rs!TTT_LIMITE_SUPERIOR, 0) = 0 Then 'ULTIMA FAIXA
                If ValorLote >= Nvl("" & Rs!TTT_LIMITE_INFERIOR, 0) Then
                    AliquotaProprio = Nvl("" & Rs!TTT_ALIQUOTA_PROPRIO, 0)
                    AliquotaFinanciado = Nvl("" & Rs!TTT_ALIQUOTA_FINANCIADO, 0)
                    TipoTransferencia = Nvl("" & Rs!TTT_COD_ALIQUOTA, 0)
                    Exit Do
                End If
            ElseIf ValorLote >= Rs!TTT_LIMITE_INFERIOR And ValorLote <= Rs!TTT_LIMITE_SUPERIOR Then
                AliquotaProprio = Nvl("" & Rs!TTT_ALIQUOTA_PROPRIO, 0)
                AliquotaFinanciado = Nvl("" & Rs!TTT_ALIQUOTA_FINANCIADO, 0)
                TipoTransferencia = Nvl("" & Rs!TTT_COD_ALIQUOTA, 0)
                Exit Do
            End If
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
End Sub
Public Sub ImprimeITBI(Documento As String)
    Dim Rpt As VSRelatorio
    Set Rpt = New VSRelatorio
    Dim Rs As VSRecordset
    Dim titulo As String, obs As String
    titulo = "CERTIDÃO DE TRANSMISSÃO DE BENS E IMÓVEIS"
    obs = ""
    If Documento = "" Then
       Util.Informa "Informe Um Documento para impressão!"
       Exit Sub
    End If
       
    If Bdados.AbreTabela("SELECT * FROM VIS_TRANSFERENCIA_IMOVEL_RURAL  WHERE TTI_TOC_COD_OBRIGACAO = " & Documento, Rs) Then
        If IsNull(Rs("TTI_TIM_IC")) Or Len(Rs("TTI_TIM_IC")) = 0 Then
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path + "\TCertidaoITBI.rpt") Then Exit Sub
                
                .Formulas "TITULO", titulo & " - RURAL"
                .Formulas "CODIGO_VERIFICADOR", Rs("TTI_COD_VERIFICADOR")
                .Formulas "EMISSAO", Format(Rs("TTI_DATA_GERACAO"), "DD/MM/YYYY")
                .Formulas "CODIGO_VERIFICADOR", Rs("TTI_COD_VERIFICADOR")
                .Formulas "IM", " " & Rs("TTI_TCI_IM_ADQUIRENTE")
                .Formulas "CPF/CNPJ", " " & Rs("TCI_CGC_CPF")
                .Formulas "RAZAOSOCIAL", " " & Rs("TCI_NOME")
                If IsNull(Rs("TCI_FANTASIA")) Then
                    .Formulas "NOMEFANTASIA", " " & Rs("TCI_NOME")
                Else
                    .Formulas "NOMEFANTASIA", " " & Rs("TCI_FANTASIA")
                End If
                .Formulas "ENDERECO", " " & Rs("TTI_TCI_ENDERECO_ADQUIRENTE")
                .Formulas "BAIRRO", " " & Rs("TCI_BAIRRO")
                .Formulas "CIDADE", " " & Rs("TCI_CIDADE")
                
                .Formulas "RAZAOTRANSMITENTE", " " & Rs("TTI_CEDENTE_NOME")
                .Formulas "ENDERECOTRANSMITENTE", IIf(IsNull(Rs("TTI_CEDENTE_ENDERECO")), "", Rs("TTI_CEDENTE_ENDERECO"))
                
                '.Formulas "IC_IMOVEL", Rs("TTI_TIM_IC")
                .Formulas "ENDERECO_IMOVEL", Rs("TTI_TCI_ENDERECO_IMOVEL")
                .Formulas "BAIRRO_IMOVEL", "ZONA RURAL"
                If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                    .Formulas "CIDADE_IMOVEL", "CODÓ"
                End If
                If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then
                    .Formulas "CIDADE_IMOVEL", "GRAJÁU"
                End If
                
                .Formulas "PROCESSO", Rs("TTI_PROCESSO")
                
                
                .Formulas "AREA_TERRENO", Format(0, "#,##0.00")
                .Formulas "AREA_IMOVEL", Format(0, "#,##0.00")
                
                
                .Formulas "V_IMOVEL", Format(Rs("TTI_VALOR_AVISTA") + Rs("TTI_VALOR_FINANCIADO") + Rs("TTI_VALOR_FGTS"), "R$ #,##0.00")
                .Formulas "DOCUMENTO", Documento
                .Formulas "V_VISTA", Format(Rs("TTI_VALOR_AVISTA"), "R$ #,##0.00")
                .Formulas "A_VISTA", Format(Rs("TTI_ALIQUOTA_PROPRIO"), "#,##0.00") & "%"
                
                .Formulas "A_FGTS", Format(Rs("TTI_ALIQUOTA_FGTS"), "#,##0.00") & "%"
                .Formulas "V_FGTS", Format(Rs("TTI_VALOR_FGTS"), "R$ #,##0.00")
                
                .Formulas "V_FINANCIADO", Format(Rs("TTI_VALOR_FINANCIADO"), "R$ #,##0.00")
                .Formulas "A_FINANCIADO", Format(Rs("TTI_ALIQUOTA_FINANCIADO"), "#,##0.00") & "%"
                .Formulas "ITBI_TOTAL", Format(Rs("TTI_VALOR_OBRIGACAO"), "R$ #,##0.00")
                If Not IsNull(Rs("TTI_OBSERVACAO")) Then
                    obs = Rs("TTI_OBSERVACAO")
                Else
                    obs = ""
                End If
                .Formulas "OBSERVACAO", obs
                .Visualizar
            End With
        Else
            If Bdados.AbreTabela("SELECT * FROM VIS_TRANSFERENCIA_IMOVEL WHERE TTI_TOC_COD_OBRIGACAO = " & Documento, Rs) Then
                With Rpt
                    If Not .DefinirArquivo(Bdados, App.Path + "\TCertidaoITBI.rpt") Then Exit Sub
                    .Formulas "TITULO", titulo
                    obs = "Verificou-se também não haver existência de débitos referente ao Imposto Territorial Urbano - IPTU."
                    .Formulas "OBS", obs
                    .Formulas "CODIGO_VERIFICADOR", Rs("TTI_COD_VERIFICADOR")
                    .Formulas "EMISSAO", Format(Rs("TTI_DATA_GERACAO"), "DD/MM/YYYY")
                    .Formulas "CODIGO_VERIFICADOR", Rs("TTI_COD_VERIFICADOR")
                    .Formulas "IM", Rs("TTI_TCI_IM_ADQUIRENTE")
                    .Formulas "CPF/CNPJ", Rs("TCI_CGC_CPF")
                    .Formulas "RAZAOSOCIAL", Rs("TCI_NOME")
                    If Not IsNull(Rs("TCI_FANTASIA")) Then
                        .Formulas "NOMEFANTASIA", Rs("TCI_FANTASIA")
                    Else
                        .Formulas "NOMEFANTASIA", Rs("TCI_NOME")
                    End If
                    .Formulas "ENDERECO", Rs("TTI_TCI_ENDERECO_ADQUIRENTE")
                    .Formulas "BAIRRO", Rs("TCI_BAIRRO")
                    If IsNull(Rs("TCI_CIDADE")) Then
                        .Formulas "CIDADE", ""
                    Else
                        .Formulas "CIDADE", Rs("TCI_CIDADE")
                    End If
                    .Formulas "RAZAOTRANSMITENTE", Rs("TTI_CEDENTE_NOME")
                    If Not IsNull(Rs("TTI_CEDENTE_ENDERECO")) Then
                        .Formulas "ENDERECOTRANSMITENTE", Rs("TTI_CEDENTE_ENDERECO")
                    Else
                        .Formulas "ENDERECOTRANSMITENTE", ""
                    End If
                    .Formulas "IC_IMOVEL", Rs("TTI_TIM_IC")
                    .Formulas "ENDERECO_IMOVEL", Rs("TTI_TCI_ENDERECO_IMOVEL")
                    .Formulas "BAIRRO_IMOVEL", Rs("TBA_NOME")
                    '.Formulas "CIDADE_IMOVEL", Rs("TMU_NOME")
                    If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                        .Formulas "CIDADE_IMOVEL", "CODÓ"
                    End If
                    If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then
                        .Formulas "CIDADE_IMOVEL", "GRAJÁU"
                    End If
                
                    .Formulas "PROCESSO", Rs("TTI_PROCESSO")
                    
                    
                    Dim ri As VSRecordset, rt As VSRecordset
                    Dim di As Double, dt As Double
                    di = 0
                    dt = 0
                    If Bdados.AbreTabela("SELECT TDI_VALOR_ITEM FROM TAB_DETALHE_IMOVEL WHERE TDI_TGC_COD_GRUPO=108 AND TDI_TIM_IC='" & Rs("TTI_TIM_IC") & "'", rt) Then
                        dt = rt(0)
                    End If
                    If Bdados.AbreTabela("SELECT TDI_VALOR_ITEM FROM TAB_DETALHE_IMOVEL WHERE TDI_TGC_COD_GRUPO=113 AND TDI_TIM_IC='" & Rs("TTI_TIM_IC") & "'", ri) Then
                        di = ri(0)
                    End If
                    .Formulas "AREA_TERRENO", Format(dt, "#,##0.00")
                    .Formulas "AREA_IMOVEL", Format(di, "#,##0.00")
                    
                    .Formulas "A_FGTS", Format(Rs("TTI_ALIQUOTA_FGTS"), "#,##0.00") & "%"
                    .Formulas "V_FGTS", Format(Rs("TTI_VALOR_FGTS"), "R$ #,##0.00")
                    
                    .Formulas "V_IMOVEL", Format(Rs("TTI_VALOR_AVISTA") + Rs("TTI_VALOR_FINANCIADO") + Rs("TTI_VALOR_FGTS"), "R$ #,##0.00")
                    .Formulas "DOCUMENTO", Documento
                    .Formulas "V_VISTA", Format(Rs("TTI_VALOR_AVISTA"), "R$ #,##0.00")
                    .Formulas "A_VISTA", Format(Rs("TTI_ALIQUOTA_PROPRIO"), "#,##0.00") & "%"
                    .Formulas "V_FINANCIADO", Format(Rs("TTI_VALOR_FINANCIADO"), "R$ #,##0.00")
                    .Formulas "A_FINANCIADO", Format(Rs("TTI_ALIQUOTA_FINANCIADO"), "#,##0.00") & "%"
                    .Formulas "ITBI_TOTAL", Format(Rs("TTI_VALOR_OBRIGACAO"), "R$ #,##0.00")
                    If Not IsNull(Rs("TTI_OBSERVACAO")) Then
                        obs = Rs("TTI_OBSERVACAO")
                    Else
                        obs = ""
                    End If
                    .Formulas "OBSERVACAO", obs
                    .Visualizar
                End With
            End If
        End If
    End If
    
End Sub


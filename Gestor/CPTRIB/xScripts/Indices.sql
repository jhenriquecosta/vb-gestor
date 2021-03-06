--Deleta os indices atuais

DROP INDEX TAB_ATIVIDADE_ESTIMADA.CAE
GO
DROP INDEX TAB_ATIVIDADE_ESTIMADA.lim_inf
GO
DROP INDEX TAB_COMPONENTE_AVANCADO.IDX_COD_COMPONENTE
GO
DROP INDEX TAB_COMPONENTE_AVANCADO.IDX_TCO_GRUPO
GO
DROP INDEX TAB_COMPONENTE_AVANCADO.IDX_COMPONENTE_GRUPO
GO
DROP INDEX TAB_DETALHE_IMOVEL.IC
GO
DROP INDEX TAB_DETALHE_IMOVEL.UND
GO
DROP INDEX TAB_DETALHE_IMOVEL.GRUPO
GO
DROP INDEX TAB_DETALHE_IMOVEL.COMP
GO
DROP INDEX TAB_DETALHE_LOGRADOURO.detlogradouro
GO
DROP INDEX TAB_DETALHE_OBRIGACAO.IMPOSTO
GO
DROP INDEX TAB_GERACAO_TRIBUTO_PARCELADO.DOC
GO
DROP INDEX TAB_IMOVEL.IC
GO
DROP INDEX TAB_IMOVEL.IDX_TLG_COD_LOGRADOURO
GO
DROP INDEX TAB_IMOVEL.IM
GO
DROP INDEX TAB_OBRIGACAO_CONTRIBUINTE.INSC
GO
DROP INDEX TAB_OBRIGACAO_CONTRIBUINTE.PERIODO
GO
DROP INDEX TAB_OBRIGACAO_CONTRIBUINTE.IMPOSTO
GO
DROP INDEX TAB_PAGAMENTO_EXTRATO.DOC
GO
DROP INDEX TAB_PAGAMENTO_EXTRATO.EXTRATO
GO
DROP INDEX TAB_TRECHO.LOGR
GO
DROP INDEX TAB_TRECHO.TRECHO
GO
DROP INDEX Tab_Componente_Logradouro.componente
GO
DROP INDEX Tab_Conta_Contribuinte.PK_Tab_Conta_Contribuinte
GO
DROP INDEX Tab_Conta_Contribuinte.IM
GO
DROP INDEX Tab_Conta_Contribuinte.IMPOSTO
GO
DROP INDEX Tab_Conta_Contribuinte.PERIODO
GO
DROP INDEX Tab_Conta_Contribuinte.conta
GO
DROP INDEX Tab_Conta_Transacao.CONTA
GO
DROP INDEX Tab_Conta_Transacao.DATA
GO
DROP INDEX Tab_Conta_Transacao.TRANSAC
GO
DROP INDEX Tab_Contribuinte.tcs_tae_cae
GO
DROP INDEX Tab_Contribuinte.tcs_tnj_cod_natureza
GO
DROP INDEX Tab_Darm_Recebido.tdr_im
GO
DROP INDEX Tab_Darm_Recebido.tdr_tgt_cod_pagamento
GO
DROP INDEX Tab_Darm_Recebido.tdr_tim_ic
GO
DROP INDEX Tab_Darm_Recebido.inscricao
GO
DROP INDEX Tab_Geracao_Tributo.im
GO
DROP INDEX Tab_Geracao_Tributo.periodo
GO
DROP INDEX Tab_Geracao_Tributo.ic
GO
DROP INDEX Tab_Geracao_Tributo.imp
GO
DROP INDEX Tab_Geracao_Tributo_Parcela.doc
GO
DROP INDEX Tab_Geracao_Tributo_Parcela.im
GO
DROP INDEX Tab_Geracao_Tributo_Parcela.imposto
GO
DROP INDEX Tab_Geracao_Tributo_Parcela.ic
GO
DROP INDEX Tab_Geracao_Tributo_Parcela.parcela
GO
DROP INDEX Tab_Isento.IM
GO
DROP INDEX Tab_Isento.IMPOSTO
GO
DROP INDEX Tab_Isento.PERIODO
GO
DROP INDEX Tab_Isento.TIPO
GO
DROP INDEX Tab_Isento.IC
GO
DROP INDEX Tab_Parcelamento.TPA_NUM_PARCELAMENTO
GO
DROP INDEX Tab_Parcelamento.TPA_TIM_IC
GO


--Recria todos os indices
 CREATE  INDEX [cae] ON [dbo].[TAB_ATIVIDADE_ESTIMADA]([TAT_TAE_CAE]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [lim_inf] ON [dbo].[TAB_ATIVIDADE_ESTIMADA]([TAT_LIMITE_INFERIOR]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IDX_COD_COMPONENTE] ON [dbo].[TAB_COMPONENTE_AVANCADO]([tco_cod_componente]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IDX_TCO_GRUPO] ON [dbo].[TAB_COMPONENTE_AVANCADO]([tco_grupo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IDX_COMPONENTE_GRUPO] ON [dbo].[TAB_COMPONENTE_AVANCADO]([tco_cod_componente], [tco_grupo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IC] ON [dbo].[TAB_DETALHE_IMOVEL]([tdi_tim_ic]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [UND] ON [dbo].[TAB_DETALHE_IMOVEL]([tdi_tim_ic_unidade]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [GRUPO] ON [dbo].[TAB_DETALHE_IMOVEL]([tdi_tgc_cod_grupo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [COMP] ON [dbo].[TAB_DETALHE_IMOVEL]([tdi_tco_cod_componente]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [detlogradouro] ON [dbo].[TAB_DETALHE_LOGRADOURO]([tdl_tlg_cod_logradouro], [tdl_tcl_cod_componente], [tdl_tgl_cod_grupo], [tdl_num_trecho]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IMPOSTO] ON [dbo].[TAB_DETALHE_OBRIGACAO]([TDO_TIP_COD_IMPOSTO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [DOC] ON [dbo].[TAB_GERACAO_TRIBUTO_PARCELADO]([tgt_cod_pagamento]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IC] ON [dbo].[TAB_IMOVEL]([tim_ic]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IDX_TLG_COD_LOGRADOURO] ON [dbo].[TAB_IMOVEL]([tim_tlg_cod_logradouro]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IM] ON [dbo].[TAB_IMOVEL]([tim_tci_im]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [INSC] ON [dbo].[TAB_OBRIGACAO_CONTRIBUINTE]([TOC_INSCRICAO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [PERIODO] ON [dbo].[TAB_OBRIGACAO_CONTRIBUINTE]([TOC_PERIODO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IMPOSTO] ON [dbo].[TAB_OBRIGACAO_CONTRIBUINTE]([TOC_TIP_COD_IMPOSTO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [DOC] ON [dbo].[TAB_PAGAMENTO_EXTRATO]([TPE_TGT_COD_PAGAMENTO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [EXTRATO] ON [dbo].[TAB_PAGAMENTO_EXTRATO]([TPE_COD_PAGAMENTO_EXTRATO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [LOGR] ON [dbo].[TAB_TRECHO]([TTC_TLG_COD_LOGRADOURO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [TRECHO] ON [dbo].[TAB_TRECHO]([TTC_COD_TRECHO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [componente] ON [dbo].[Tab_Componente_Logradouro]([tcl_cod_componente], [tcl_grupo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [PK_Tab_Conta_Contribuinte] ON [dbo].[Tab_Conta_Contribuinte]([tcc_codigo_conta]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IM] ON [dbo].[Tab_Conta_Contribuinte]([tcc_im]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IMPOSTO] ON [dbo].[Tab_Conta_Contribuinte]([tcc_tip_cod_imposto]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [PERIODO] ON [dbo].[Tab_Conta_Contribuinte]([tcc_periodo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [conta] ON [dbo].[Tab_Conta_Contribuinte]([tcc_codigo_conta]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [CONTA] ON [dbo].[Tab_Conta_Transacao]([tct_tcc_codigo_conta]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [DATA] ON [dbo].[Tab_Conta_Transacao]([tct_data_transacao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [TRANSAC] ON [dbo].[Tab_Conta_Transacao]([tct_tipo_transacao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [tcs_tae_cae] ON [dbo].[Tab_Contribuinte]([tci_tae_cae]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [tcs_tnj_cod_natureza] ON [dbo].[Tab_Contribuinte]([tci_tnj_cod_natureza]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [tdr_im] ON [dbo].[Tab_Darm_Recebido]([tdr_im]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [tdr_tgt_cod_pagamento] ON [dbo].[Tab_Darm_Recebido]([tdr_tgt_cod_pagamento]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [tdr_tim_ic] ON [dbo].[Tab_Darm_Recebido]([tdr_tim_ic]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [inscricao] ON [dbo].[Tab_Darm_Recebido]([TDR_INSCRICAO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [im] ON [dbo].[Tab_Geracao_Tributo]([tgt_im]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [periodo] ON [dbo].[Tab_Geracao_Tributo]([tgt_periodo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [ic] ON [dbo].[Tab_Geracao_Tributo]([tgt_tim_ic]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [imp] ON [dbo].[Tab_Geracao_Tributo]([tgt_tip_cod_imposto]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [doc] ON [dbo].[Tab_Geracao_Tributo_Parcela]([tgt_cod_pagamento]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [im] ON [dbo].[Tab_Geracao_Tributo_Parcela]([tgt_im]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [imposto] ON [dbo].[Tab_Geracao_Tributo_Parcela]([tgt_tip_cod_imposto]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [ic] ON [dbo].[Tab_Geracao_Tributo_Parcela]([tgt_tim_ic]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [parcela] ON [dbo].[Tab_Geracao_Tributo_Parcela]([tgt_parcela]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IM] ON [dbo].[Tab_Isento]([TIS_TCI_IM]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IMPOSTO] ON [dbo].[Tab_Isento]([TIS_TIP_COD_IMPOSTO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [PERIODO] ON [dbo].[Tab_Isento]([TIS_PERIODO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [TIPO] ON [dbo].[Tab_Isento]([TIS_TIPO_ISENSAO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IC] ON [dbo].[Tab_Isento]([TIS_TIM_IC]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  UNIQUE  INDEX [TPA_NUM_PARCELAMENTO] ON [dbo].[Tab_Parcelamento]([TPA_NUM_PARCELAMENTO]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [TPA_TIM_IC] ON [dbo].[Tab_Parcelamento]([TPA_TIM_IC]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO


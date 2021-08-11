package com.bcptecnologias.migracao.gestortributario.negocio.gestor;

public enum NumeroCorrelativo {
	CREDENCIAMENTO_GRAFIA(22,0),
	AIDF(23,0),
	DIVIDA_ATIVA(66,0),
	EXTRATO_LANCAMENTO(99,0),
	INSCRICAO_CONTRIBUINTE(11,0),
	INSCRICAO_IMOVEL(31,0),
	IPTU_11(11,0),
	IPTU_75(75,0),
	ISS_RETIDO_PELA_PREFEITURA(82,0),
	NOTA_FISCAL_AVULSA(65,0),
	NUMERO_AIDF(23,0),
	PAGAMENTO_ALVARA(91,0),
	PAGAMENTO_IPTU(71,0),
	PAGAMENTO_ISS(81,0),
	PARCELAMENTO(55,0),
	DEBITOS_PARCELADOS(64,0),
	ITBI(93,0),
	LOTE_E_MODIFICACAO_CADASTRO(0,0),
	
	NAO_IDENTIFICADO(37,0),
	NAO_IDENTIFICADO1(97,0),
	NAO_IDENTIFICADO2(95,0),
	NAO_IDENTIFICADO3(96,0),
	NAO_IDENTIFICADO4(67,0),
	NAO_IDENTIFICADO5(64,0),
	NAO_IDENTIFICADO6(53,0),
	NAO_IDENTIFICADO7(78,0),
	
	;
	NumeroCorrelativo(Integer operacao, Integer valorPadrao){
		this.operacao=operacao;
		this.valorPadrao=valorPadrao;
	}
	private Integer operacao;
	private Integer valorPadrao;
	public Integer getOperacao() {
		return operacao;
	}
	public Integer getValorPadrao() {
		return valorPadrao;
	}
	
	
}

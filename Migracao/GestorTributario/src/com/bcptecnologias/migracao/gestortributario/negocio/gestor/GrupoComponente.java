package com.bcptecnologias.migracao.gestortributario.negocio.gestor;

public enum GrupoComponente implements Componente {
	TERRENO_OCUPACAO_LOTE(26),
	TERRENO_PATRIMONIO(27),
	TERRENO_LIMITE(33),
	EDIFICACAO_TIPOLOGIA(39),
	TERRENO_SITUACAO(43),
	TERRENO_TOPOGRAFIA(44),
	TERRENO_PEDOLOGIA(45),
	EDIFICACAO_DESTINACAO(77),
	EDIFICACAO_ESTRUTURA(78),
	EDIFICACAO_PADRAO(81),
	EDIFICACAO_CONSERVACAO(92),
	TESTADA_PRINCIPAL(100),
	TESTADA_2(101),
	EDIFICACAO_TIPO_COBRANCA(102),
	TESTADA_3(103),
	TESTADA_4(105),
	AREA_LOTE(108),
	PAVIMENTOS(110),
	ANO_CONTRUCAO(111),
	AREA_EDIFICADA_UNIDADE(112),
	AREA_EDIFICADA_TOTAL(113),
	PROFUNDIDADE(115),
	UNIDADE_FISCAL(400),
	VALOR_FIXO(1000),
	;
	private Integer valor;
	GrupoComponente(Integer valor){
		this.valor=valor;
	}
	public Integer getValor(){
		return valor;
	}
}

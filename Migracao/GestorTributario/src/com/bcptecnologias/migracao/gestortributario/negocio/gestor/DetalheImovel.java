package com.bcptecnologias.migracao.gestortributario.negocio.gestor;

import com.bcptecnologias.migracao.gestortributario.negocio.Imovel;

public class DetalheImovel {
	public enum TerrenoOcupacaoLote implements Componente{
		NAO_CONSTRUIDO(1),
		RUINAS(2),
		EM_DEMOLICAO(3),
		CONSTRUCAO_PARALISADA(4),
		CONSTRUCAO_EM_ANDAMENTO(5),
		AGROPECUARIA(6),
		CONSTRUIDO(7),
		;
		private Integer valor;
		TerrenoOcupacaoLote(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum TerrenoPatrimonio implements Componente{
		PARTICULAR(1),
		RELIGIOSO(2),
		PUBLICO_FEDERAL(3),
		PUBLICO_ESTADUAL(4),
		PUBLICO_MUNICIPAL(5),
		ASSOCIACAO(6),
		;
		private Integer valor;
		TerrenoPatrimonio(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum TerrenoLimite implements Componente{
		SEM_MURO_CERCA(1),
		CERCADO(2),
		MURADO(3),
		CERCA_MURO(4),
		;
		private Integer valor;
		TerrenoLimite(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	public enum TerrenoSituacao implements Componente{
		MEIO_DE_QUADRA_COM_UMA_FRENTE(1),
		MEIO_DE_QUADRA_COM_DUAS_FRENTES(2),
		ESQUINA_COM_MAIS_DE_UMA_FRENTE(3),
		GLEBA(4),
		FUNDOS(5),
		ENCRAVADOS(6),
		;
		private Integer valor;
		TerrenoSituacao(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	public enum TerrenoTopografia implements Componente{
		PLANA(1),
		IRREGULAR(2),
		ACLIVE_SUAVE(3),
		ACLIVE_ACENTUADO(4),
		DECLIVE_SUAVE(5),
		DECLIVE_ACENTUADO(6),
		;
		private Integer valor;
		TerrenoTopografia(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	public enum TerrenoPedologia implements Componente{
		FIRME(1),
		INUNDAVEL(2),
		ALAGADO(3),
		ROCHOSO(4),
		ARENOSO(5),
		COMBINACAO_DAS_DEMAIS(6),
		;
		private Integer valor;
		TerrenoPedologia(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	 
	public enum EdificacaoTipologia implements Componente{
		CASA(1),
		CONSTRUCAO_PRECARIA(2),
		BARRACO(3),
		APARTAMENTO(4),
		LOJA_SALA_CONJUNTO(5),
		GALPAO(6),
		DEPOSITO(7),
		TALHEIRO(8),
		OUTROS(9),
		;
		private Integer valor;
		EdificacaoTipologia(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	public enum EdificacaoDestinacao implements Componente{
		TERRENO_SEM_USO(1),
		RESIDENCIAL(2),
		INDUSTRIAL(3),
		COMERCIAL(4),
		SERVICOS(5),
		AGROPECUARIA(6),
		TEMPLO(7),
		FUNDACAO(8),
		OUTROS(9),
		;
		private Integer valor;
		EdificacaoDestinacao(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	public enum EdificacaoEstrutura implements Componente{
		ALVENARIA(1),
		MADEIRA(2),
		METALICO(3),
		CONCRETO(4),
		ALVENARIA_MADEIRA(5),
		ALVENARIA_CONCRETO(6),
		;
		private Integer valor;
		EdificacaoEstrutura(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum EdificacaoPadrao implements Componente{
		ALTO(1),
		MEDIO(2),
		BAIXO(3),
		;
		private Integer valor;
		EdificacaoPadrao(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum EdificacaoConservacao implements Componente{
		OTIMO(1),
		BOA(2),
		REGULAR(3),
		PRECARIO(4),
		;
		private Integer valor;
		EdificacaoConservacao(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum EdificacaoTipoCobranca implements Componente{
		NORMAL(1),
		ISENTO(2),
		IMUNE(3),
		;
		private Integer valor;
		EdificacaoTipoCobranca(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	private Imovel imovel;
	public Imovel getImovel() {
		return imovel;
	}
	public Componente getGrupoComponente() {
		return grupoComponente;
	}
	public Componente getComponente() {
		return componente;
	}
	public Double getValorItem() {
		return valorItem;
	}
	private Componente grupoComponente;
	private Componente componente;
	private Double valorItem;
	
	/*
	public DetalheImovel(Imovel imovel, Componente grupoComponente, Componente componente,Double valorItem){
		this.imovel=imovel;
		this.grupoComponente=grupoComponente;
		this.componente=componente;
		this.valorItem=valorItem;
	}
	*/
	public DetalheImovel(Imovel imovel, Componente grupoComponente, Componente componente){
		this.imovel=imovel;
		this.grupoComponente=grupoComponente;
		this.componente=componente;
		this.valorItem=0.0d;
	}
	public DetalheImovel(Imovel imovel, Componente grupoComponente, Double valorItem){
		this.imovel=imovel;
		this.grupoComponente=grupoComponente;
		this.componente=grupoComponente;
		this.valorItem=valorItem;
	}
	
}

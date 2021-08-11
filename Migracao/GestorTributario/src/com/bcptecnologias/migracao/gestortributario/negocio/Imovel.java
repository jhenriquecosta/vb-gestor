package com.bcptecnologias.migracao.gestortributario.negocio;

import java.util.Date;

public class Imovel {
	public enum TipoImovel{
		NAO_INFORMADO(0),
		CEDIDO(1),
		PROPRIO(2),
		ALUGADO(3),
		;
		private Integer valor;
		TipoImovel(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum SituacaoCadastral{
		HABILITADO(1),
		;
		private Integer valor;
		SituacaoCadastral(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum SituacaoLote{
		NAO_IDENTIFICADO(0),
		PREDIAL(1),
		TERRITORIAL(2),
		CONDOMINIO(3),
		;
		private Integer valor;
		SituacaoLote(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}	
	private String inscricao;
	private Double unidade;
	private String inscricaoAuxiliar;
	private Contribuinte contrbuinte;
	private String inscricaoAnterior;
	private Logradouro logradouro;
	private String numero;
	private String complemento;
	private String cep;
	private Integer anoAquisicao;
	private Double valor;
	private TipoImovel tipoImovel;
	private String ocupante;
	private String cpfCnpjOcupante;
	private String inscricaoCondominio;
	private Bairro bairro;
	private Integer codMensagem;
	private Integer zona;
	private Double valorTerreno;
	private Double valorEdificao;
	private String loteamento;
	private String secao;
	private String quadra;
	private String lote;
	private SituacaoCadastral situacaoCadastral;
	private boolean aforado;
	private String observacao;
	private String registroAforamento;
	private Date dataRegistroAforamento;
	private String numeroAforamento;
	private String fichaAforamento;
	private String livroAforamenteo;
	private String folhaAforamento;
	private Date dataAforamento;
	private String motivoAlteracao;
	private Double valorTerrenoMercado;
	private Double valorEdificadoMercado;
	private SituacaoLote situacaoLote;
	private String subUnidade;
	private String usuario;
	private Date dataCadastro;
	private String tipoBoletim;
	private Integer codEdificio;
	private String bloco;
	private String apto;
	private String salaLoja;
	private String sequenciaTrecho;
	private String codigoTrecho;

	public String getInscricao() {
		return inscricao;
	}
	public void setInscricao(String inscricao) {
		this.inscricao = inscricao;
	}
	public Double getUnidade() {
		return unidade;
	}
	public void setUnidade(Double unidade) {
		this.unidade = unidade;
	}
	public String getInscricaoAuxiliar() {
		return inscricaoAuxiliar;
	}
	public void setInscricaoAuxiliar(String inscricaoAuxiliar) {
		this.inscricaoAuxiliar = inscricaoAuxiliar;
	}
	public Contribuinte getContrbuinte() {
		return contrbuinte;
	}
	public void setContrbuinte(Contribuinte contrbuinte) {
		this.contrbuinte = contrbuinte;
	}
	public String getInscricaoAnterior() {
		return inscricaoAnterior;
	}
	public void setInscricaoAnterior(String inscricaoAnterior) {
		this.inscricaoAnterior = inscricaoAnterior;
	}
	public Logradouro getLogradouro() {
		return logradouro;
	}
	public void setLogradouro(Logradouro logradouro) {
		this.logradouro = logradouro;
	}
	public String getNumero() {
		return numero;
	}
	public void setNumero(String numero) {
		this.numero = numero;
	}
	public String getComplemento() {
		return complemento;
	}
	public void setComplemento(String complemento) {
		this.complemento = complemento;
	}
	public String getCep() {
		return cep;
	}
	public void setCep(String cep) {
		this.cep = cep;
	}
	public Integer getAnoAquisicao() {
		return anoAquisicao;
	}
	public void setAnoAquisicao(Integer anoAquisicao) {
		this.anoAquisicao = anoAquisicao;
	}
	public Double getValor() {
		return valor;
	}
	public void setValor(Double valor) {
		this.valor = valor;
	}
	public TipoImovel getTipoImovel() {
		return tipoImovel;
	}
	public void setTipoImovel(TipoImovel tipoImovel) {
		this.tipoImovel = tipoImovel;
	}
	public String getOcupante() {
		return ocupante;
	}
	public void setOcupante(String ocupante) {
		this.ocupante = ocupante;
	}
	public String getCpfCnpjOcupante() {
		return cpfCnpjOcupante;
	}
	public void setCpfCnpjOcupante(String cpfCnpjOcupante) {
		this.cpfCnpjOcupante = cpfCnpjOcupante;
	}
	public String getInscricaoCondominio() {
		return inscricaoCondominio;
	}
	public void setInscricaoCondominio(String inscricaoCondominio) {
		this.inscricaoCondominio = inscricaoCondominio;
	}
	public Bairro getBairro() {
		return bairro;
	}
	public void setBairro(Bairro bairro) {
		this.bairro = bairro;
	}
	public Integer getCodMensagem() {
		return codMensagem;
	}
	public void setCodMensagem(Integer codMensagem) {
		this.codMensagem = codMensagem;
	}
	public Integer getZona() {
		return zona;
	}
	public void setZona(Integer zona) {
		this.zona = zona;
	}
	public Double getValorTerreno() {
		return valorTerreno;
	}
	public void setValorTerreno(Double valorTerreno) {
		this.valorTerreno = valorTerreno;
	}
	public Double getValorEdificao() {
		return valorEdificao;
	}
	public void setValorEdificao(Double valorEdificao) {
		this.valorEdificao = valorEdificao;
	}
	public String getLoteamento() {
		return loteamento;
	}
	public void setLoteamento(String loteamento) {
		this.loteamento = loteamento;
	}
	public String getSecao() {
		return secao;
	}
	public void setSecao(String secao) {
		this.secao = secao;
	}
	public String getQuadra() {
		return quadra;
	}
	public void setQuadra(String quadra) {
		this.quadra = quadra;
	}
	public String getLote() {
		return lote;
	}
	public void setLote(String lote) {
		this.lote = lote;
	}
	public SituacaoCadastral getSituacaoCadastral() {
		return situacaoCadastral;
	}
	public void setSituacaoCadastral(SituacaoCadastral situacaoCadastral) {
		this.situacaoCadastral = situacaoCadastral;
	}
	public boolean getAforado() {
		return aforado;
	}
	public void setAforado(boolean aforado) {
		this.aforado = aforado;
	}
	public String getObservacao() {
		return observacao;
	}
	public void setObservacao(String observacao) {
		this.observacao = observacao;
	}
	public String getRegistroAforamento() {
		return registroAforamento;
	}
	public void setRegistroAforamento(String registroAforamento) {
		this.registroAforamento = registroAforamento;
	}
	public Date getDataRegistroAforamento() {
		return dataRegistroAforamento;
	}
	public void setDataRegistroAforamento(Date dataRegistroAforamento) {
		this.dataRegistroAforamento = dataRegistroAforamento;
	}
	public String getNumeroAforamento() {
		return numeroAforamento;
	}
	public void setNumeroAforamento(String numeroAforamento) {
		this.numeroAforamento = numeroAforamento;
	}
	public String getFichaAforamento() {
		return fichaAforamento;
	}
	public void setFichaAforamento(String fichaAforamento) {
		this.fichaAforamento = fichaAforamento;
	}
	public String getLivroAforamenteo() {
		return livroAforamenteo;
	}
	public void setLivroAforamenteo(String livroAforamenteo) {
		this.livroAforamenteo = livroAforamenteo;
	}
	public String getFolhaAforamento() {
		return folhaAforamento;
	}
	public void setFolhaAforamento(String folhaAforamento) {
		this.folhaAforamento = folhaAforamento;
	}
	public Date getDataAforamento() {
		return dataAforamento;
	}
	public void setDataAforamento(Date dataAforamento) {
		this.dataAforamento = dataAforamento;
	}
	public String getMotivoAlteracao() {
		return motivoAlteracao;
	}
	public void setMotivoAlteracao(String motivoAlteracao) {
		this.motivoAlteracao = motivoAlteracao;
	}
	public Double getValorTerrenoMercado() {
		return valorTerrenoMercado;
	}
	public void setValorTerrenoMercado(Double valorTerrenoMercado) {
		this.valorTerrenoMercado = valorTerrenoMercado;
	}
	public Double getValorEdificadoMercado() {
		return valorEdificadoMercado;
	}
	public void setValorEdificadoMercado(Double valorEdificadoMercado) {
		this.valorEdificadoMercado = valorEdificadoMercado;
	}
	public SituacaoLote getSituacaoLote() {
		return situacaoLote;
	}
	public void setSituacaoLote(SituacaoLote situacaoLote) {
		this.situacaoLote = situacaoLote;
	}
	public String getSubUnidade() {
		return subUnidade;
	}
	public void setSubUnidade(String subUnidade) {
		this.subUnidade = subUnidade;
	}
	public String getUsuario() {
		return usuario;
	}
	public void setUsuario(String usuario) {
		this.usuario = usuario;
	}
	public Date getDataCadastro() {
		return dataCadastro;
	}
	public void setDataCadastro(Date dataCadastro) {
		this.dataCadastro = dataCadastro;
	}
	public String getTipoBoletim() {
		return tipoBoletim;
	}
	public void setTipoBoletim(String tipoBoletim) {
		this.tipoBoletim = tipoBoletim;
	}
	public Integer getCodEdificio() {
		return codEdificio;
	}
	public void setCodEdificio(Integer codEdificio) {
		this.codEdificio = codEdificio;
	}
	public String getBloco() {
		return bloco;
	}
	public void setBloco(String bloco) {
		this.bloco = bloco;
	}
	public String getApto() {
		return apto;
	}
	public void setApto(String apto) {
		this.apto = apto;
	}
	public String getSalaLoja() {
		return salaLoja;
	}
	public void setSalaLoja(String salaLoja) {
		this.salaLoja = salaLoja;
	}
	public String getSequenciaTrecho() {
		return sequenciaTrecho;
	}
	public void setSequenciaTrecho(String sequenciaTrecho) {
		this.sequenciaTrecho = sequenciaTrecho;
	}
	public String getCodigoTrecho() {
		return codigoTrecho;
	}
	public void setCodigoTrecho(String codigoTrecho) {
		this.codigoTrecho = codigoTrecho;
	}
	
	
	
	
	
	
	
}

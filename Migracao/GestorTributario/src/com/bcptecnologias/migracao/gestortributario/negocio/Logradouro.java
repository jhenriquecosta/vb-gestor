package com.bcptecnologias.migracao.gestortributario.negocio;


public class Logradouro {
	public enum TipoLogradouro{
		NAO_IDENTIFICADO(0),
		FAZENDA(1),
		AVENIDA(2),
		PRACA(3),
		PRAIA(4),
		POVOADO(5),
		ALAMEDA(6),
		CONDOMINIO(7),
		CAMIMHO(8),
		ILHA(9),
		LADEIRA(10),
		LOTEAMENTO(11),
		LUGAREJO(12),
		PARQUE(13),
		PONTE(14),
		VIADUTO(15),
		RODOVIA(16),
		RUA(17),
		TRAVESSA(18),
		ESTRADA(19),
		BECO(20),
		VILA(21),
		CARRASCO(22),
		BR(23),
		CALCADA(24),
		CONJUNTO(25),
		;
		private Integer valor;
		TipoLogradouro(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	
	
	private Integer codigo;
	private String nome;
	private Bairro bairro;
	private String cep;
	//QUANDO FOR O PROPRIO NOME
	private String nomeTipoLogradouro;
	private TipoLogradouro tipoLogradouro;
	
	public String getNomeTipoLogradouro() {
		return nomeTipoLogradouro;
	}
	public TipoLogradouro getTipoLogradouro() {
		return tipoLogradouro;
	}
	public void setTipoLogradouro(TipoLogradouro tipoLogradouro) {
		this.tipoLogradouro = tipoLogradouro;
	}
	public void setNomeTipoLogradouro(String nomeTipoLogradouro) {
		this.nomeTipoLogradouro = nomeTipoLogradouro;
	}
	public Bairro getBairro() {
		return bairro;
	}
	public void setBairro(Bairro bairro) {
		this.bairro = bairro;
	}
	public String getCep() {
		return cep;
	}
	public void setCep(String cep) {
		this.cep = cep;
	}
	public String getNome() {
		return nome;
	}
	public void setNome(String nome) {
		this.nome = nome;
	}
	public Integer getCodigo() {
		return codigo;
	}
	public void setCodigo(Integer codigo) {
		this.codigo = codigo;
	}
}

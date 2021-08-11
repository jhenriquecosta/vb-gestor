package com.bcptecnologias.migracao.gestortributario.negocio.gestor;

public enum TabelaCpSeg {
	TAB_ACESSO_USUARIO("TAB_ACESSO_USUARIO"),
	TAB_ACESSO_USUARIO_PRIVILEGIO("TAB_ACESSO_USUARIO_PRIVILEGIO"),
	TAB_USUARIO("TAB_USUARIO"),
	TAB_BAIRRO("TAB_BAIRRO"),
	TAB_LOGRADOURO("TAB_LOGRADOURO"),
	;
	private String nome;
	TabelaCpSeg(String nome){
		this.nome=nome;
	}
	public String getNome(){
		return nome;
	}
}

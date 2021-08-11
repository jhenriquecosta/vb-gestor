package com.bcptecnologias.migracao.gestortributario.negocio;

import java.util.Date;


public class Contribuinte {
	public enum SituacaoCadastral{
		NAO_IDENTIFICADO(0),
		HABILITADO(1),
		SUSPENSO_POR_OFICIO(2),
		BAIXADO(3),
		INATIVO(4),
		CANCELADO(5),
		;
		private Integer valor;
		SituacaoCadastral(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum GrupoAtividade{
		NAO_IDENTIFICADO(0),
		AGROPECUARIA(1),
		INDUSTRIA(2),
		COMERCIO(3),
		PRESTACAO_SERVICO(4),
		CULTO_RELIGIOSO(5),
		ORGAO_PUBLICO(6),
		ASSOCIACAO_CLASSE(7),
		;
		private Integer valor;
		GrupoAtividade(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum NaturezaJuridica{
		NAO_IDENTIFICADO(0),
		PESSOA_FISICA(1),
		PESSOA_JURIDICA(2),
		SOCIEDADE_CIVIL(3),
		RELIGIOSO(4),
		ORGAO_PUBLICO(5),
		;
		private Integer valor;
		NaturezaJuridica(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum AtividadePoder{
		NAO_IDENTIFICADO(0),
		PARTICULAR(1),
		MUNICIPAL(2),
		ESTADUAL(3),
		FEDERAL(4),
		;
		private Integer valor;
		AtividadePoder(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum Estabelecido{
		NAO_IDENTIFICADO(0),
		SIM(1),
		NAO(2),
		OUTROS(3),
		;
		private Integer valor;
		Estabelecido(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum TipoContribuinte{
		NAO_IDENTIFICADO(0),
		FISICO(1),
		JURIDICO(2),
		AMBOS(3),
		;
		private Integer valor;
		TipoContribuinte(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum TipoRecolhimento{
		NAO_IDENTIFICADO(0),
		SEM_RECOLHIMENTO(1),
		RECOLHIMENTO_MENSAL(2),
		RECOLHIMENTO_FIXO_ANUAL(3),
		;
		private Integer valor;
		TipoRecolhimento(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum PorteEmpresa{
		NAO_IDENTIFICADO(0),
		MICRO(1),
		PEQUENO(2),
		MEDIO(3),
		GRANDE(4),
		;
		private Integer valor;
		PorteEmpresa(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum SituacaoAlvara{
		NAO_IDENTIFICADO(0),
		DEFINITIVO(1),
		PROVISORIO(2);
		private Integer valor;
		SituacaoAlvara(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum TipoEmpresa{
		NAO_IDENTIFICADO(0),
		MATRIZ(1),
		FILIAL(2);
		private Integer valor;
		TipoEmpresa(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum TipoImovel{
		NAO_IDENTIFICADO(0),
		CEDIDO(1),
		ALUGADO(2),
		PROPRIO(3),
		;
		private Integer valor;
		TipoImovel(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	public enum Escolaridade{
		NAO_IDENTIFICADO(0),
		NAO_ALFABETIZADO(1),
		FUNDAMENTAL_INCOMPLETO(2),
		FUNDAMENTAL_COMPLETO(3),
		MEDIO_INCOMPLETO(4),
		MEDIO_COMPLETO(5),
		SUPERIOR_INCOMPLETO(6),
		SUPERIOR_COMPLETO(7),
		ESPECIALIZADO(8),
		ALFABETIZADO(9),
		
		;
		private Integer valor;
		Escolaridade(Integer valor){
			this.valor=valor;
		}
		public Integer getValor(){
			return valor;
		}
	}
	private String codigo;
	private String codigoOrigem;
	private String inscricao;
	private String inscricaoAnterior;
	private String inscricaoAuxiliar;
	private String cpfCnpj;
	private String nome;
	private String nomeFantasia;
	private String logradouro;
	private String nomeLogradouro;
	private String numero;
	private String nomeBairro;
	private String complemento;
	private String cep;
	private Cidade cidade;
	private String uf;
	private Date dataCadastro;
	private Date dataAlteracao;
	private Date inicioAtividade;
	private Date fimAtividade;
	private Date dataEncerramento;
	
	private Date inicioPrestacaoServico;
	private String email;
	private String identidadeRg;
	private SituacaoCadastral situacao; 
	private String usuario;
	private GrupoAtividade grupoAtividade;
	private NaturezaJuridica naturezaJuridica; 
	private String cae;
	private Integer grupoCae;
	private AtividadePoder atividadePoder;
	private Estabelecido estabelecido;
	private TipoContribuinte tipoContribuinte;
	private TipoRecolhimento tipoRecolhimento;
	private PorteEmpresa porteEmpresa;
	private boolean isento;
	private Long fatorAlvara;
	private String foneFax;
	private String cnh;
	private String categoriaCnh;
	private Logradouro fkLogradouro;
	private Bairro bairro;
	private SituacaoAlvara situacaoAlvara;
	private TipoEmpresa tipoEmpresa;
	private String telefone;
	private TipoImovel tipoImovel;
	private Integer numeroEmpregados;
	private Integer tipoCadastro;
	private boolean imovelProprio;
	private Integer registro;
	private Escolaridade escolaridade;
	private String ruc;
	private String conselho;
	private String caeSecundario;
	private String caeTerciario;
	private String protocolo;
	private Integer ramo;
	private String autorizacao;
	private String pontoRecepcao; 
	private boolean optanteSimplesNacional;
	
	public String getCodigo() {
		return codigo;
	}
	public void setCodigo(String codigo) {
		this.codigo = codigo;
	}
	public boolean isOptanteSimplesNacional() {
		return optanteSimplesNacional;
	}
	public void setOptanteSimplesNacional(boolean optanteSimplesNacional) {
		this.optanteSimplesNacional = optanteSimplesNacional;
	}
	public String getAutorizacao() {
		return autorizacao;
	}
	public void setAutorizacao(String autorizacao) {
		this.autorizacao = autorizacao;
	}
	public String getPontoRecepcao() {
		return pontoRecepcao;
	}
	public void setPontoRecepcao(String pontoRecepcao) {
		this.pontoRecepcao = pontoRecepcao;
	}
	public Integer getRamo() {
		return ramo;
	}
	public void setRamo(Integer ramo) {
		this.ramo = ramo;
	}
	public String getProtocolo() {
		return protocolo;
	}
	public void setProtocolo(String protocolo) {
		this.protocolo = protocolo;
	}
	public String getCaeSecundario() {
		return caeSecundario;
	}
	public Date getDataEncerramento() {
		return dataEncerramento;
	}
	public void setDataEncerramento(Date dataEncerramento) {
		this.dataEncerramento = dataEncerramento;
	}
	public String getIdentidadeRg() {
		return identidadeRg;
	}
	public void setIdentidadeRg(String identidadeRg) {
		this.identidadeRg = identidadeRg;
	}
	public String getCodigoOrigem() {
		return codigoOrigem;
	}
	public void setCodigoOrigem(String codigoOrigem) {
		this.codigoOrigem = codigoOrigem;
	}
	public void setCaeSecundario(String caeSecundario) {
		this.caeSecundario = caeSecundario;
	}
	public String getCaeTerciario() {
		return caeTerciario;
	}
	public void setCaeTerciario(String caeTerciario) {
		this.caeTerciario = caeTerciario;
	}
	public String getConselho() {
		return conselho;
	}
	public void setConselho(String conselho) {
		this.conselho = conselho;
	}
	public String getRuc() {
		return ruc;
	}
	public void setRuc(String ruc) {
		this.ruc = ruc;
	}
	public Escolaridade getEscolaridade() {
		return escolaridade;
	}
	public void setEscolaridade(Escolaridade escolaridade) {
		this.escolaridade = escolaridade;
	}
	public Integer getRegistro() {
		return registro;
	}
	public void setRegistro(Integer registro) {
		this.registro = registro;
	}
	public boolean isImovelProprio() {
		return imovelProprio;
	}
	public void setImovelProprio(boolean imovelProprio) {
		this.imovelProprio = imovelProprio;
	}
	public Date getFimAtividade() {
		return fimAtividade;
	}
	public void setFimAtividade(Date fimAtividade) {
		this.fimAtividade = fimAtividade;
	}
	public Date getDataAlteracao() {
		return dataAlteracao;
	}
	public void setDataAlteracao(Date dataAlteracao) {
		this.dataAlteracao = dataAlteracao;
	}
	public GrupoAtividade getGrupoAtividade() {
		return grupoAtividade;
	}
	public void setGrupoAtividade(GrupoAtividade grupoAtividade) {
		this.grupoAtividade = grupoAtividade;
	}
	public void setNaturezaJuridica(NaturezaJuridica naturezaJuridica) {
		this.naturezaJuridica = naturezaJuridica;
	}
	public String getNomeBairro() {
		return nomeBairro;
	}
	public void setNomeBairro(String nomeBairro) {
		this.nomeBairro = nomeBairro;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public Date getInicioPrestacaoServico() {
		return inicioPrestacaoServico;
	}
	public void setInicioPrestacaoServico(Date inicioPrestacaoServico) {
		this.inicioPrestacaoServico = inicioPrestacaoServico;
	}
	public Integer getTipoCadastro() {
		return tipoCadastro;
	}
	public void setTipoCadastro(Integer tipoCadastro) {
		this.tipoCadastro = tipoCadastro;
	}
	
	public Integer getNumeroEmpregados() {
		return numeroEmpregados;
	}
	public void setNumeroEmpregados(Integer numeroEmpregados) {
		this.numeroEmpregados = numeroEmpregados;
	}
	public TipoImovel getTipoImovel() {
		return tipoImovel;
	}
	public void setTipoImovel(TipoImovel tipoImovel) {
		this.tipoImovel = tipoImovel;
	}
	public Date getInicioAtividade() {
		return inicioAtividade;
	}
	public void setInicioAtividade(Date inicioAtividade) {
		this.inicioAtividade = inicioAtividade;
	}
	public String getNumero() {
		return numero;
	}
	public void setNumero(String numero) {
		this.numero = numero;
	}
	
	public NaturezaJuridica getNaturezaJuridica() {
		return naturezaJuridica;
	}
	public String getCae() {
		return cae;
	}
	public void setCae(String cae) {
		this.cae = cae;
	}
	public Integer getGrupoCae() {
		return grupoCae;
	}
	public void setGrupoCae(Integer grupoCae) {
		this.grupoCae = grupoCae;
	}
	public AtividadePoder getAtividadePoder() {
		return atividadePoder;
	}
	public void setAtividadePoder(AtividadePoder atividadePoder) {
		this.atividadePoder = atividadePoder;
	}
	public Estabelecido getEstabelecido() {
		return estabelecido;
	}
	public void setEstabelecido(Estabelecido estabelecido) {
		this.estabelecido = estabelecido;
	}
	public TipoContribuinte getTipoContribuinte() {
		return tipoContribuinte;
	}
	public void setTipoContribuinte(TipoContribuinte tipoContribuinte) {
		this.tipoContribuinte = tipoContribuinte;
	}
	public TipoRecolhimento getTipoRecolhimento() {
		return tipoRecolhimento;
	}
	public void setTipoRecolhimento(TipoRecolhimento tipoRecolhimento) {
		this.tipoRecolhimento = tipoRecolhimento;
	}
	public PorteEmpresa getPorteEmpresa() {
		return porteEmpresa;
	}
	public void setPorteEmpresa(PorteEmpresa porteEmpresa) {
		this.porteEmpresa = porteEmpresa;
	}
	public boolean isIsento() {
		return isento;
	}
	public void setIsento(boolean isento) {
		this.isento = isento;
	}
	public Long getFatorAlvara() {
		return fatorAlvara;
	}
	public void setFatorAlvara(Long fatorAlvara) {
		this.fatorAlvara = fatorAlvara;
	}
	public String getFoneFax() {
		return foneFax;
	}
	public void setFoneFax(String foneFax) {
		this.foneFax = foneFax;
	}
	public String getCnh() {
		return cnh;
	}
	public void setCnh(String cnh) {
		this.cnh = cnh;
	}
	public String getCategoriaCnh() {
		return categoriaCnh;
	}
	public void setCategoriaCnh(String categoriaCnh) {
		this.categoriaCnh = categoriaCnh;
	}
	public Logradouro getFkLogradouro() {
		return fkLogradouro;
	}
	public void setFkLogradouro(Logradouro fkLogradouro) {
		this.fkLogradouro = fkLogradouro;
	}
	public Bairro getBairro() {
		return bairro;
	}
	public void setBairro(Bairro bairro) {
		this.bairro = bairro;
	}
	public SituacaoAlvara getSituacaoAlvara() {
		return situacaoAlvara;
	}
	public void setSituacaoAlvara(SituacaoAlvara situacaoAlvara) {
		this.situacaoAlvara = situacaoAlvara;
	}
	public TipoEmpresa getTipoEmpresa() {
		return tipoEmpresa;
	}
	public void setTipoEmpresa(TipoEmpresa tipoEmpresa) {
		this.tipoEmpresa = tipoEmpresa;
	}
	public String getTelefone() {
		return telefone;
	}
	public void setTelefone(String telefone) {
		this.telefone = telefone;
	}
	public String getInscricao() {
		return inscricao;
	}
	public void setInscricao(String inscricao) {
		this.inscricao = inscricao;
	}
	public String getCpfCnpj() {
		return cpfCnpj;
	}
	public void setCpfCnpj(String cpfCnpj) {
		this.cpfCnpj = cpfCnpj;
	}
	public String getNome() {
		return nome;
	}
	public void setNome(String nome) {
		this.nome = nome;
	}
	public String getNomeFantasia() {
		return nomeFantasia;
	}
	public void setNomeFantasia(String nomeFantasia) {
		this.nomeFantasia = nomeFantasia;
	}
	public String getLogradouro() {
		return logradouro;
	}
	public void setLogradouro(String logradouro) {
		this.logradouro = logradouro;
	}
	public String getNomeLogradouro() {
		return nomeLogradouro;
	}
	public void setNomeLogradouro(String nomeLogradouro) {
		this.nomeLogradouro = nomeLogradouro;
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
	public Cidade getCidade() {
		return cidade;
	}
	public void setCidade(Cidade cidade) {
		this.cidade = cidade;
	}
	public String getUf() {
		return uf;
	}
	public void setUf(String uf) {
		this.uf = uf;
	}
	public Date getDataCadastro() {
		return dataCadastro;
	}
	public void setDataCadastro(Date dataCadastro) {
		this.dataCadastro = dataCadastro;
	}
	public SituacaoCadastral getSituacao() {
		return situacao;
	}
	public void setSituacao(SituacaoCadastral situacao) {
		this.situacao = situacao;
	}
	public String getUsuario() {
		return usuario;
	}
	public void setUsuario(String usuario) {
		this.usuario = usuario;
	}
	public String getInscricaoAnterior() {
		return inscricaoAnterior;
	}
	public void setInscricaoAnterior(String inscricaoAnterior) {
		this.inscricaoAnterior = inscricaoAnterior;
	}
	public String getInscricaoAuxiliar() {
		return inscricaoAuxiliar;
	}
	public void setInscricaoAuxiliar(String inscricaoAuxiliar) {
		this.inscricaoAuxiliar = inscricaoAuxiliar;
	}
}


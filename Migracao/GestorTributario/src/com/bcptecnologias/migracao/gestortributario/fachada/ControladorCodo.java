package com.bcptecnologias.migracao.gestortributario.fachada;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import com.bcptecnologias.migracao.gestortributario.negocio.Bairro;
import com.bcptecnologias.migracao.gestortributario.negocio.Cidade;
import com.bcptecnologias.migracao.gestortributario.negocio.Contribuinte;
import com.bcptecnologias.migracao.gestortributario.negocio.Contribuinte.SituacaoCadastral;
import com.bcptecnologias.migracao.gestortributario.negocio.Imovel;
import com.bcptecnologias.migracao.gestortributario.negocio.Logradouro;
import com.bcptecnologias.migracao.gestortributario.negocio.Uf;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.DetalheImovel;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.GrupoComponente;
import com.bcptecnologias.migracao.gestortributario.util.Utilidade;


public class ControladorCodo {
	private static Connection conexaoCodo;
	public static void setConexaoCodo(Connection conexao){
		conexaoCodo=conexao;
	}
	//-BAIRROS
	public static List<Bairro> listarBairros(Cidade cidade)throws SQLException{
		List<Bairro>bairros=new ArrayList();
		String sql="SELECT * FROM BAIRROS ORDER BY BAIRRO_CODIGO";
		ResultSet rsB=conexaoCodo.createStatement().executeQuery(sql);
		while(rsB.next()){
			Bairro bairro=new Bairro();
			bairro.setCodigo(rsB.getInt("BAIRRO_CODIGO"));
			bairro.setNome(rsB.getString("BAIRRO_NOME"));
			bairro.setCidade(cidade);
			bairros.add(bairro);
		}
		return bairros;
	}
	//-LOGRADOUROS
	public static List<Logradouro> listarLogradouros()throws SQLException{
		List<Logradouro>logradouros=new ArrayList();
		String sql="SELECT * FROM LOGRADOUROS ORDER BY LOG_CODIGO";
		ResultSet rs=conexaoCodo.createStatement().executeQuery(sql);
		while(rs.next()){
			Logradouro logradouro=new Logradouro();
			logradouro.setCodigo(rs.getInt("LOG_CODIGO"));
			logradouro.setNome(rs.getString("LOG_NOME"));
			logradouro.setNomeTipoLogradouro(rs.getString("LOG_TIPO"));
			logradouro.setBairro(ControladorCpSeg.buscarBairro(validarBairroRepetido(rs.getInt("LOG_COD_BAI"))));
			logradouro.setCep("65400-000");
			logradouro.setTipoLogradouro(tipoLogradouro(logradouro.getNomeTipoLogradouro()));
			logradouros.add(logradouro);
		}
		return logradouros;
	}
	//-CONTRIBUINTES
	public static List<Contribuinte> listarContribuintes(Cidade cidade) throws SQLException{
		//String sql="select * from contribuintes where CONTRIB_NOME LIKE 'MARIA DA CONCEIÇÃO%' order by CONTRIB_NOME ;";
		String sql="select * from contribuintes order by CONTRIB_NOME ;";
		Statement st=conexaoCodo.createStatement();
		ResultSet rs=st.executeQuery(sql);
		//PRIMEIRA TABELA CONTRIBUINTES
		List<Contribuinte>contribuintes=listarDaTabelaContribuinte(rs,cidade);
		contribuintes=listarDaTabelaCadastroMobiliario(contribuintes);
		return contribuintes;
		/* OS CAMPOS DO GESTOR
		c.setInscricao("");
		c.setInscricaoAnterior("");
		c.setInscricaoAuxiliar("");
		c.setCpfCnpj("");
		c.setNome("");
		c.setNomeFantasia("");
		c.setLogradouro("");
		c.setNomeLogradouro("");
		c.setNumero("");
		c.setComplemento("");
		c.setNomeBairro("");
		c.setCep("");
		c.setCidade(cidade);
		c.setDataCadastro(null);
		c.setDataAlteracao(null);
		c.setSituacao(null);
		c.setUsuario("");
		c.setInicioAtividade(null);
		c.setGrupoAtividade(null);
		c.setNaturezaJuridica(null);
		c.setCae(0L);
		c.setAtividadePoder(null);
		c.setEstabelecido(null);
		c.setGrupoCae(null);
		c.setTipoContribuinte(null);
		c.setTipoRecolhimento(null);
		c.setRuc("");
		c.setConselho("");
		c.setImovelProprio(false);
		c.setRegistro(0);
		c.setNumeroEmpregados(0);
		c.setPorteEmpresa(null);
		c.setCaeSecundario(0L);
		c.setCaeTerciario(0L);
		c.setEscolaridade(null);
		c.setProtocolo("");
		c.setIsento(false);
		c.setFatorAlvara(0L);
		c.setRamo(0);
		c.setFoneFax("");
		c.setCnh("");
		c.setCategoriaCnh("");
		c.setAutorizacao("");
		c.setPontoRecepcao("");
		c.setFkLogradouro(null);
		c.setBairro(null);
		c.setTipoContribuinte(null);
		c.setTipoCadastro(0);
		c.setInicioPrestacaoServico(null);
		c.setSituacaoAlvara(null);
		c.setPorteEmpresa(null);
		c.setDataEncerramento(null);
		c.setTelefone("");
		c.setEmail("");
		c.setIdentidadeRg("");
		c.setOptanteSimplesNacional(false);
		*/
	}
	private static List<Contribuinte> listarDaTabelaContribuinte(ResultSet rs,Cidade cidade)throws SQLException{
		List<Contribuinte>contribuintes=new ArrayList();
		//Integer inscricao=1;
		while(rs.next()){
			Contribuinte c=new Contribuinte();
				c.setCae("0");
				c.setCaeSecundario("0");
				c.setCaeTerciario("0");
				c.setInscricao(Utilidade.gerarInscricaoPeloCodigoExistente(rs.getString("CONTRIB_CODIGO")));
				
				//inscricao++;
				
				c.setTipoImovel(Contribuinte.TipoImovel.NAO_IDENTIFICADO);
				c.setPorteEmpresa(Contribuinte.PorteEmpresa.PEQUENO);
				c.setEscolaridade(Contribuinte.Escolaridade.ALFABETIZADO);
				c.setSituacaoAlvara(Contribuinte.SituacaoAlvara.PROVISORIO);
				c.setTipoEmpresa(Contribuinte.TipoEmpresa.MATRIZ);
				
				c.setCodigo(rs.getString("CONTRIB_CODIGO"));
				c.setNaturezaJuridica(tipoNaturezaJuridica(rs.getInt("CONTRIB_NAT_JURIDICA")));
				c.setNome(rs.getString("CONTRIB_NOME"));
				c.setNomeFantasia(rs.getString("CONTRIB_NOME"));
				c.setInscricaoAnterior(Utilidade.gerarInscricaoPeloCodigoExistente(rs.getString("CONTRIB_CODIGO")));
				c.setInscricaoAuxiliar(rs.getString("CONTRIB_INSC_MUNICIPAL"));
				c.setCpfCnpj(rs.getString("CONTRIB_CNPJ"));
				if(c.getCpfCnpj()==null || c.getCpfCnpj().trim().length()==0){
					c.setCpfCnpj(rs.getString("CONTRIB_CPF"));
					c.setTipoContribuinte(Contribuinte.TipoContribuinte.FISICO);
					
				}else{
					c.setTipoContribuinte(Contribuinte.TipoContribuinte.JURIDICO);
					
				}
				if(c.getCpfCnpj()==null || c.getCpfCnpj().trim().length()==0)
					c.setTipoContribuinte(Contribuinte.TipoContribuinte.NAO_IDENTIFICADO);
				
				
				c.setLogradouro(tipoLogradouro(rs.getString("CONTRIB_TIPO_LOGRADOURO")).toString());
				c.setNomeLogradouro(rs.getString("CONTRIB_NOME_LOGRA"));
				c.setNumero(rs.getString("CONTRIB_NUM_LOGRA"));
				Bairro bairro=ControladorCpSeg.buscarBairro(validarBairroRepetido((rs.getInt("CONTRIB_COD_BAIRRO"))));
				Logradouro logradouro=ControladorCpSeg.buscarLogradouro(rs.getInt("CONTRIB_COD_LOGRADOURO"));
				if(bairro!=null){
					c.setBairro(bairro);
					c.setNomeBairro(bairro.getNome());
				}
				c.setFkLogradouro(logradouro);
				c.setCep(rs.getString("CONTRIB_CEP"));
				c.setCidade(cidade);
				c.setUsuario("MIGRACAO");
				
				c.setComplemento("");
				c.setSituacao(Contribuinte.SituacaoCadastral.HABILITADO);
				c.setGrupoAtividade(Contribuinte.GrupoAtividade.COMERCIO);
				c.setAtividadePoder(Contribuinte.AtividadePoder.MUNICIPAL);
				c.setEstabelecido(Contribuinte.Estabelecido.SIM);
				c.setTipoRecolhimento(Contribuinte.TipoRecolhimento.RECOLHIMENTO_MENSAL);
				c.setPorteEmpresa(Contribuinte.PorteEmpresa.PEQUENO);
				c.setSituacaoAlvara(Contribuinte.SituacaoAlvara.PROVISORIO);
				c.setEscolaridade(Contribuinte.Escolaridade.ALFABETIZADO);
				
				c.setDataAlteracao(Utilidade.dataAtual());
				c.setDataCadastro(Utilidade.dataAtual());
				c.setInicioAtividade(Utilidade.dataAtual());
				c.setInicioPrestacaoServico(Utilidade.dataAtual());
				
				
				c.setGrupoCae(0);
				
				c.setRuc("");
				c.setConselho("");
				c.setImovelProprio(false);
				c.setRegistro(0);
				c.setNumeroEmpregados(0);
				c.setProtocolo("");
				c.setIsento(false);
				c.setFatorAlvara(0L);
				c.setRamo(0);
				c.setFoneFax("");
				c.setCnh("");
				c.setCategoriaCnh("");
				c.setAutorizacao("");
				c.setPontoRecepcao("");
				c.setTipoCadastro(0);
				c.setDataEncerramento(null);
				c.setTelefone("");
				c.setEmail("");
				c.setIdentidadeRg("");
				c.setOptanteSimplesNacional(false);
				
				
				/*ESTOU MIGRANDO TODOS DE CONTRIBUINTES, MAS ALGUMAS INFORMACOES
				IREI ACRESCENTAR AGORA DE CADASTRO MOBILIARIO VENDO SE EXISTE ALGUEM COM
				O NOME CORRELATIVO
				*/
				contribuintes.add(c);
				System.out.println("CONTRIBUINTES --- " + c.getNome());
				
				
		}
		return contribuintes;
	}
	private static Integer validarBairroRepetido(Integer codigo) {
		Integer codLocal=codigo;
		if(codigo==274 ||codigo==256 ||codigo==258){
			codLocal=7;//CENTRO
		}
		if(codigo==265){
			codLocal=8;//SAO FRANCISCO
		}
		if(codigo==269){
			codLocal=25;//ZONA RURAL
		}
		return codLocal;
	}
	private static Bairro buscarBairroCodo(Integer codigo) throws SQLException{
		Bairro bairro=null;
		//COMO ALGUNS BAIRROS ESTAO REPETIDOS EU OS DESCONSIDERAREI AQUI
		Integer codLocal=validarBairroRepetido(codigo);
		String sql="SELECT BAIRRO_CODIGO, BAIRRO_NOME FROM BAIRROS WHERE BAIRRO_CODIGO=" + codLocal;
		ResultSet rsB=conexaoCodo.createStatement().executeQuery(sql);
		while(rsB.next()){
			bairro=new Bairro();
			bairro.setCodigo(rsB.getInt("BAIRRO_CODIGO"));
			bairro.setNome(rsB.getString("BAIRRO_NOME"));
			Uf uf=new Uf();
			uf.setCodigo(10);
			uf.setNome("MARANHÃO");
			uf.setSigla("MA");
			Cidade cidade=new Cidade();
			cidade.setCodigo(1179);
			cidade.setNome("CODÓ");
			cidade.setUf(uf);
			bairro.setCidade(cidade);
		}
		rsB.close();
		return bairro;
	}
	private static List<Contribuinte> listarDaTabelaCadastroMobiliario(List<Contribuinte> contribuintes)throws SQLException{
		Statement stCM=conexaoCodo.createStatement();
		for(Contribuinte c: contribuintes){
			//PRIMEIRO SELECT PELO CODIGO
			String sqlCM ="select * from cadastro_mobiliario where mobil_cod_contribuinte='" + c.getCodigo() + "'";
			ResultSet rsCM=stCM.executeQuery(sqlCM);
			boolean encontrouPeloCodigo=false;
			while (rsCM.next()){
				atualizaCamposCM(c, rsCM);
				encontrouPeloCodigo=true;
			}
			//ENCONTRANDO NADA PELO NOME
			if(!encontrouPeloCodigo){//SE NAO ENCONTROU PELO CODIGO, PROCURA PELO NOME
				rsCM.close();
				sqlCM ="select * from cadastro_mobiliario where mobil_nome_razao='" + c.getNome() + "'";
				rsCM=stCM.executeQuery(sqlCM);
				while (rsCM.next()){
					atualizaCamposCM(c, rsCM);
				}
			}
			System.out.println("CADASTRO MOBILIARIO --- " + c.getNome());
			rsCM.close();
		}
		return contribuintes;
	}
	private static void atualizaCamposCM(Contribuinte c,ResultSet rsCM) throws SQLException{
		if(c.getCpfCnpj()==null || c.getCpfCnpj().trim().length()==0){
			c.setCpfCnpj(rsCM.getString("MOBIL_CNPJ"));
			if(c.getCpfCnpj()==null || c.getCpfCnpj().trim().length()==0){
				c.setTipoContribuinte(Contribuinte.TipoContribuinte.NAO_IDENTIFICADO);
			}else{
				if(c.getCpfCnpj().trim().length()<=11){
					c.setTipoContribuinte(Contribuinte.TipoContribuinte.FISICO);
				}else{
					c.setTipoContribuinte(Contribuinte.TipoContribuinte.JURIDICO);
				}
			}
		}
		//c.setInscricaoAnterior(rsCM.getString("MOBIL_INSC_MUNICIPAL"));
		c.setInscricaoAuxiliar(rsCM.getString("MOBIL_INSC_MUNICIPAL"));
		c.setNaturezaJuridica(tipoNaturezaJuridica(rsCM.getInt("MOBIL_COD_NAT_JURID")));
		c.setEstabelecido(tipoEstabelecido(rsCM.getInt("MOBIL_ESTABELECIMENTO")));
		c.setInicioAtividade(rsCM.getDate("MOBIL_INICIO_ATIVIDADE"));
		c.setFimAtividade(rsCM.getDate("MOBIL_FIM_ATIVIDADE"));
		c.setNumeroEmpregados(rsCM.getInt("MOBIL_QTE_EMPREGADOS"));
		c.setCae(rsCM.getString("MOBIL_COD_ATIVIDADE1"));
		c.setCaeSecundario(rsCM.getString("MOBIL_COD_ATIVIDADE2"));
		c.setCaeTerciario(rsCM.getString("MOBIL_COD_ATIVIDADE3"));
		c.setEscolaridade(tipoEscolaridade(rsCM.getString("MOBIL_COD_ESCOLARIDADE")));
		c.setDataCadastro(rsCM.getDate("MOBIL_DATA_CADASTRO"));
	}
	private static Contribuinte.Escolaridade tipoEscolaridade(String codigo){
		Contribuinte.Escolaridade escolaridade=Contribuinte.Escolaridade.ALFABETIZADO;
		if(codigo=="1")
			escolaridade=Contribuinte.Escolaridade.NAO_ALFABETIZADO;
		else if(codigo=="3")
			escolaridade=Contribuinte.Escolaridade.FUNDAMENTAL_COMPLETO;
		else if(codigo=="5")
			escolaridade=Contribuinte.Escolaridade.MEDIO_COMPLETO;
		else if(codigo=="6")
			escolaridade=Contribuinte.Escolaridade.SUPERIOR_INCOMPLETO;
		else if(codigo=="7")
			escolaridade=Contribuinte.Escolaridade.SUPERIOR_COMPLETO;
		else if(codigo=="8")
			escolaridade=Contribuinte.Escolaridade.ESPECIALIZADO;
		return escolaridade;
	}
	private static Contribuinte.NaturezaJuridica tipoNaturezaJuridica(Integer codigo){
		Contribuinte.NaturezaJuridica natureza=Contribuinte.NaturezaJuridica.PESSOA_FISICA;
		if(codigo==0)
			natureza=Contribuinte.NaturezaJuridica.PESSOA_FISICA;
		else if (codigo==1)
			natureza=Contribuinte.NaturezaJuridica.PESSOA_JURIDICA;
		return natureza;
	}
	private static Contribuinte.TipoContribuinte tipoContribuinte(Integer codigo){
		Contribuinte.TipoContribuinte tipoContribuinte=Contribuinte.TipoContribuinte.FISICO;
		if(codigo==0)
			tipoContribuinte=Contribuinte.TipoContribuinte.FISICO;
		else if (codigo==1)
			tipoContribuinte=Contribuinte.TipoContribuinte.JURIDICO;
		return tipoContribuinte;
	}
	private static Contribuinte.Estabelecido tipoEstabelecido(Integer codigo){
		Contribuinte.Estabelecido estabelecido=Contribuinte.Estabelecido.NAO;
		if(codigo==0)
			estabelecido=Contribuinte.Estabelecido.NAO;
		else if (codigo==1)
			estabelecido=Contribuinte.Estabelecido.SIM;
		return estabelecido;
	}
	private static Logradouro.TipoLogradouro tipoLogradouro(String descricao){
		Logradouro.TipoLogradouro tipo=Logradouro.TipoLogradouro.RUA;
		descricao=descricao.trim();
		if(descricao.equals("AVN"))
			tipo=Logradouro.TipoLogradouro.AVENIDA;
		else if(descricao.equals("CNJ"))
			tipo=Logradouro.TipoLogradouro.CONJUNTO;
		else if(descricao.equals("PCA"))
			tipo=Logradouro.TipoLogradouro.PRACA;
		else if(descricao.equals("POV"))
			tipo=Logradouro.TipoLogradouro.POVOADO;
		else if(descricao.equals("RUA"))
			tipo=Logradouro.TipoLogradouro.RUA;
		else if(descricao.equals("TRV"))
			tipo=Logradouro.TipoLogradouro.TRAVESSA;
		else if(descricao.equals("VILA"))
			tipo=Logradouro.TipoLogradouro.VILA;
		else
			tipo=Logradouro.TipoLogradouro.RUA;
		return tipo;
	}
	//-IMOVEIS
	public static List<Imovel> listarImoveis() throws SQLException{
		Statement st=conexaoCodo.createStatement();
		String sql="SELECT * FROM IMOVEIS ORDER BY IMOVEL_INSCRICAO";
		ResultSet rs=st.executeQuery(sql);
		List<Imovel> imoveis=new ArrayList();
		while(rs.next()){
			Imovel i=new Imovel();
			
			Contribuinte c=new Contribuinte();
			c.setInscricao(Utilidade.gerarInscricaoPeloCodigoExistente(rs.getString("IMOVEL_PROP_CODIGO")));
			//ESSA FASE DEPOIS QUE RODAR A ATUALIZACAO DAS INSCRICOES PELO GESTOR
			c=ControladorCpTrib.buscarContribuinteInscricaoAnterior(c.getInscricao());
			i.setContrbuinte(c);
			Logradouro l=ControladorCpSeg.buscarLogradouro(rs.getInt("IMOVEL_LOC_COD_LOGRA"));
			i.setTipoImovel(Imovel.TipoImovel.PROPRIO);
			i.setInscricao(rs.getString("IMOVEL_INSCRICAO"));
			i.setUnidade(0.0d);
			i.setInscricaoAuxiliar("");
			i.setContrbuinte(c);
			i.setInscricaoAnterior(rs.getString("IMOVEL_INSCRICAO"));
			i.setLogradouro(l);
			i.setNumero(rs.getString("IMOVEL_LOC_NUM_LOGRA"));
			i.setComplemento(rs.getString("IMOVEL_LOC_COMPL_LOGRA"));
			i.setCep("65400000");
			i.setAnoAquisicao(rs.getInt("IMOVEL_EXERCICIO"));
			i.setValor(rs.getDouble("IMOVEL_VVI"));
			i.setOcupante(c==null?null:c.getNome());
			i.setCpfCnpjOcupante(c==null?null:c.getCpfCnpj());
			i.setInscricaoCondominio("0");
			i.setBairro(l==null?null:l.getBairro());
			i.setCodMensagem(0);
			i.setZona(0);
			i.setValorTerreno(rs.getDouble("IMOVEL_VVT"));
			i.setValorEdificao(rs.getDouble("IMOVEL_VVE"));
			i.setLoteamento("");
			i.setSecao(rs.getString("IMOVEL_LOC_SECAO_LOGRA"));
			i.setQuadra(rs.getString("IMOVEL_LOC_QUADRA"));
			i.setLote(rs.getString("IMOVEL_LOC_LOTE"));
			i.setSituacaoCadastral(Imovel.SituacaoCadastral.HABILITADO);
			i.setAforado(rs.getBoolean("IMOVEL_INFGER_AFORAMENTO"));
			i.setObservacao("");
			i.setRegistroAforamento("");
			i.setDataRegistroAforamento(null);
			i.setNumeroAforamento("");
			i.setFichaAforamento("");
			i.setLivroAforamenteo("");
			i.setFolhaAforamento("");
			i.setDataAforamento(null);
			i.setMotivoAlteracao("");
			i.setValorTerrenoMercado(rs.getDouble("IMOVEL_VVE"));
			i.setValorEdificadoMercado(0.0);
			i.setSituacaoLote(Imovel.SituacaoLote.TERRITORIAL);
			i.setSubUnidade("");
			i.setUsuario("MIGRACAO");
			i.setDataCadastro(new java.sql.Date(Utilidade.dataAtual().getTime()));
			i.setTipoBoletim("");
			i.setCodEdificio(0);
			i.setBloco("");
			i.setApto("");
			i.setSalaLoja("");
			i.setSequenciaTrecho("");
			i.setCodigoTrecho("");
			imoveis.add(i);
		}
		rs.close();
		st.close();
		return imoveis;
	}
	public static List<DetalheImovel> listarDetalhesImovel(List<Imovel> imoveis) throws SQLException{
		List<DetalheImovel> detalhes=new ArrayList();
		Statement st=conexaoCodo.createStatement();
		DetalheImovel di=null;
		for(Imovel i: imoveis){
			String sql="SELECT * FROM IMOVEIS WHERE IMOVEL_INSCRICAO = '" + i.getInscricao() + "'";
			ResultSet rs=st.executeQuery(sql);
			while(rs.next()){
				di=new DetalheImovel(i,GrupoComponente.TERRENO_OCUPACAO_LOTE,tipoOpcupacaoLote(rs.getInt("IMOVEL_INFGER_OCUPACAO")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TERRENO_PATRIMONIO,tipoPatrimonio(rs.getInt("IMOVEL_INFGER_PATRIMONIO")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TERRENO_LIMITE,DetalheImovel.TerrenoLimite.SEM_MURO_CERCA);
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.EDIFICACAO_TIPOLOGIA,tipoTipologia(rs.getInt("IMOVEL_TIPO")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TERRENO_SITUACAO,tipoSituacao(rs.getInt("IMOVEL_INFTER_SITUACAO")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TERRENO_TOPOGRAFIA,tipoTopografia(rs.getInt("IMOVEL_INFTER_TOPOGRAFIA")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TERRENO_PEDOLOGIA,tipoPedologia(rs.getInt("IMOVEL_INFTER_PEDOLOGIA")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.EDIFICACAO_DESTINACAO,tipoDestinacao(rs.getInt("IMOVEL_INFGER_UTILIZACAO")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.EDIFICACAO_ESTRUTURA,tipoEstrutura(rs.getInt("IMOVEL_ESTRUTURA")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.EDIFICACAO_PADRAO,DetalheImovel.EdificacaoPadrao.MEDIO);
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.EDIFICACAO_CONSERVACAO,tipoConservacao(rs.getInt("IMOVEL_ESTADO_CONSERVACAO")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TESTADA_PRINCIPAL,rs.getDouble("IMOVEL_TESTADA_PRINCIPAL"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.EDIFICACAO_TIPO_COBRANCA,tipoCobranca(rs.getInt("IMOVEL_INFGER_ISENTO_IPTU")));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TESTADA_2,rs.getDouble("IMOVEL_TESTADA_2"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TESTADA_3,rs.getDouble("IMOVEL_TESTADA_3"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.TESTADA_4,rs.getDouble("IMOVEL_TESTADA_4"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.AREA_LOTE,rs.getDouble("IMOVEL_AREA_TOTAL_TERRENO"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.PAVIMENTOS,0.0d);
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.ANO_CONTRUCAO,rs.getDouble("IMOVEL_INFGER_ANO_CONSTR"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.AREA_EDIFICADA_UNIDADE,rs.getDouble("IMOVEL_AREA_CONST_UNIDADE"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.AREA_EDIFICADA_TOTAL,rs.getDouble("IMOVEL_AREA_TOTAL_CONSTRUIDA"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.PROFUNDIDADE,rs.getDouble("IMOVEL_PROFUNDIDADE"));
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.UNIDADE_FISCAL,0.0d);
				detalhes.add(di);
				di=new DetalheImovel(i,GrupoComponente.VALOR_FIXO,0.0d);
				
				detalhes.add(di);
			}
			rs.close();
		}
		return detalhes;	
	}
	private static DetalheImovel.TerrenoOcupacaoLote tipoOpcupacaoLote(Integer codigo){
		for(DetalheImovel.TerrenoOcupacaoLote opcao : DetalheImovel.TerrenoOcupacaoLote.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;
        }   
		return DetalheImovel.TerrenoOcupacaoLote.CONSTRUIDO; 
	}
	private static DetalheImovel.TerrenoPatrimonio tipoPatrimonio(Integer codigo){
		for(DetalheImovel.TerrenoPatrimonio opcao : DetalheImovel.TerrenoPatrimonio.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;   
        }   
        return DetalheImovel.TerrenoPatrimonio.PARTICULAR;
	}
	private static DetalheImovel.EdificacaoDestinacao tipoDestinacao(Integer codigo){
		for(DetalheImovel.EdificacaoDestinacao opcao : DetalheImovel.EdificacaoDestinacao.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;   
        }   
		return DetalheImovel.EdificacaoDestinacao.COMERCIAL;
	}
	private static DetalheImovel.TerrenoLimite tipoLimite(Integer codigo){
		for(DetalheImovel.TerrenoLimite opcao : DetalheImovel.TerrenoLimite.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;   
        }   
		return DetalheImovel.TerrenoLimite.SEM_MURO_CERCA;
	}
	private static DetalheImovel.TerrenoSituacao tipoSituacao(Integer codigo){
		DetalheImovel.TerrenoSituacao opcao=DetalheImovel.TerrenoSituacao.MEIO_DE_QUADRA_COM_UMA_FRENTE;
		switch(codigo){
			case 1:
				opcao=DetalheImovel.TerrenoSituacao.MEIO_DE_QUADRA_COM_UMA_FRENTE;
				break;
			case 2:
				opcao=DetalheImovel.TerrenoSituacao.MEIO_DE_QUADRA_COM_DUAS_FRENTES;
				break;
			case 3:
				opcao=DetalheImovel.TerrenoSituacao.ESQUINA_COM_MAIS_DE_UMA_FRENTE;
				break;
			case 4:
				opcao=DetalheImovel.TerrenoSituacao.ENCRAVADOS;
				break;
			case 5:
				opcao=DetalheImovel.TerrenoSituacao.FUNDOS;
				break;
			case 6:
				opcao=DetalheImovel.TerrenoSituacao.ENCRAVADOS;
				break;
			case 7:
				opcao=DetalheImovel.TerrenoSituacao.GLEBA;
				break;
				
		}
		return opcao;
	}
	private static DetalheImovel.TerrenoTopografia tipoTopografia(Integer codigo){
		for(DetalheImovel.TerrenoTopografia opcao : DetalheImovel.TerrenoTopografia.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;   
        }   
        return DetalheImovel.TerrenoTopografia.PLANA;
	}
	private static DetalheImovel.TerrenoPedologia tipoPedologia(Integer codigo){
		for(DetalheImovel.TerrenoPedologia opcao : DetalheImovel.TerrenoPedologia.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;   
        }   
        return DetalheImovel.TerrenoPedologia.FIRME; 
	}
	private static DetalheImovel.EdificacaoTipologia tipoTipologia(Integer codigo){
		DetalheImovel.EdificacaoTipologia tipologia=DetalheImovel.EdificacaoTipologia.CASA;
		switch(codigo){
			case 1:
				tipologia=DetalheImovel.EdificacaoTipologia.CASA;
				break;
			case 2:
				tipologia=DetalheImovel.EdificacaoTipologia.CONSTRUCAO_PRECARIA;
				break;
			case 3:
				tipologia=DetalheImovel.EdificacaoTipologia.BARRACO;
				break;
			case 4:
				tipologia=DetalheImovel.EdificacaoTipologia.APARTAMENTO;
				break;
			case 5:
				tipologia=DetalheImovel.EdificacaoTipologia.LOJA_SALA_CONJUNTO;
				break;
			case 6:
				tipologia=DetalheImovel.EdificacaoTipologia.GALPAO;
				break;
			case 7:
				tipologia=DetalheImovel.EdificacaoTipologia.TALHEIRO;
				break;
			case 8:
				tipologia=DetalheImovel.EdificacaoTipologia.DEPOSITO;
				break;
			case 9:
				tipologia=DetalheImovel.EdificacaoTipologia.OUTROS;
		}
		return tipologia;
	}
	private static DetalheImovel.EdificacaoEstrutura tipoEstrutura(Integer codigo){
		for(DetalheImovel.EdificacaoEstrutura opcao : DetalheImovel.EdificacaoEstrutura.values()) {   
            if(opcao.getValor() == codigo)   
                return opcao;   
        }   
		return DetalheImovel.EdificacaoEstrutura.CONCRETO;  
	}
	private static DetalheImovel.EdificacaoConservacao tipoConservacao(Integer codigo){
		DetalheImovel.EdificacaoConservacao opcao=DetalheImovel.EdificacaoConservacao.BOA;
		switch(codigo){
			case 1:
				opcao=DetalheImovel.EdificacaoConservacao.OTIMO;
				break;
			case 2:
				opcao=DetalheImovel.EdificacaoConservacao.BOA;
				break;
			case 3:
				opcao=DetalheImovel.EdificacaoConservacao.REGULAR;
				break;
			case 4:
				opcao=DetalheImovel.EdificacaoConservacao.PRECARIO;
				break;
			case 5:
				opcao=DetalheImovel.EdificacaoConservacao.PRECARIO;
		}
		return opcao;
	}
	private static DetalheImovel.EdificacaoTipoCobranca tipoCobranca(Integer codigo){
		DetalheImovel.EdificacaoTipoCobranca opcao=DetalheImovel.EdificacaoTipoCobranca.NORMAL;
		switch(codigo){
			case 1:
				opcao=DetalheImovel.EdificacaoTipoCobranca.NORMAL;
				break;
			case 2:
				opcao=DetalheImovel.EdificacaoTipoCobranca.IMUNE;
				break;
			case 3:
				opcao=DetalheImovel.EdificacaoTipoCobranca.ISENTO;
		}
		return opcao;
	}
}
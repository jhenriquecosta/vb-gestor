package com.bcptecnologias.migracao.gestortributario.fachada;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;

import com.bcptecnologias.migracao.gestortributario.negocio.Bairro;
import com.bcptecnologias.migracao.gestortributario.negocio.Contribuinte;
import com.bcptecnologias.migracao.gestortributario.negocio.Imovel;
import com.bcptecnologias.migracao.gestortributario.negocio.Logradouro;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.DetalheImovel;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.NumeroCorrelativo;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.TabelaCpTrib;
import com.bcptecnologias.migracao.gestortributario.util.Utilidade;
import com.microsoft.sqlserver.jdbc.SQLServerException;
import com.sun.xml.internal.ws.api.model.wsdl.WSDLBoundOperation.ANONYMOUS;

public class ControladorCpTrib {
	private static Connection conexao;
	public static void setConexao(Connection cnd){
		conexao=cnd;
	}
	public static Contribuinte buscarContribuinteInscricaoAnterior(Object valor) throws SQLException{
		return buscarContribuinte("TCI_INSCRICAO_ANTERIOR", valor);
	}
	public static Contribuinte buscarContribuinte(String campo, Object valor)throws SQLException{
		Contribuinte c=null;
		Statement st=conexao.createStatement();
		String sql="SELECT tci_im, TCI_INSCRICAO_ANTERIOR, tci_im_auxiliar FROM TAB_CONTRIBUINTE WHERE " + campo +  "='" + valor.toString() +"'";
		ResultSet rs=st.executeQuery(sql);
		while(rs.next()){
			c=new Contribuinte();
			c.setInscricao(rs.getString("TCI_IM" ));
			c.setInscricaoAnterior(rs.getString("TCI_INSCRICAO_ANTERIOR"));
			c.setInscricaoAuxiliar(rs.getString("tci_im_auxiliar"));
		}
		rs.close();
		st.close();
		return c;
	}
	public static boolean encontrouContribuintePeloNome(String nome)throws SQLException{
		Statement st=conexao.createStatement();
		String sql="SELECT TCI_NOME FROM TAB_CONTRIBUINTE WHERE TCI_NOME='"  + nome + "'";
		ResultSet rs=st.executeQuery(sql);
		boolean encontrado=false;
		while (rs.next()){
			encontrado=true;
		}
		rs.close();
		return encontrado;
	}
	public static Integer migrarContribuintes(List<Contribuinte> contribuintes) throws SQLException{
		String sql="insert into tab_contribuinte(" +
			"tci_im, TCI_INSCRICAO_ANTERIOR, tci_im_auxiliar, tci_cgc_cpf, tci_nome, tci_fantasia, tci_logradouro, " +
			"tci_nome_logradouro, tci_numero, tci_complemento,tci_bairro, tci_cep, tci_cidade, tci_UF, tci_data_cadastro, " +
			"tci_data_modific, tci_tsc_cod_sit_cad, tci_tus_cod_usuario, tci_inicio_atividade, tci_tga_cod_grupo," +
			"tci_tnj_cod_natureza, tci_tae_cae, tci_tap_cod_ativ_poder, tci_estab, tci_grupo_cae, tci_tipo_contribuinte, " +
			"tci_tipo_recolhimento_iss, tci_ruc, tci_conselho,tci_imovel_proprio, tci_registro, tci_num_empregado, " +
			"tci_porte_empresa, tci_tae_cae_secund, tci_tae_cae_terc, tci_nivel_escolar, tci_protocolo, tci_tim_ic," + 
			"TCI_ISENTO, TCI_FATOR_ALVARA, TCI_TRA_COD_RAMO, TCI_FONE_FAX, TCI_CNH, TCI_CATEGORIA, TCI_AUTORIZACAO, " +
			"TCI_PONTO_RECEPCAO,TCI_THF_COD_HORARIO, TCI_COD_LOGRADOURO, TCI_COD_BAIRRO, TCI_TIPO_PESSOA, TCI_ALVARA_LIBERADO, " +
			"TCI_TAE_CAE_TRANSPORTE, tci_tipo_cadastro,TCI_INICIO_PREST_SERV, TCI_SIT_ALVARA, tci_matriz_filial, " +
			"tci_data_encerramento, tci_data_reabertura, tci_telefone, tci_email, TCI_RG, tci_optante_simples" +
        ") "+
			"values("+ Utilidade.criarInterrogacoes(62) +")";
		Integer migrados=0;
		PreparedStatement pst=conexao.prepareStatement(sql);
		for(Contribuinte c: contribuintes){
			if(!ControladorCpTrib.encontrouContribuintePeloNome(c.getNome())){
				pst.setString(1, c.getInscricao());
				pst.setString(2, c.getInscricaoAnterior());
				pst.setString(3, c.getInscricaoAuxiliar());
				pst.setString(4, c.getCpfCnpj());
				pst.setString(5, c.getNome());
				pst.setString(6, c.getNomeFantasia());
				pst.setString(7, c.getLogradouro());
				pst.setString(8, c.getNomeLogradouro());
				pst.setString(9, c.getNumero());
				pst.setString(10, c.getComplemento());
				pst.setString(11, c.getNomeBairro());
				pst.setString(12, c.getCep());
				pst.setString(13, c.getCidade().getNome());
				pst.setString(14, c.getCidade().getUf().getSigla());
				pst.setInt(17, c.getSituacao().getValor());
				pst.setString(18, c.getUsuario());
				pst.setInt(20, c.getGrupoAtividade().getValor());
				pst.setInt(21, c.getNaturezaJuridica().getValor());
				pst.setString(22, c.getCae());
				pst.setInt(23, c.getAtividadePoder().getValor());
				pst.setInt(24, c.getEstabelecido().getValor());
				pst.setInt(25, c.getGrupoCae());
				pst.setInt(26, c.getTipoContribuinte().getValor());
				pst.setInt(27, c.getTipoRecolhimento().getValor());
				pst.setString(28, c.getRuc());
				pst.setString(29, c.getConselho());
				pst.setInt(30, c.getTipoImovel().getValor());
				pst.setInt(31, c.getRegistro());
				pst.setInt(32, c.getNumeroEmpregados());
				pst.setInt(33, c.getPorteEmpresa().getValor());
				pst.setString(34, c.getCaeSecundario());
				pst.setString(35, c.getCaeTerciario());
				pst.setInt(36, c.getEscolaridade().getValor());
				pst.setString(37, c.getProtocolo());
				pst.setString(38, "");
				pst.setInt(39, c.isIsento()? 1:0);
				pst.setLong(40, c.getFatorAlvara());
				pst.setInt(41, c.getRamo());
				pst.setString(42, c.getFoneFax());
				pst.setString(43, c.getCnh());
				pst.setString(44, c.getCategoriaCnh());
				pst.setString(45, c.getAutorizacao());
				pst.setString(46, c.getPontoRecepcao());
				pst.setInt(47, 0);
				pst.setInt(48, c.getFkLogradouro().getCodigo());
				//PEGAR O BAIRRO
				pst.setInt(49, c.getBairro().getCodigo());
				pst.setInt(50, c.getTipoEmpresa().getValor());
				pst.setInt(51,0);
				pst.setString(52,"0");
				pst.setInt(53, c.getTipoCadastro());
				pst.setInt(55, c.getSituacaoAlvara().getValor());
				pst.setInt(56, c.getTipoEmpresa().getValor());
				pst.setTimestamp(57, c.getDataEncerramento()==null?null: new java.sql.Timestamp(c.getDataEncerramento().getTime()));
				pst.setDate(58, null);
				pst.setString(59, c.getTelefone());
				pst.setString(60, c.getEmail());
				pst.setString(61, c.getIdentidadeRg());
				pst.setInt(62, c.isOptanteSimplesNacional()? 1:0);
				//pst.setString(63, c.getCodigoOrigem());
				
				Date dataAtual = Utilidade.dataAtual();
				pst.setTimestamp(15,new java.sql.Timestamp(dataAtual.getTime()));
				pst.setTimestamp(16,new java.sql.Timestamp(dataAtual.getTime()));
				pst.setTimestamp(19,new java.sql.Timestamp(dataAtual.getTime()));
				pst.setTimestamp(54, new java.sql.Timestamp(dataAtual.getTime()));
				pst.execute();
				migrados++;
				System.out.println("INCLUINDO CONTRIBUINTE --- " + c.getNome());
			}else
				System.out.println("CONTRIBUINTE " + c.getNome() + " --- JA EXISTENTE EM NOSSA BASE ");
		}
		pst.close();
		return migrados;
	}
	public static boolean encontrouImovelPelaInscricao(String inscricao)throws SQLException{
		Statement st=conexao.createStatement();
		String sql="SELECT TIM_IC FROM TAB_IMOVEL WHERE TIM_IC='"  + inscricao + "'";
		ResultSet rs=st.executeQuery(sql);
		boolean encontrado=false;
		while (rs.next()){
			encontrado=true;
		}
		rs.close();
		return encontrado;
	}
	public static Integer migrarImoveis(List<Imovel> imoveis) throws SQLException{
		String sql="INSERT INTO TAB_IMOVEL("+
		"tim_ic, tim_unidade, tim_ic_auxiliar, tim_tci_im, tim_tci_im_auxiliar, TIM_TCI_INSCRICAO_ANTERIOR, TIM_TCI_CADASTRO_ANTERIOR, tim_ic_anterior,"+ 
        "tim_tlg_cod_logradouro, tim_numero, tim_complemento, tim_cep, tim_ano_aquis, tim_valor, tim_tipo_imovel, tim_ocupante, tim_cgc_cpf_ocupante,"+ 
        "TIM_IC_CONDOMINIO, TIM_TTL_COD_TIP_LOGR, TIM_TBA_COD_BAIRRO, TIM_COD_MENSAGEM, TIM_ZONA, TIM_VALOR_TERRENO, TIM_VALOR_EDIFIC,"+ 
        "tim_loteamento, tim_secao, TIM_QUADRA, tim_lote, tim_tsc_cod_sit_cad, tim_aforado, tim_obs, TIM_AFORAMENTO_REGISTRO, TIM_DATA_REGISTRO,"+ 
        "TIM_AFORAMENTO_NUMERO, TIM_AFORAMENTO_FICHA, TIM_AFORAMENTO_LIVRO, TIM_AFORAMENTO_FOLHA, TIM_AFORAMENTO_DATA,"+ 
        "TIM_MOTIVO_ALTERACAO, TIM_VALOR_TERRENO_MERCADO, TIM_VALOR_EDIFICACAO_MERCADO, TIM_SITUACAO_LOTE, TIM_SUB_UNIDADE,"+ 
        "TIM_TUS_COD_USUARIO, TIM_DATA_CADASTRO, TIM_TIPO_BOLETIM, TIM_TED_COD_EDIFICIO, TIM_BLOCO, TIM_APTO, TIM_SALA_LOJA, TIM_TTC_SEQ_TRECHO,"+ 
        "TIM_TTC_COD_TRECHO, TIM_TUS_USUARIO"+
		") "+
			"values("+ Utilidade.criarInterrogacoes(53) +")"; 
		PreparedStatement pst=conexao.prepareStatement(sql);
		Integer migrados=0;
		for(Imovel i:imoveis){
			if(!encontrouImovelPelaInscricao(i.getInscricao())){
				Logradouro l=i.getLogradouro();
				Contribuinte c=i.getContrbuinte();
				pst.setString(1, i.getInscricao());
				pst.setDouble(2, i.getUnidade());
				pst.setString(3, i.getInscricaoAuxiliar());
				pst.setString(4, c==null?null:c.getInscricao());
				pst.setString(5, c==null?null:c.getInscricaoAuxiliar());
				pst.setString(6, c==null?null:c.getInscricaoAnterior());
				pst.setString(7, c==null?null:c.getInscricaoAnterior());
				pst.setString(8, i.getInscricaoAnterior());
				pst.setString(9, l==null?null:l.getCodigo().toString());
				pst.setString(10, i.getNumero());
				pst.setString(11, i.getComplemento());
				pst.setString(12, i.getCep());
				pst.setInt(13, i.getAnoAquisicao());
				pst.setDouble(14, i.getValor());
				pst.setInt(15, i.getTipoImovel().getValor());
				pst.setString(16, i.getOcupante());
				pst.setString(17, i.getCpfCnpjOcupante());
				pst.setString(18, i.getInscricaoCondominio());
				pst.setInt(19, l==null?null:l.getTipoLogradouro().getValor());
				pst.setInt(20, i.getBairro()==null?null:i.getBairro().getCodigo());
				pst.setInt(21, i.getCodMensagem());
				pst.setInt(22, i.getZona());
				pst.setDouble(23, i.getValorTerreno());
				pst.setDouble(24, i.getValorEdificao());
				pst.setString(25, i.getLoteamento());
				pst.setString(26, i.getSecao());
				pst.setString(27, i.getQuadra());
				pst.setString(28, i.getLote());
				pst.setInt(29, i.getSituacaoCadastral().getValor());
				pst.setString(30, i.getAforado()?"1":"0");
				pst.setString(31, i.getObservacao());
				pst.setString(32, i.getRegistroAforamento());
				pst.setDate(33, i.getDataRegistroAforamento()==null?null: new java.sql.Date(i.getDataRegistroAforamento().getTime()));
				pst.setString(34, i.getNumeroAforamento());
				pst.setString(35, i.getFichaAforamento());
				pst.setString(36, i.getLivroAforamenteo());
				pst.setString(37, i.getFolhaAforamento());
				pst.setTimestamp(38, i.getDataAforamento()==null?null: new java.sql.Timestamp(i.getDataAforamento().getTime()));
				pst.setString(39, i.getMotivoAlteracao());
				pst.setDouble(40, i.getValorTerrenoMercado());
				pst.setDouble(41, i.getValorEdificadoMercado());
				pst.setInt(42, i.getSituacaoLote().getValor());
				pst.setString(43, i.getSubUnidade());
				pst.setString(44, i.getUsuario());
				pst.setTimestamp(45, i.getDataCadastro()==null?null: new java.sql.Timestamp(i.getDataCadastro().getTime()));
				pst.setString(46,i.getTipoBoletim());
				pst.setInt(47, i.getCodEdificio());
				pst.setString(48, i.getBloco());
				pst.setString(49, i.getApto());
				pst.setString(50, i.getSalaLoja());
				pst.setString(51, i.getSequenciaTrecho());
				pst.setString(52, i.getCodigoTrecho());
				pst.setString(52, i.getSequenciaTrecho());
				pst.setString(53, i.getUsuario());
				pst.execute();
				migrados++;
				System.out.println("INCLUINDO IMOVEL --- " + i.getInscricao() + " -- "+ i.getLogradouro().getNome() + " -- " + i.getLogradouro().getBairro().getNome());
			}else{
				System.out.println("IMOVEL --- " + i.getInscricao() + " -- "+ i.getLogradouro().getNome() + " -- " + i.getLogradouro().getBairro().getNome() + " JA MIGRADO");
			}
				
		}
		pst.close();
		return migrados;
	}
	public static Integer migrarDetalheImoveis(List<DetalheImovel> detalhes) throws SQLException{
		Integer migrados=0;
		String sql="INSERT INTO TAB_DETALHE_IMOVEL(tdi_tim_ic,tdi_tgc_cod_grupo, tdi_tco_cod_componente,tdi_valor_item,  " +
		" tdi_valor_calc, tdi_tim_ic_unidade,tdi_tim_sub_unidade"+
		") "+
		"values("+ Utilidade.criarInterrogacoes(7) +")";
		PreparedStatement pst=conexao.prepareStatement(sql);
		for(DetalheImovel di:detalhes){
			pst.setString(1, di.getImovel().getInscricao());
			pst.setInt(2,di.getGrupoComponente().getValor());
			pst.setInt(3,di.getComponente().getValor());
			pst.setDouble(4,di.getValorItem());
			pst.setInt(5, 0);
			pst.setInt(6, 0);
			pst.setInt(7, 0);
			pst.execute();
			migrados++;
		}
		return migrados;
	}
	public static void resetaNumCorrelativa()throws SQLException{
		for(NumeroCorrelativo num: NumeroCorrelativo.values()){
			String sql="UPDATE " + TabelaCpTrib.TAB_NUM_CORRELATIVO + " SET TNC_SEQUENCIA = " + num.getValorPadrao() + " WHERE TNC_TIPO_OPERACAO=" + num.getOperacao();
			conexao.createStatement().execute(sql);
		}
	}
	public static void limparBaseCPTRIB() throws SQLException{
		for(TabelaCpTrib tab: TabelaCpTrib.values()){
			if(!tab.equals(TabelaCpTrib.TAB_NUM_CORRELATIVO))
				conexao.createStatement().execute(Utilidade.instrucaoDelete(tab.getNome()));
		}
	}
	public static String gerarInscricaoContribuinte() throws SQLException{
		Integer operacao=NumeroCorrelativo.INSCRICAO_CONTRIBUINTE.getOperacao();
		Statement st=conexao.createStatement();
		
		String sql="";
		ResultSet rs=null;
	    String inscricao="";
	    Double posicao=0.0d;
	    sql = "Update Tab_Num_Correlativo set tnc_sequencia = tnc_sequencia+1 where tnc_tipo_operacao = " + operacao;
	    st.execute(sql);
	    sql = "SELECT tnc_sequencia from tab_num_correlativo where tnc_tipo_operacao = " + operacao;
	    rs=st.executeQuery(sql);
	    while (rs.next()){
	    	posicao=new Double(rs.getDouble(1))==null?1:rs.getDouble(1);
	    }
	    rs.close();
	    st.close();
	    
	    //MONTAR INSCRICAO
	    inscricao=operacao+"";
	    return inscricao;
	    
	    /*
	    inscricao = operacao + Format(posicao, "000000")
	    inscricao = Inscricao & CStr(GeraDV(Inscricao, 9))
	    inscricao = Inscricao & GeraDV(Inscricao, 10)
	    GeraInscMunicipal = Format(Inscricao, "00000000-00")
	    */
	}
	
}

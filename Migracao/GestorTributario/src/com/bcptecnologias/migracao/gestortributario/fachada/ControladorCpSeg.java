package com.bcptecnologias.migracao.gestortributario.fachada;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

import com.bcptecnologias.migracao.gestortributario.negocio.Bairro;
import com.bcptecnologias.migracao.gestortributario.negocio.Cidade;
import com.bcptecnologias.migracao.gestortributario.negocio.Logradouro;
import com.bcptecnologias.migracao.gestortributario.negocio.Uf;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.TabelaCpSeg;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.TabelaCpTrib;
import com.bcptecnologias.migracao.gestortributario.util.Utilidade;

public class ControladorCpSeg {
	private static Connection conexao;
	public static void setConexao(Connection cnd){
		conexao=cnd;
	}
	public static Uf buscarUf(Integer codigo) throws SQLException{
		Uf uf=null;
		String sql="SELECT TUF_COD_UF, TUF_NOME, TUF_UF FROM TAB_UF WHERE TUF_COD_UF=" + codigo;
		ResultSet rs=conexao.createStatement().executeQuery(sql);
		while(rs.next()){
			uf=new Uf();
			uf.setCodigo(rs.getInt("TUF_COD_UF"));
			uf.setNome(rs.getString("TUF_NOME"));
			uf.setSigla(rs.getString("TUF_UF"));
		}
		rs.close();
		return uf;
	}
	
	public static Cidade buscarCidade(Integer codigo)throws SQLException{
		Cidade cidade=null;
		String sql="SELECT TMU_COD_MUNICIPIO, TMU_TUF_COD_UF, TMU_NOME FROM TAB_MUNICIPIO WHERE TMU_COD_MUNICIPIO=" + codigo;
		ResultSet rs=conexao.createStatement().executeQuery(sql);
		while(rs.next()){
			cidade=new Cidade();
			cidade.setCodigo(rs.getInt("TMU_COD_MUNICIPIO"));
			cidade.setNome(rs.getString("TMU_NOME"));
			cidade.setUf(buscarUf(rs.getInt("TMU_TUF_COD_UF")));
		}
		rs.close();
		return cidade;
	}
	public static Bairro buscarBairro(Integer codigo)throws SQLException{
		Bairro bairro=null;
		String sql="SELECT TBA_TMU_COD_MUNICIPIO, TBA_COD_BAIRRO, TBA_NOME FROM TAB_BAIRRO WHERE TBA_COD_BAIRRO=" + codigo;
		ResultSet rs=conexao.createStatement().executeQuery(sql);
		while(rs.next()){
			bairro=new Bairro();
			bairro.setCodigo(rs.getInt("TBA_COD_BAIRRO"));
			bairro.setNome(rs.getString("TBA_NOME"));
			bairro.setCidade(buscarCidade(rs.getInt("TBA_TMU_COD_MUNICIPIO")));
		}
		rs.close();
		return bairro;
	}
	public static void deletarBairro(Integer[] codigos) throws SQLException{
		for(int codigo=0;codigo<codigos.length;codigo++){
			deletarBairro(codigos[codigo]);
		}
	}
	public static void deletarBairro(Integer codigo) throws SQLException{
		conexao.createStatement().execute(Utilidade.instrucaoDelete(TabelaCpSeg.TAB_BAIRRO.getNome(), "TBA_COD_BAIRRO=" + codigo));
	}
	public static Logradouro buscarLogradouro(Integer codigo)throws SQLException{
		Logradouro logradouro=null;
		String sql="SELECT TLG_TBA_COD_BAIRRO,TLG_COD_LOGRADOURO,TLG_NOME,TLG_TTL_COD_TIP_LOGR FROM TAB_LOGRADOURO WHERE TLG_COD_LOGRADOURO=" + codigo;
		ResultSet rs=conexao.createStatement().executeQuery(sql);
		while(rs.next()){
			logradouro=new Logradouro();
			logradouro.setCodigo(rs.getInt("TLG_COD_LOGRADOURO"));
			logradouro.setNome(rs.getString("TLG_NOME"));
			logradouro.setBairro(buscarBairro(rs.getInt("TLG_TBA_COD_BAIRRO")));
			logradouro.setTipoLogradouro(tipoLogradouroGestor(rs.getInt("TLG_TTL_COD_TIP_LOGR")));
		}
		rs.close();
		return logradouro;
	}
	private static Logradouro.TipoLogradouro tipoLogradouroGestor(Integer codigo){
		Logradouro.TipoLogradouro tipo=Logradouro.TipoLogradouro.RUA;
		switch(codigo){
			case 1:
				tipo=Logradouro.TipoLogradouro.FAZENDA;
				break;
			case 2:
				tipo=Logradouro.TipoLogradouro.AVENIDA;
				break;
			case 3:
				tipo=Logradouro.TipoLogradouro.PRACA;
				break;
			case 4:
				tipo=Logradouro.TipoLogradouro.PRAIA;
				break;
			case 5:
				tipo=Logradouro.TipoLogradouro.POVOADO;
				break;
			case 6:
				tipo=Logradouro.TipoLogradouro.ALAMEDA;
				break;
			case 7:
				tipo=Logradouro.TipoLogradouro.CONDOMINIO;
				break;
			case 8:
				tipo=Logradouro.TipoLogradouro.CAMIMHO;
				break;
			case 9:
				tipo=Logradouro.TipoLogradouro.ILHA;
				break;
			case 10:
				tipo=Logradouro.TipoLogradouro.LADEIRA;
				break;
			case 11:
				tipo=Logradouro.TipoLogradouro.LOTEAMENTO;
				break;
			case 12:
				tipo=Logradouro.TipoLogradouro.LUGAREJO;
				break;
			case 13:
				tipo=Logradouro.TipoLogradouro.PARQUE;
				break;
			case 14:
				tipo=Logradouro.TipoLogradouro.PONTE;
				break;
			case 15:
				tipo=Logradouro.TipoLogradouro.VIADUTO;
				break;
			case 16:
				tipo=Logradouro.TipoLogradouro.RODOVIA;
				break;
			case 17:
				tipo=Logradouro.TipoLogradouro.RUA;
				break;
			case 18:
				tipo=Logradouro.TipoLogradouro.TRAVESSA;
				break;
			case 19:
				tipo=Logradouro.TipoLogradouro.ESTRADA;
				break;
			case 20:
				tipo=Logradouro.TipoLogradouro.BECO;
				break;
			case 21:
				tipo=Logradouro.TipoLogradouro.VILA;
				break;
			case 22:
				tipo=Logradouro.TipoLogradouro.CARRASCO;
				break;
			case 23:
				tipo=Logradouro.TipoLogradouro.BR;
				break;
			case 24:
				tipo=Logradouro.TipoLogradouro.CALCADA;
				break;
			case 25:
				tipo=Logradouro.TipoLogradouro.CONJUNTO;
		}
		
		return tipo;
	}
	public static Integer migrarBairros(List<Bairro>bairros,Cidade cidade) throws SQLException{
		PreparedStatement pst=conexao.prepareStatement("INSERT INTO TAB_BAIRRO(TBA_TMU_COD_MUNICIPIO,TBA_COD_BAIRRO,TBA_NOME) values("+ Utilidade.criarInterrogacoes(3) + ")"); 
		
		Integer migrados=0;
		for(Bairro b: bairros){
			pst.setInt(1, b.getCidade().getCodigo());
			pst.setInt(2, b.getCodigo());
			pst.setString(3, b.getNome());
			pst.execute();
			migrados++;
		}
		pst.close();
		return migrados;
		
	}
	public static Integer migrarLogradouros(List<Logradouro>logradouros) throws Exception{
		String sql="INSERT INTO TAB_LOGRADOURO(tlg_tmu_cod_municipio, tlg_tba_cod_bairro, tlg_ttl_cod_tip_logr, " +
		"tlg_cod_logradouro, tlg_secao, tlg_nome, tlg_ttr_cod_trecho, tlg_quadra,tlg_cod_logradouro_inicial, tlg_cod_logradouro_final, " +
		"tlg_cod_bairro_inicial, tlg_cod_bairro_final, tlg_cep) " 
		+
		"values("+ Utilidade.criarInterrogacoes(13) +")";
		Integer migrados=0;
		PreparedStatement pst=conexao.prepareStatement(sql);
		for(Logradouro l: logradouros){
			if(l!=null){
				Bairro b=l.getBairro();
				if(b==null){
					b=buscarBairro(1000);//NAO IDENTIFICADO
				}
				pst.setInt(1, b.getCidade().getCodigo());
				pst.setInt(2, b.getCodigo());
				pst.setInt(3, l.getTipoLogradouro().getValor());
				pst.setInt(4, l.getCodigo());
				pst.setString(5, "0");
				pst.setString(6, l.getNome());
				pst.setInt(7, 0);
				pst.setString(8, "0");
				pst.setString(9, "0");
				pst.setString(10, "0");
				pst.setInt(11, 0);
				pst.setInt(12, 0);
				pst.setString(13, l.getCep());
				pst.execute();
				migrados++;
			}
		}
		pst.close();
		return migrados;
	}
	public static void criarUmTipoLogradouro(Integer codigo,String nome,String sigla) throws SQLException{
		PreparedStatement pst=conexao.prepareStatement("INSERT INTO TAB_TIPO_LOGR(TTL_COD_TIP_LOGR, TTL_NOME, TTL_SIGLA) values("+ Utilidade.criarInterrogacoes(3) + ")"); 
		pst.setInt(1, codigo);
		pst.setString(2, nome);
		pst.setString(3, sigla);
		pst.execute();
	}
	public static void limparBaseCpSeg() throws SQLException{
		for(TabelaCpSeg tab: TabelaCpSeg.values()){
			if (tab.equals(TabelaCpSeg.TAB_ACESSO_USUARIO )|| tab.equals(TabelaCpSeg.TAB_ACESSO_USUARIO_PRIVILEGIO) || tab.equals(TabelaCpSeg.TAB_USUARIO)){
				conexao.createStatement().execute(Utilidade.instrucaoDelete(TabelaCpSeg.TAB_ACESSO_USUARIO.getNome(),"TAU_TUS_COD_USUARIO<>'HENRIQUE' OR TAU_TUS_COD_USUARIO<>'GLEYSON'"));
				conexao.createStatement().execute(Utilidade.instrucaoDelete(TabelaCpSeg.TAB_ACESSO_USUARIO_PRIVILEGIO.getNome(),"TAU_TUS_COD_USUARIO<>'HENRIQUE' OR TAU_TUS_COD_USUARIO<>'GLEYSON'"));
				conexao.createStatement().execute(Utilidade.instrucaoDelete(TabelaCpSeg.TAB_USUARIO.getNome(),"TUS_COD_USUARIO<>'HENRIQUE'"));
			}else{
				conexao.createStatement().execute(Utilidade.instrucaoDelete(tab.getNome()));
			}
				
		}
	}
	
}

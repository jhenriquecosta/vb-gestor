package com.bcptecnologias.migracao.gestortributario;

import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import com.bcptecnologias.migracao.gestortributario.fachada.ControladorCodo;
import com.bcptecnologias.migracao.gestortributario.fachada.ControladorCpSeg;
import com.bcptecnologias.migracao.gestortributario.fachada.ControladorCpTrib;
import com.bcptecnologias.migracao.gestortributario.negocio.Bairro;
import com.bcptecnologias.migracao.gestortributario.negocio.Cidade;
import com.bcptecnologias.migracao.gestortributario.negocio.Contribuinte;
import com.bcptecnologias.migracao.gestortributario.negocio.Imovel;
import com.bcptecnologias.migracao.gestortributario.negocio.Logradouro;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.DetalheImovel;
import com.bcptecnologias.migracao.gestortributario.negocio.gestor.GrupoComponente;
import com.bcptecnologias.migracao.gestortributario.util.Conexao;
import com.bcptecnologias.migracao.gestortributario.util.Utilidade;

public class Migrador {
	public static void main (String [] args){
		try{
			for(int ano=2011;ano<2030;ano++){
				for(int mes=1;mes<13;mes++){
					String sql="INSERT INTO TAB_MONETARIA(TMO_DATA, TMO_VALOR, TMO_UNIDADE, TMO_VALOR_MOSTRA, TMO_ANO, TMO_MES, TMO_PERIODO, TMO_VALOR_DEBITO)" +
					" VALUES ('2011-07-05', 1, 2, NULL," + ano + "," + mes +"," + ano+ Utilidade.formatar("00", mes) + ",0);";
					System.out.println(sql);
				}
			}
			/*
			 * 001931040, ACESSORIO ESCRITURAL 4316876, AG:791-9 CC:1667-5 
			for(int ano=2011;ano<2030;ano++){
				for(int mes=1;mes<13;mes++){
					String sql="INSERT INTO TAB_MONETARIA(TMO_DATA, TMO_VALOR, TMO_UNIDADE, TMO_VALOR_MOSTRA, TMO_ANO, TMO_MES, TMO_PERIODO, TMO_VALOR_DEBITO)" +
					" VALUES ('2011-07-05', 1, 2, NULL," + ano + "," + mes +"," + ano+ Utilidade.formatar("00", mes) + ",0);";
					System.out.println(sql);
				}
			}
			*/
			//CRIANDO CONEXOES GESTOR
			Connection conexaoCPSEG=Conexao.getConexaoSQL("cpseg");
			Connection conexaoCPTRIB=Conexao.getConexaoSQL("cptrib");
			//SETANDO AS CONEXOES 
			ControladorCpSeg.setConexao(conexaoCPSEG);
			ControladorCpTrib.setConexao(conexaoCPTRIB);
			
			//procedimentoCodo();
			
			
		}catch(Exception ex){
			ex.printStackTrace();
		}finally {
			System.exit(0);
		}
	}
	private static void procedimentoD2Ti() throws Exception{
		
		System.out.println("MIGRACAO CONCLUIDA");
	}
	private static void procedimentoCodo() throws Exception{
		//01010010038002 e 01010010066002
		Connection conexaoCodo=Conexao.getConexaoFireBird();
		ControladorCodo.setConexaoCodo(conexaoCodo);
		//ControladorCpSeg.limparBaseCpSeg();
		//ControladorCpTrib.limparBaseCPTRIB();
		//ControladorCpTrib.resetaNumCorrelativa();
		Cidade cidade=ControladorCpSeg.buscarCidade(1179);
		//faseBairro(cidade);
		//faseLogradouro();
		//faseContribuinte(cidade);
		//ESSA FASE DEPOIS QUE RODAR A ATUALIZACAO DAS INSCRICOES PELO GESTOR
		faseImovel();
		System.out.println("MIGRACAO CONCLUIDA");
	}
	private static void faseBairro(Cidade cidade) throws Exception{
		List<Bairro>bairros=ControladorCodo.listarBairros(cidade);
		System.out.println(bairros.size() + " BAIRROS - ENCONTRADOS");
		System.out.println(ControladorCpSeg.migrarBairros(bairros,cidade) + " BAIRROS - IMPORTADOS COM SUCESSO");
		
		Bairro bairro=new Bairro();
		bairro.setCidade(bairros.get(0).getCidade());
		bairro.setCodigo(1000);
		bairro.setNome("NAO IDENTIDICADO");
		bairros=new ArrayList();
		bairros.add(bairro);
		ControladorCpSeg.migrarBairros(bairros,cidade);
		
		Integer[] codBairros={274,256,258,265,269};
		ControladorCpSeg.deletarBairro(codBairros);
		
	}
	private static void faseLogradouro()throws Exception{
		List<Logradouro>logradouros=ControladorCodo.listarLogradouros();
		System.out.println(logradouros.size() + " LOGRADOUROS - ENCONTRADOS");
		System.out.println(ControladorCpSeg.migrarLogradouros(logradouros)+ " LOGRADOUROS - IMPORTADOS COM SUCESSO");
		//ControladorCpSeg.criarUmTipoLogradouro(25, "CONJUNTO", "CNJ");
	}
	private static void faseContribuinte(Cidade cidade)throws Exception{
		List<Contribuinte>contribuintes=ControladorCodo.listarContribuintes(cidade);
		System.out.println(contribuintes.size() + " CONTRIBUINTES - ENCONTRADOS");
		System.out.println(ControladorCpTrib.migrarContribuintes(contribuintes)+ " CONTRIBUINTES - IMPORTADOS COM SUCESSO");
	}
	private static void faseImovel()throws Exception{
		List<Imovel>imoveis=ControladorCodo.listarImoveis();
		System.out.println(imoveis.size() + " IMOVEIS - ENCONTRADOS");
		System.out.println(ControladorCpTrib.migrarImoveis(imoveis) + " IMOVEIS - IMPORTADOS COM SUCESSO");
		
		List<DetalheImovel>detalhesImovel=ControladorCodo.listarDetalhesImovel(imoveis);
		System.out.println(detalhesImovel.size() + " DETALHES DO IMOVEL - ENCONTRADOS");
		System.out.println(ControladorCpTrib.migrarDetalheImoveis(detalhesImovel) + " DETALHES DO IMOVEL - IMPORTADOS COM SUCESSO");
	}
}

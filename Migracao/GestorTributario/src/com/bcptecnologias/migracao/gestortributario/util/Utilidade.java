package com.bcptecnologias.migracao.gestortributario.util;

import java.util.Date;

import com.bcptecnologias.migracao.gestortributario.negocio.gestor.NumeroCorrelativo;

public final class Utilidade {
	public static Date dataAtual(){
		Date dataAtual = new Date(System.currentTimeMillis());
		return dataAtual;
	}
	public static String gerarInscricaoPeloCodigoExistente(String codigoExistente){
		String codigoFormatado="";
		if(codigoExistente.trim().length()>0){
			char[] numeros=new Integer(codigoExistente).toString().toCharArray();
			for(int n=0;n<numeros.length;n++){
				if((n+2)==numeros.length){
					codigoFormatado=codigoFormatado+"-";
					codigoFormatado=codigoFormatado + numeros[n];
				}else
					codigoFormatado=codigoFormatado + numeros[n];
			}
		}
		return codigoFormatado;
	}
	
	
	public static String instrucaoDelete(String nomeTabela){
		return instrucaoDelete(nomeTabela,null);
	}
	public static String instrucaoDelete(String nomeTabela, String condicao){
		String insrucao= "DELETE FROM "+ nomeTabela ;
		if(condicao==null){
			System.out.println("DELETANDO A TABELA --- " + nomeTabela);
			return  insrucao + ";";
			
		}else{
			System.out.println("DELETANDO A TABELA --- " + nomeTabela);
			return  insrucao + " WHERE " + condicao + ";";
		}
		
	}
	public static String criarInterrogacoes(Integer colunas){
		StringBuffer sb=new StringBuffer("");
		for(int c=1;c<=colunas;c++){
			sb.append("?,");
		}
		return sb.toString().substring(0, sb.toString().trim().length()-1);
	}
	public static String formatar(String formato, Object valor){
		int tamanho=valor.toString().length();
		formato=formato.substring(0,(formato.length()-tamanho));
		formato=formato + valor.toString();
		return formato;
		
	}
	
}
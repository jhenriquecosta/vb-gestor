package com.bcptecnologias.migracao.gestortributario;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;

public class GetLet {
	public static void main(String [] args){
		File f=new File("c:\\campos.txt");
		FileReader fr=null;
		try {
			fr = new FileReader(f);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		BufferedReader leitor = new BufferedReader(fr);
		String campo=null;
		int l=1;
		try {
			while(leitor.ready()){
				if(l==5){
					l=1;
				}
				campo=leitor.readLine();
				System.out.println(metotoGet(campo));
				System.out.println(metotoLet(campo));
				l++;
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	private static String metotoGet(String campo){
		StringBuffer sb=new StringBuffer();
		sb.append("Public Property Get " + campo + " () As String \n");
		sb.append("     "+campo+ "=m"+campo + "\n");
	    sb.append("End Property"); 
		return sb.toString();
	}
	private static String metotoLet(String campo){
		StringBuffer sb=new StringBuffer();
		sb.append("Public Property Let " + campo + "(ByVal valor As String)\n");
		sb.append("     m"+campo+ "=valor\n");
	    sb.append("End Property"); 
		return sb.toString();
	}
}

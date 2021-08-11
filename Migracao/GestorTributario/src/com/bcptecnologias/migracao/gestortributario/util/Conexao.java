package com.bcptecnologias.migracao.gestortributario.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
public class Conexao {
	public static Connection getConexaoFireBird() throws ClassNotFoundException, SQLException{
		//
		Connection conexao=null;
		Class.forName("org.firebirdsql.jdbc.FBDriver");
		//String endereco="C:\\firebd\\BACKUP.GDB";
		//endereco=endereco.replace("/", "\'");
		conexao= DriverManager.getConnection("jdbc:firebirdsql:localhost/3050:C:/firebd/tributos.gdb", "SYSDBA", "masterkey"); 
		return conexao;
	}
	public static Connection getConexaoSQL(String banco) throws ClassNotFoundException, SQLException{
		Connection conexao=null;
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		conexao= DriverManager.getConnection("jdbc:sqlserver://localhost:1433;databaseName="+ banco+";selectMethod=cursor","sa","kabecao");
		return conexao;
	}
	
}

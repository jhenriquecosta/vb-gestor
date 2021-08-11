using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;

namespace VSClass
{
	public class VSConexao
	{

		//=========================================================

		// VBto upgrade warning: Conn As Connection	OnWrite(Connection, int)
		 private Connection Conn = new Connection();
		 public string Dsn;
		 public string User;
		 public string Password;
		 public string Catalog;
		 public TipoBanco FormatoBanco;
		enum TipoBanco {
			Access = 0,
			SQLServer = 1,
			oracle = 2,
			Interbase = 3,
			Postgres = 4
		};

		public void /* Connection */ DBConnection
		{
			get
			{
				if (CMOD.bRegistrado) DBConnection = Conn;
			}
		}

		public void DBConnection( /* Connection NewDBConnection */ )
		{
			if (CMOD.bRegistrado) Conn = NewDBConnection;
		}

		public void BeginTrans()
		{
			if (CMOD.bRegistrado) Conn.BeginTrans();
		}

		public void CommitTrans()
		{
			if (CMOD.bRegistrado) Conn.CommitTrans();
		}

		public void RollbackTrans()
		{
			if (CMOD.bRegistrado) Conn.RollbackTrans();
		}

		public void Abrir(string ConnStr)
		{
			if (CMOD.bRegistrado) {
				Conn.Open(ConnStr);
				Conn.CommandTimeout = 0;
			}
		}

		public void /* ADODB.Errors */ Errors
		{
			get
			{
				if (CMOD.bRegistrado) Errors = Conn.Errors;
			}
		}

		public void Execute(string Str)
		{
			if (CMOD.bRegistrado) Conn.Execute(Str);
		}

		public int State
		{
			get
			{
				int State = 0;
				if (CMOD.bRegistrado) State = Conn.State;
				return State;
			}
		}

		public void Fechar()
		{
			if (CMOD.bRegistrado) Conn.Close();
		}

		public string ConnectionString
		{
			get
			{
				string ConnectionString = "";
				if (CMOD.bRegistrado) ConnectionString = Conn.ConnectionString;
				return ConnectionString;
			}
		}

		private VSConexao() : base()
		{
			CMOD.ValidaComponente("CLASS");
			if (CMOD.bRegistrado) {
				Conn = new Connection();
			}
		}

		~VSConexao()
		{
			if (CMOD.bRegistrado) Conn = null;
		}

	}
}
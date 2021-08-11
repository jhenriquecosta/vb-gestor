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
	public class VSDados
	{

		//=========================================================
		 public bool ModoTexto; // SQz (22/09/02) : vide descricao da propriedade
		// VBto upgrade warning: Conexao As VSClass.VSConexao	OnWrite(int)
		 public VSClass.VSConexao Conexao = new VSClass.VSConexao();
		// VBto upgrade warning: Tabela As VSRecordset	OnWrite(int)
		 public VSRecordset Tabela = new VSRecordset();
		 private VSUtil Util = new VSUtil();
		 private VSTexto Edita = new VSTexto();
		 private string strUltimaQuery;
		 private string cteSeparador;

		enum TipoConversao {
			TCBit,
			TCSimples,
			TCInteiro,
			TCByte,
			TCMonetario,
			TCDuplo,
			TCDataHora,
			tctexto,
			TCMemo,
			TCLob,
			TCBinario
		};
		enum TipoCursor {
			SomenteAvanco = 0,
			Registros = 1,
			Dinamico = 2,
			Estatico = 3
		};
		enum TipoTrava {
			SomenteLeitura = 1,
			Pessimista = 2,
			Otimista = 3,
			OtimistaBatch = 3
		};
		enum TipoParteTexto {
			LeftVs = 1,
			MidVs = 2,
			RightVs = 3
		};

		public void AbreTrans()
		{
			if (CMOD.bRegistrado) Conexao.BeginTrans();
		}

		public void GravaTrans()
		{
			if (CMOD.bRegistrado) Conexao.CommitTrans();
		}

		public void CancelaTrans()
		{
			if (CMOD.bRegistrado) Conexao.RollbackTrans();
		}

		public bool AbreBanco( /* TipoBanco Tipo */ string DataSource, string User)
		{
			return AbreBanco( /* Tipo */ DataSource, User, "");
		}
		public bool AbreBanco( /* TipoBanco Tipo */ string DataSource, string User, string Password)
		{
			return AbreBanco( /* Tipo */ DataSource, User, Password, "");
		}
		public bool AbreBanco( /* TipoBanco Tipo */ string DataSource, string User, string Password, string Parametro_Opcional)
		{
			bool AbreBanco = false;
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Str = "";

					if (Tipo==Access)
					{

						Str = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=True;Password=''"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Jet OLEDB:Database Password="+Password;

					}
					else if (Tipo==SQLServer)
					{
						Str = "Provider=SQLOLEDB.1;Persist Security Info=True"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Password="+Password;
						if (Parametro_Opcional!="") Str = Str+";Initial Catalog="+Parametro_Opcional;

					}
					else if (Tipo==oracle)
					{
						Str = "Provider=MSDAORA.1;Persist Security Info=True;"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Password="+Password;
					}
					else if (Tipo==Interbase)
					{
						Str = "Provider=IbOleDb.1;Persist Security Info=True;Mode=ReadWrite"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Password="+Password;
						if (Parametro_Opcional!="") Str = Str+";Loation="+Parametro_Opcional;

					}
					FechaBanco();
					Conexao = new VSConexao();
					Conexao.Abrir(Str);

					Conexao.Dsn = DataSource;
					Conexao.User = User;
					Conexao.Password = Password;
					Conexao.Catalog = Parametro_Opcional;
					Conexao.FormatoBanco = Tipo;

					AbreBanco = true;
				}
				return AbreBanco;
			}
			catch
			{	// Trata:
				if (Conexao.Errors.Count>0) {
					AbreBanco = false;
					Util.Avisa("Erro: "+Conexao.Errors(0).Number+" - "+Conexao.Errors(0).Description+".");
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return AbreBanco;
		}

		public bool AbreBancoDados(string Banco)
		{
			bool AbreBancoDados = false;
			try
			{	// On Error GoTo Trata
				 TipoBanco Tipo = new TipoBanco(); string DataSource = "",  User = "",  Password = "",  Parametro_Opcional = "";
				if (CMOD.bRegistrado) {
					 string Str = "";
					 VSInstala Instala = new VSInstala();

					Tipo = Instala.PegaConfiguracao(Banco, tcTipo);
					DataSource = Instala.PegaConfiguracao(Banco, tcDsn);
					User = Instala.PegaConfiguracao(Banco, tcUser);
					Parametro_Opcional = Instala.PegaConfiguracao(Banco, tcCatalog);
					Password = Instala.PegaConfiguracao(Banco, tcPassword);

					if (Tipo==Access)
					{

						Str = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=True;Password=''"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Jet OLEDB:Database Password="+Password;

					}
					else if (Tipo==SQLServer)
					{
						Str = "Provider=SQLOLEDB.1;Persist Security Info=True"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Password="+Password;
						if (Parametro_Opcional!="") Str = Str+";Initial Catalog="+Parametro_Opcional;

					}
					else if (Tipo==oracle)
					{
						Str = "Provider=MSDAORA.1;Persist Security Info=True;"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Password="+Password;
					}
					else if (Tipo==Interbase)
					{
						Str = "Provider=IbOleDb.1;Persist Security Info=True;Mode=ReadWrite"+";Data Source="+DataSource+";User ID="+User;

						if (Password!="") Str = Str+";Password="+Password;
						if (Parametro_Opcional!="") Str = Str+";Loation="+Parametro_Opcional;
					}
					else if (Tipo==4)
					{
						Str = "PROVIDER=PostgreSQL OLE DB Provider;DATA SOURCE="+DataSource+"; USER ID= "+User+"; PASSWORD="+Password+";";
					}
					FechaBanco();
					Conexao = new VSConexao();
					Conexao.Abrir(Str);

					Conexao.Dsn = DataSource;
					Conexao.User = User;
					Conexao.Password = Password;
					Conexao.Catalog = Parametro_Opcional;
					Conexao.FormatoBanco = Tipo;

					AbreBancoDados = true;
				}
				return AbreBancoDados;
			}
			catch
			{	// Trata:
				if (Conexao.Errors.Count>0) {
					AbreBancoDados = false;
					Util.Avisa("Erro: "+Conexao.Errors(0).Number+" - "+Conexao.Errors(0).Description+".");
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return AbreBancoDados;
		}

		public bool Executa(string Sql)
		{
			bool Executa = false;
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					if (!ModoTexto) {
						Conexao.Execute(Sql);
					}

					Executa = true;
					strUltimaQuery = Sql;
				}
				return Executa;
			}
			catch
			{	// Trata:
				if (Conexao.Errors.Count>0) {
					Executa = false;
					Util.Erro("Erro: "+Conexao.Errors(0).Number+" - "+Conexao.Errors(0).Description+".");
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return Executa;
		}

		public bool AbreTabela(string Sql)
		{
			return AbreTabela(Sql, null);
		}
		public bool AbreTabela(string Sql, ref object Record)
		{
			return AbreTabela(Sql, Record, TipoCursor.SomenteAvanco);
		}
		public bool AbreTabela(string Sql, ref object Record, TipoCursor TipoAbertura)
		{
			return AbreTabela(Sql, Record, TipoAbertura, TipoTrava.SomenteLeitura);
		}
		public bool AbreTabela(string Sql, ref object Record, TipoCursor TipoAbertura, TipoTrava TipoTrava)
		{
			bool AbreTabela = false;
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					if (!IsMissing(Record)) {
						FechaTabela(ref Record);
						Record = new VSRecordset();
						Record.Abrir(Sql, Conexao, TipoAbertura, TipoTrava);
						if (!Record.Eof) AbreTabela = true;
					} else {
						FechaTabela(ref Tabela);
						Tabela = new VSRecordset();
						Tabela.Abrir(Sql, Conexao, TipoAbertura, TipoTrava);
						if (!Tabela.Eof) AbreTabela = true;
					}
					strUltimaQuery = Sql;
				}
				return AbreTabela;

			}
			catch
			{	// Trata:
				if (Conexao.Errors.Count!=0) {
					AbreTabela = false;
					Util.Avisa("Erro: "+Conexao.Errors(0).Number+" - "+Conexao.Errors(0).Description+".");
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return AbreTabela;
		}

		public void ApagaTabela(string Tabela)
		{
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					Executa("DROP TABLE "+Tabela);
					strUltimaQuery = "DROP TABLE "+Tabela;
				}
				return;

			}
			catch
			{	// Trata:
				if (Conexao.Errors.Count>0) {
					Util.Avisa("Erro: "+Conexao.Errors(0).Number+" - "+Conexao.Errors(0).Description+".");
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
		}

		public bool InsereDados(string Tabela, string Valores)
		{
			return InsereDados(Tabela, Valores, "");
		}
		public bool InsereDados(string Tabela, string Valores, string Campos)
		{
			bool InsereDados = false;
			// VSClass.VSDados.Function InsereDados
			// ================================================================================
			// Queiroz em VTDES_01
			// 24/05/2002-12:11:28
			// 
			// Descricao  : Gera uma sql do tipo Insert, e envia
			// 
			// Parametros : Tabela (String)
			// Valores (String)
			// Campos (String)
			// 
			// Ex:
			// --------------------------------------------------------------------------------
			 string Sql = "";
			 int i;
			 string Valor = "";
			 string ListaDeValores = "";
			 int PosCaracter;
			 string ValorData = "";
			// VBto upgrade warning: X As int	OnWrite(int)
			 int X;
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					Sql = "INSERT into "+Tabela+" "+(Strings.Trim(Campos)=="" ? "" : "("+Campos+")");
					Sql = Sql+" VALUES(";

					i = 1;
					X = Valores.Length;
					do {
						Valor = Util.ParseString(Valores, cteSeparador, i);
						// If (UCase(Left(Trim(Valor), 8)) = "CONVERT(" And Right(Trim(Valor), 1) = ")") Then
						if ((Conexao.FormatoBanco==SQLServer && (Strings.UCase(Strings.Left(Strings.Trim(Valor), 8))=="CONVERT(" && Strings.Right(Strings.Trim(Valor), 1)==")")) || (Conexao.FormatoBanco==oracle && (Strings.UCase(Strings.Left(Strings.Trim(Valor), 3))=="TO_" && Strings.Right(Strings.Trim(Valor), 1)==")"))) {
							Sql = Sql+(i==1 ? "" : ",")+(Edita.PosPic(Valor, "'")==0 ? Strings.Left(Strings.Left(Valor, Edita.PosPic(Valor, ","))+"'"+Strings.Mid(Valor, Edita.PosPic(Valor, ",")+1), Valor.Length)+"')" : Valor);
						} else {
							Sql = Sql+(i==1 ? "" : ",")+FormataValorCampo(Valor);
						}
						i += 1;
						X -= (Valor.Length+cteSeparador.Length);
					} while (!(X==0));
					Sql = Sql+")";

					InsereDados = Executa(Sql);
					strUltimaQuery = Sql;
				}
				return InsereDados;
			}
			catch
			{	// Trata:
				if (Information.Err().Number==6) {
					Resume Next;
				}
			}
			return InsereDados;
		}

		public bool DeletaDados(string Tabela)
		{
			return DeletaDados(Tabela, "");
		}
		public bool DeletaDados(string Tabela, string Condicao)
		{
			bool DeletaDados = false;
			 string Sql = "";

			if (CMOD.bRegistrado) {
				Sql = "DELETE  from "+Tabela+" ";
				if (Strings.Trim(Condicao)!="") Sql = Sql+" WHERE "+Condicao;

				DeletaDados = Executa(Sql);
				strUltimaQuery = Sql;
			}
			return DeletaDados;
		}

		public bool AtualizaDados(string Tabela, string Valores, string Campos)
		{
			return AtualizaDados(Tabela, Valores, Campos, "");
		}
		public bool AtualizaDados(string Tabela, string Valores, string Campos, string Condicao)
		{
			bool AtualizaDados = false;
			// VSClass.VSDados.Function AtualizaDados
			// ================================================================================
			// Queiroz em VTDES_01
			// 24/05/2002-14:34:11
			// 
			// Descricao  : Prepara uma sql do tipo UPDATE e envia para o banco. O valor de retorno
			// indica sucesso ou fracasso da execucao.
			// 
			// Parametros : Tabela (String) - Nome da tabela
			// Valores (String) - Valores a serem atualizados (preparada pela PreparaValores)
			// Campos (String) - Nomes dos campos, separados por virgula
			// Condicao (String) - Condicao de atualizacao (clausula Where)
			// 
			// Ex:
			// --------------------------------------------------------------------------------
			try
			{	// On Error GoTo ErroAtualiza
				 string Sql = "";
				 int i;
				 string Campo = "",  Valor = "";
				 string ListaDeValores = "";
				 int PosCaracter;
				 string ValorData = "";

				if (CMOD.bRegistrado) {
					Sql = "UPDATE "+Tabela+" SET ";

					i = 1;
					do {
						Campo = Util.ParseString(Campos, ",", i);
						if (Campo!="") {
							Sql = Sql+(i==1 ? "" : ",")+Campo+" = ";
							Valor = Util.ParseString(Valores, cteSeparador, i);
							if (Strings.UCase(Strings.Left(Strings.Trim(Valor), 8))=="CONVERT(" && Strings.Right(Strings.Trim(Valor), 1)==")") {
								Sql = Sql+(Edita.PosPic(Valor, "'")==0 ? Strings.Left(Strings.Left(Valor, Edita.PosPic(Valor, ","))+"'"+Strings.Mid(Valor, Edita.PosPic(Valor, ",")+1), Valor.Length)+"')" : Valor);
							} else {
								Sql = Sql+FormataValorCampo(Valor);
							}
							i += 1;
						} else {
							break;
						}
					} while (!(Campo==""));

					if (Condicao!="") Sql = Sql+" WHERE "+Condicao;

					AtualizaDados = Executa(Sql);
					strUltimaQuery = Sql;
				}
				return AtualizaDados;

			}
			catch
			{	// ErroAtualiza:
				MessageBox.Show(Information.Err().Description);
			}
			return AtualizaDados;
		}

		public bool GravaDados(string Tabela, string Valores, string Campos, string Condicao)
		{
			bool GravaDados = false;
			 VSRecordset RSTEMP = new VSRecordset();
			if (CMOD.bRegistrado) {
				if (AbreTabela("SELECT * FROM "+Tabela+" WHERE "+Condicao, ref RSTEMP)) {
					GravaDados = AtualizaDados(Tabela, Valores, Campos, Condicao);
				} else {
					GravaDados = InsereDados(Tabela, Valores, Campos);
				}
				FechaTabela(ref RSTEMP);
			}
			return GravaDados;
		}

		// VBto upgrade warning: Record As object	OnWrite(VSRecordset, int)
		public void FechaTabela()
		{
			FechaTabela(null);
		}
		public void FechaTabela(ref object Record)
		{
			if (CMOD.bRegistrado) {
				if (!IsMissing(Record)) {
					if (!Record==null) {
						if (Record.State!=adStateClosed) {
							Record.Fechar();
							Record = null;
						}
					}
				} else {
					if (!Tabela==null) {
						if (Tabela.State!=adStateClosed) {
							Tabela.Fechar();
							Tabela = null;
						}
					}
				}
			}
		}

		public object PreparaValor()
		{
			object PreparaValor = 0; Valores(()); /*? ) As */ String();
			// VBto upgrade warning: i As byte	OnWrite(byte, int)
			 byte i,  Min,  Max;
			 string Valor = "";

			Valor = "";
			if (CMOD.bRegistrado) {
				Min = LBound(Valores);
				Max = UBound(Valores);
				for(i=Min; i<=Max; i++) {
					Valor = Valor+Util.Nvl((Valores(i)).ToString(), "Null")+cteSeparador;
					// If CStr(Valores(i)) <> "" Then
					// Valor = Valor & Seg.Criptografa(Seg.Criptografa(Seg.Criptografa(CStr(Valores(i))))) & cteSeparador
					// Else
					// Valor = Valor & "Null" & cteSeparador
					// End If
				}
				PreparaValor = Valor;
			}
			return PreparaValor;
		}

		public void FechaBanco()
		{
			if (CMOD.bRegistrado) {
				if (!Conexao==null) {
					if (Conexao.State!=adStateClosed) {
						Conexao.Fechar();
						Conexao = null;
					}
				}
			}
		}

		public string BuscaCodigo(string Tabela)
		{
			string BuscaCodigo = "";
			 VSRecordset RS = new VSRecordset();

			if (CMOD.bRegistrado) {
				if (AbreTabela(Tabela, ref RS)) {
					BuscaCodigo = (IsNull(RS(0)) ? "0" : RS(0));
				} else {
					BuscaCodigo = "0";
				}
				FechaTabela(ref RS);
				strUltimaQuery = Tabela;
			}
			return BuscaCodigo;
		}

		public string DescricaoGeral(string Tabela, int Codigo)
		{
			string DescricaoGeral = "";
			 VSRecordset RS = new VSRecordset();
			 string strSql = "";

			if (CMOD.bRegistrado) {
				strSql = "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO="+Codigo+" AND TGE_TIPO="+"(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='"+Tabela+"')";

				if (AbreTabela(strSql, ref RS)) {
					DescricaoGeral = RS!TGE_NOME;
				}
				FechaTabela(ref RS);
				strUltimaQuery = strSql;
			}
			return DescricaoGeral;
		}

		public int CodigoGeral(string Tabela, string Descricao)
		{
			int CodigoGeral = 0;
			 VSRecordset RS = new VSRecordset();
			 string strSql = "";

			if (CMOD.bRegistrado) {
				strSql = "SELECT TGE_CODIGO FROM TAB_GERAL WHERE TGE_NOME='"+Descricao+"' AND TGE_TIPO="+"(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='"+Tabela+"')";

				if (AbreTabela(strSql, ref RS)) {
					CodigoGeral = RS!TGE_CODIGO;
				}
				FechaTabela(ref RS);
				strUltimaQuery = strSql;
			}
			return CodigoGeral;
		}

		private VSDados() : base()
		{
			CMOD.ValidaComponente("CLASS");
			if (CMOD.bRegistrado) {
				Conexao = new VSConexao();
				cteSeparador = "VTSEP";
			}
		}

		public string Concatena()
		{
			string Concatena = "";
			if (CMOD.bRegistrado) {
				VBtoVar = Conexao.FormatoBanco;
				if (VBtoVar==Access)
				{

					Concatena = " & ";
				}
				else if (VBtoVar==SQLServer)
				{
					Concatena = " + ";
				}
				else if (VBtoVar==oracle)
				{
					Concatena = " || ";
				}
				else if (VBtoVar==Interbase)
				{
					Concatena = " || ";
				}
			}
			return Concatena;
		}

		// VBto upgrade warning: Dado As object --> As string
		public string Converte(string Dado, TipoConversao Tipo)
		{
			string Converte = "";
			if (CMOD.bRegistrado) {
				 VSTexto T = new VSTexto();
				 int pos;
				VBtoVar = Conexao.FormatoBanco;
				if (VBtoVar==Access)
				{

					if (Tipo==TipoConversao.TCDataHora)
					{

						if (Strings.InStr(1, Dado, ",103", CompareMethod.Text)>0) {
							Dado = Strings.Mid(Dado, 1, Dado.Length-4);
						}
						if (IsDate(Dado)) {
							Converte = " cdate('"+Dado+"') ";
						} else {
							Converte = " cdate("+Dado+") ";
						}
					}
					else if ((Tipo==TipoConversao.TCLob) || (Tipo==TipoConversao.TCMemo) || (Tipo==TipoConversao.tctexto))
					{
						if (Strings.Trim(Dado)!="") Converte = " CSTR('"+Dado+"') ";
					}
					else if (Tipo==TipoConversao.TCMonetario)
					{
						Converte = " ccur('"+Dado+"') ";
					}
					else if (Tipo==TipoConversao.TCByte)
					{
						Converte = " cbyte('"+Dado+"') ";
					}
					else if (Tipo==TipoConversao.TCInteiro)
					{
						Converte = " cint('"+Dado+"') ";
					}
					else if ((Tipo==TipoConversao.TCDuplo) || (Tipo==TipoConversao.TCSimples))
					{
						Converte = " cdbl('"+Dado+"') ";
						// Case TCDuplo, TCSimples: Dado = Format(Dado, "#,##0.00"): Converte = " cdbl(" & Dado & ") "
					}
					else 
					{
						Util.Erro("Tipo de dados não programado.");
					}
				}
				else if (VBtoVar==SQLServer)
				{
					if (Information.IsNumeric(Dado)) {
						 string ValorAux = "";
						ValorAux = Edita.TiraPic((Dado).ToString(), ".");
						ValorAux = Edita.TrocaPic((ValorAux).ToString(), ",", ".");
						Converte = " convert("+NomeConvSQL(Tipo)+","+(Tipo==TipoConversao.tctexto ? "'"+ValorAux+"'" : ValorAux)+")";
					} else {
						// If Trim(Dado) = "" Then Dado = "Null"
						if (Tipo==TipoConversao.tctexto) {
							if (Strings.Trim(Dado)!="") Converte = " convert("+NomeConvSQL(Tipo)+"("+Dado.Length+"),'"+Dado+"')";
						} else {
							if (Strings.Trim(Dado)!="") Converte = " convert("+NomeConvSQL(Tipo)+","+(Tipo==TipoConversao.TCDataHora ? " '"+Dado+"',103" : Dado)+")";
						}
					}
				}
				else if (VBtoVar==oracle)
				{
					if (Tipo==TipoConversao.TCDataHora) {
						Converte = " to_date('"+Dado+"','dd/mm/yyyy') ";
					} else if (Tipo==TipoConversao.tctexto) {
						Converte = " to_char('"+Dado+"') ";
					} else if (Tipo==TipoConversao.TCMemo) {
						Converte = " to_long('"+Dado+"') ";
					} else {
						if (Tipo==TipoConversao.TCMonetario) {
							Dado = (Dado).ToString("#,##0.00");
						}
						if (VBtoConverter.Fix(Dado)>0) {
							pos = Strings.InStr(Dado, ",", CompareMethod.Text);
							Dado = T.TiraTudo((Dado).ToString());
							if (pos>0) {
								Dado = Strings.Left(Dado, Dado.Length-2)+","+Strings.Right(Dado, 2);
							}
						}
						if (Tipo==TipoConversao.TCMonetario) {
							Converte = " to_number('"+T.TrocaPic((Dado).ToString(), ",", ".")+"','9999999999.99')";
						} else {
							Converte = " to_number('"+T.TrocaPic((Dado).ToString(), ",", ".")+"','9999999999.9999')";
						}
					}
				}
				else if (VBtoVar==Interbase)
				{
					if (Tipo==TipoConversao.TCDataHora) {
						Converte = " to_date("+Dado+",'dd/mm/yyyy') ";
					} else if (Tipo==TipoConversao.tctexto) {
						Converte = " to_char("+Dado+") ";
					} else if (Tipo==TipoConversao.TCMemo) {
						Converte = " to_long("+Dado+") ";
					} else {
						Converte = " to_number("+Dado+") ";
					}
				}
			}
			return Converte;
		}

		private string NomeConvSQL(TipoConversao EnumConv)
		{
			string NomeConvSQL = "";
			if (CMOD.bRegistrado) {
				if (EnumConv==TipoConversao.TCBit)
				{

					NomeConvSQL = "Bit";
				}
				else if (EnumConv==TipoConversao.TCSimples)
				{
					NomeConvSQL = "int";
				}
				else if (EnumConv==TipoConversao.TCInteiro)
				{
					NomeConvSQL = "SmallInt";
				}
				else if (EnumConv==TipoConversao.TCByte)
				{
					NomeConvSQL = "tinyint";
				}
				else if (EnumConv==TipoConversao.TCMonetario)
				{
					NomeConvSQL = "Money";
				}
				else if (EnumConv==TipoConversao.TCDuplo)
				{
					NomeConvSQL = "Real";
				}
				else if (EnumConv==TipoConversao.TCDataHora)
				{
					NomeConvSQL = "DateTime";
				}
				else if (EnumConv==TipoConversao.tctexto)
				{
					NomeConvSQL = "Char";
				}
				else if (EnumConv==TipoConversao.TCMemo)
				{
					NomeConvSQL = "VarChar";
				}
				else if (EnumConv==TipoConversao.TCLob)
				{
					NomeConvSQL = "Text";
				}
				else if (EnumConv==TipoConversao.TCBinario)
				{
					NomeConvSQL = "VarBinary";
				}
			}
			return NomeConvSQL;
		}

		public string ParteTexto(object Dado, TipoParteTexto Parte)
		{
			return ParteTexto(Dado, Parte, 1);
		}
		public string ParteTexto(object Dado, TipoParteTexto Parte, int Inicio)
		{
			return ParteTexto(Dado, Parte, Inicio, 0);
		}
		public string ParteTexto(object Dado, TipoParteTexto Parte, int Inicio, int Tamanho)
		{
			return ParteTexto(Dado, Parte, Inicio, Tamanho, false);
		}
		public string ParteTexto(object Dado, TipoParteTexto Parte, int Inicio, int Tamanho, bool DadoEhCampo)
		{
			string ParteTexto = "";
			if (CMOD.bRegistrado) {
				VBtoVar = Conexao.FormatoBanco;
				if (VBtoVar==Access)
				{


					if (Parte==TipoParteTexto.LeftVs) {
						ParteTexto = Convert.ToString(" Left("+Dado+","+Tamanho+")");
					} else if (Parte==TipoParteTexto.MidVs) {
						ParteTexto = " Mid("+Dado+","+Inicio+(Tamanho==0 ? ")" : ","+Tamanho+")");
					} else if (Parte==TipoParteTexto.RightVs) {
						ParteTexto = Convert.ToString(" Right("+Dado+","+Tamanho+")");
					}
				}
				else if (VBtoVar==SQLServer)
				{
					if (Parte==TipoParteTexto.LeftVs) {
						ParteTexto = Convert.ToString(" Left("+Dado+","+Tamanho+")");
					} else if (Parte==TipoParteTexto.MidVs) {
						// <Removed by: Queiroz at: 28/05/2002-20:08:16 on machine: VTDES_01>
						// ParteTexto = " Left(Right(" & Dado & "," & Len(Dado) + 1 - Inicio & ")," & IIf(Tamanho <> 0, Tamanho, Len(Dado) + 1 - Inicio) & ")"
						// </Removed by: Queiroz at: 28/05/2002-20:08:16 on machine: VTDES_01>

						// <Added by: Queiroz at: 28/05/2002-20:12:39 on machine: VTDES_01>
						ParteTexto = Convert.ToString(" Substring(Cast("+(DadoEhCampo ? "" : "'")+Dado+(DadoEhCampo ? "" : "'")+" AS varchar),"+Inicio+","+Tamanho+")");
						// </Added by: Queiroz at: 28/05/2002-20:12:39 on machine: VTDES_01>
					} else if (Parte==TipoParteTexto.RightVs) {
						ParteTexto = Convert.ToString(" Right("+Dado+","+Tamanho+")");
					}
				}
				else if (VBtoVar==oracle)
				{
					if (Parte==TipoParteTexto.LeftVs) {
						ParteTexto = Convert.ToString(" substr("+Dado+","+Tamanho+")");
					} else if (Parte==TipoParteTexto.MidVs) {
						ParteTexto = " substr("+Dado+","+Inicio+(Tamanho==0 ? ")" : ","+Tamanho+")");
					} else if (Parte==TipoParteTexto.RightVs) {
						ParteTexto = Convert.ToString(" right("+Dado+","+Tamanho+")"); // Falta este
					}
				}
				else if (VBtoVar==Interbase)
				{
					if (Parte==TipoParteTexto.LeftVs) {
						ParteTexto = Convert.ToString(" substr("+Dado+","+Tamanho+")");
					} else if (Parte==TipoParteTexto.MidVs) {
						ParteTexto = " substr("+Dado+","+Inicio+(Tamanho==0 ? ")" : ","+Tamanho+")");
					} else if (Parte==TipoParteTexto.RightVs) {
						ParteTexto = Convert.ToString(" right("+Dado+","+Tamanho+")"); // Falta este
					}
				}
			}
			return ParteTexto;
		}

		public string UltimoComando
		{
			get
			{
				string UltimoComando = "";
				if (CMOD.bRegistrado) UltimoComando = strUltimaQuery;
				return UltimoComando;
			}
		}

		// VBto upgrade warning: Valor As object --> As string
		public string FormataValorCampo(string Valor)
		{
			string FormataValorCampo = "";
			// VSClass.VSDados.Function FormataValorCampo
			// ================================================================================
			// Queiroz em VTDES_01
			// 24/05/2002-14:01:22
			// 
			// Descricao  : Formata um valor da forma que a sintaxe sql dos bancos solicita
			// 
			// Parametros : Valor - Valor a ser formatado
			// 
			// Ex: FormataValorCampo("27/03/1978") = Convert('27/03/1978', 103, DateTime), num banco SQL Server
			// --------------------------------------------------------------------------------
			try
			{	// On Error GoTo ErroFormata
				if (CMOD.bRegistrado) {
					 string resultado = "";
					if (Strings.Trim(Strings.UCase(Valor))=="NULL") {
						resultado = "Null";
					} else {
						if (Information.IsNumeric(Valor)) {
							if (Conexao.FormatoBanco==Access) {
								// resultado = "'" & Valor & "'"
								resultado = Valor;
							} else {
								resultado = Edita.TiraPic((Valor).ToString(), ".");
								resultado = Edita.TrocaPic(resultado, ",", ".");
							}
						} else if (Strings.Mid(Valor, 3, 1)=="/" && Strings.Mid(Valor, 6, 1)=="/" && Valor.Length==10) {
							resultado = Converte(Valor, TipoConversao.TCDataHora);
						} else {
							if (Conexao.FormatoBanco==oracle) {
								resultado = (Valor).ToString();
							} else {
								resultado = Edita.TiraPic((Valor).ToString(), "'");
							}
							if (Conexao.FormatoBanco==Access) {
								if (!((Strings.UCase(Strings.Left(Strings.Trim(Valor), 5))=="CSTR(" || Strings.UCase(Strings.Left(Strings.Trim(Valor), 5))=="CDBL(" || Strings.UCase(Strings.Left(Strings.Trim(Valor), 6))=="CDATE(") && Strings.Right(Strings.Trim(Valor), 1)==")")) {
									resultado = "'"+resultado+"'";
								} else {
									resultado = (Edita.PosPic((Valor).ToString(), "'")==0 ? Strings.Left(Strings.Left(Valor, Edita.PosPic((Valor).ToString(), "("))+"'"+Strings.Mid(Valor, Edita.PosPic((Valor).ToString(), "(")+1), Valor.Length)+"')" : Valor);
								}
							} else if (Conexao.FormatoBanco==SQLServer) {
								resultado = "'"+resultado+"'";
							} else if (Conexao.FormatoBanco==oracle) {
								if (!Strings.UCase(Strings.Left(Strings.Trim(Valor), 3))=="TO_") {
									resultado = Converte(resultado, TipoConversao.tctexto);
								}
							}
						}
					}
					FormataValorCampo = resultado;
				}
				return FormataValorCampo;
			}
			catch
			{	// ErroFormata:
				MessageBox.Show(Information.Err().Description);
			}
			return FormataValorCampo;
		}

		public bool AnexarDocumento(string Tabela, string Campos, string CampoDoc, ref object objRTF)
		{
			return AnexarDocumento(Tabela, Campos, CampoDoc, objRTF, "");
		}
		public bool AnexarDocumento(string Tabela, string Campos, string CampoDoc, ref object objRTF, string Condicao)
		{
			return AnexarDocumento(Tabela, Campos, CampoDoc, objRTF, Condicao, "");
		}
		public bool AnexarDocumento(string Tabela, string Campos, string CampoDoc, ref object objRTF, string Condicao, string Valores)
		{
			bool AnexarDocumento = false;
			 string Sql = "",  Valor = "";
			// VBto upgrade warning: X As int	OnWrite(int)
			 int X;
			 ADODB.Command cmmDocumento = new ADODB.Command();
			 Parameter Param1 = new Parameter();
			 Parameter Param2 = new Parameter();
			 int i;

			try
			{	// On Error GoTo Trata_Erro

				AnexarDocumento = true;


				Sql = "INSERT INTO "+Tabela+" "+(Strings.Trim(Campos)=="" ? "" : "("+Campos+")");
				Sql = Sql+" VALUES(";

				i = 1;
				X = Valores.Length;
				do {
					Valor = Util.ParseString(Valores, cteSeparador, i);
					if (Strings.UCase(Strings.Left(Strings.Trim(Valor), 8))=="CONVERT(" && Strings.Right(Strings.Trim(Valor), 1)==")") {
						Sql = Sql+(i==1 ? "" : ",")+(Edita.PosPic(Valor, "'")==0 ? Strings.Left(Strings.Left(Valor, Edita.PosPic(Valor, ","))+"'"+Strings.Mid(Valor, Edita.PosPic(Valor, ",")+1), Valor.Length)+"')" : Valor);
					} else {
						Sql = Sql+(i==1 ? "" : ",")+FormataValorCampo(Valor);
					}
					i += 1;
					X -= (Valor.Length+cteSeparador.Length);
				} while (!(X==0));
				Sql = Sql+")";

				Sql = "INSERT INTO "+Tabela+"(HRD_HRO_NUM_CORRELATIVO,HRD_ARQUIVO_DOC)";
				Sql = Sql+" VALUES(?,?)";

				cmmDocumento.ActiveConnection = Conexao;
				cmmDocumento.CommandText = Sql;
				Param1 = new Parameter();
				Param1.Direction = adParamInput;
 /*? Param1.Type = */				adNumeric();
				// Param1.Value = Num_Regime
				cmmDocumento.Parameters.Append(Param1);

				Param2 = new Parameter();
				Param2.Direction = adParamInput;
 /*? Param2.Type = */				adBSTR();
				// Param2.Value = rct_Documento.TextRTF
				cmmDocumento.Parameters.Append(Param2);

				cmmDocumento.Execute();
			}
			catch
			{	// Trata_Erro:
			}
			return AnexarDocumento;
		}

		public string sp_GeraCorrelativo(string Banco, int Operacao)
		{
			string sp_GeraCorrelativo = "";
			Util.Erro("Use a funcao Correlativo!");
			// Dim prBanco As Parameter
			// Dim prOperacao As Parameter
			// Dim prRetorno As Parameter
			// Dim prSequencial As Parameter
			// 
			// Dim Store As New adodb.Command
			// Dim Retorno As String
			// Set Store.ActiveConnection = Me.Conexao.DBConnection
			// 
			// Store.CommandText = "USE VTSeg"
			// Store.CommandType = adCmdText
			// Store.Execute
			// 
			// Store.CommandText = "gp_GeraNumCorrelativo"
			// Store.CommandType = adCmdStoredProc
			// 
			// Seta o parâmetro de retorno padrão
			// Set prRetorno = Store.CreateParameter("Return", adInteger, adParamReturnValue)
			// Store.Parameters.Append prRetorno
			// 
			// Seta o parâmetro de entrada Banco
			// Set prBanco = Store.CreateParameter("NomeDoBanco", adVarChar, adParamInput, 8)
			// Store.Parameters.Append prBanco
			// prBanco.Value = Banco
			// 
			// Seta o parâmetro de entrada Operacao
			// Set prOperacao = Store.CreateParameter("OperadorCorrelativo", adInteger, adParamInput)
			// Store.Parameters.Append prOperacao
			// prOperacao.Value = Operacao
			// 
			// Seta o parâmetro de retorno com  o número sequencial
			// Set prSequencial = Store.CreateParameter("NumCorrelativo", adInteger, adParamOutput)
			// Store.Parameters.Append prSequencial
			// 
			// Store.Execute
			// sp_GeraCorrelativo = prSequencial.Value
			// 
			// While Store.Parameters.Count > 0
			// Store.Parameters.Delete 0
			// Wend
			// 
			// 
			// Store.CommandText = "USE " & Banco
			// Store.CommandType = adCmdText
			// Store.Execute
			return sp_GeraCorrelativo;
		}

		public string Correlativo(string Sistema, int Operador)
		{
			return Correlativo(Sistema, Operador, "");
		}
		public string Correlativo(string Sistema, int Operador, string Descricao)
		{
			return Correlativo(Sistema, Operador, Descricao, true);
		}
		public string Correlativo(string Sistema, int Operador, string Descricao, bool Incrementar)
		{
			return Correlativo(Sistema, Operador, Descricao, Incrementar, "");
		}
		public string Correlativo(string Sistema, int Operador, string Descricao, bool Incrementar, string Referencia)
		{
			return Correlativo(Sistema, Operador, Descricao, Incrementar, Referencia, 1);
		}
		public string Correlativo(string Sistema, int Operador, string Descricao, bool Incrementar, string Referencia, int Incremento)
		{
			return Correlativo(Sistema, Operador, Descricao, Incrementar, Referencia, Incremento, 1);
		}
		public string Correlativo(string Sistema, int Operador, string Descricao, bool Incrementar, string Referencia, int Incremento, int Inicio)
		{
			return Correlativo(Sistema, Operador, Descricao, Incrementar, Referencia, Incremento, Inicio, "999999");
		}
		public string Correlativo(string Sistema, int Operador, string Descricao, bool Incrementar, string Referencia, int Incremento, int Inicio, string Mascara)
		{
			string Correlativo = "";
			 VSComando Comando = new VSComando();

			Comando = new VSComando();
			Comando.Texto(Me, "vtseg.dbo.sp_num_correlativo", cmdStoredProcedure);
			Comando.setarParametro("NomeDoSistema", tipChar, parEntrada, 4, Sistema);
			Comando.setarParametro("NumCorrelativo", tipInteger, parEntradaSaida, , Operador);
			Comando.setarParametro("ReferenciaDoOperador", tipLongVarChar, parEntrada, 20, Referencia);
			Comando.setarParametro("SeqDoOperador", tipInteger, parEntrada, , Inicio);
			Comando.setarParametro("IncDaSequencia", tipInteger, parEntrada, , Incremento);
			Comando.setarParametro("AutoInc", tipInteger, parEntrada, , Math.Abs(Incrementar));
			Comando.setarParametro("DisplayFormat", tipChar, parEntrada, 6, Mascara);
			Comando.setarParametro("DescrDoOperador", tipVarChar, parEntrada, 50, Descricao);
			Comando.Executa();
			Correlativo = Comando.Parametro("NumCorrelativo").Value;
			return Correlativo;
		}



	}
}
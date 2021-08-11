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
	public class VSInstala
	{

		//=========================================================

		 const string ArqAtualizacao = "\\Winft.fta";
		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "GetDriveTypeA")] private static extern unsafe int GetDriveType(IntPtr nDrive);
		private unsafe int GetDriveTypeWrp(ref string nDrive)
		{
			int ret;
			IntPtr pnDrive = VBtoConverter.GetByteFromString(nDrive);

			ret = GetDriveType(pnDrive);

			VBtoConverter.GetStringFromByte(ref nDrive, pnDrive);

			return ret;
		}

		 private VSUtil Util = new VSUtil();
		 private VSSeguranca Seguranca = new VSSeguranca();
		 private VSTemp Temp = new VSTemp();
		enum TipoConfig {
			tcTipo = 5,
			tcDsn = 6,
			tcUser = 7,
			tcCatalog = 8,
			tcPassword = 9
		};

		public bool Expirou(object Bdados)
		{
			bool Expirou = false;
			if (CMOD.bRegistrado) {
				 VSRecordset Rs2 = new VSRecordset();
				 string Sql = "";
				Sql = "SELECT * FROM TAB_ATUALIZACAO";
				if (Bdados.AbreTabela(Sql, Rs2)) {
					if (ControleOK(Bdados, (Rs2(0) ? 1 : 0), Rs2(1), Rs2(2), Rs2(3))) {
						if (Rs2(0)==false) {
							if (Convert.ToInt32(DateDiff("d", CDate(Rs2(1)), DateTime.Today))<Convert.ToInt32(Rs2(2))) {
								Expirou = false;
							} else {
								AtualizaExpiro(Bdados, "1", Rs2(1), Rs2(2), PegaSerialDisco(Bdados));
								Expirou = true;
							}
						} else {
							Expirou = true;
						}
					} else {
						Expirou = true;
					}
				} else {
					Expirou = true;
				}
				Bdados.FechaTabela(Rs2);
			}
			return Expirou;
		}

		public void AtualizaExpiro(object Bdados, string StatusDoExpiro, string Data, string Duracao, string Serial)
		{
			if (CMOD.bRegistrado) {
				 string Crc = "";
				 VSRecordset RSTEMP = new VSRecordset();
				 string Sql = "";
				Crc = GeraControle(StatusDoExpiro, Data, Duracao, Serial);
				Sql = "SELECT * FROM TAB_ATUALIZACAO";
				if (Bdados.AbreTabela(Sql, RSTEMP)) {
					Bdados.AtualizaDados("TAB_ATUALIZACAO", Bdados.PreparaValor(StatusDoExpiro, Bdados.Converte(Data, tctexto), Duracao, Crc, Bdados.Converte(Serial, tctexto)), "TAT_EXPIROU, TAT_DAT_ATUALIZADO, TAT_DURACAO, TAT_CONTROLE, TAT_SERIAL");
				} else {
					Bdados.InsereDados("TAB_ATUALIZACAO", Bdados.PreparaValor(StatusDoExpiro, Bdados.Converte(Data, tctexto), Duracao, Crc, Bdados.Converte(Serial, tctexto)), "TAT_EXPIROU, TAT_DAT_ATUALIZADO, TAT_DURACAO, TAT_CONTROLE, TAT_SERIAL");
				}
				Bdados.FechaTabela(RSTEMP);
			}
		}

		private bool ControleOK(object Bdados, string Expirou, string Data, string Duracao, string Controle)
		{
			bool ControleOK = false;
			if (CMOD.bRegistrado) ControleOK = (GeraControle(Expirou, Data, Duracao, PegaSerialDisco(Bdados))==Controle);
			return ControleOK;
		}

		public string GeraControle(string Expirou, string Data, string Duracao, string Serial)
		{
			string GeraControle = "";
			if (CMOD.bRegistrado) {
				 string Resp = "";
				 string A = "";
				 string B = "";
				 string C = "";
				 string D = "";
				 string E = "";
				 string F = "";
				 string X1 = "";
				 string X2 = "";
				 string X3 = "";
				 string X4 = "";
				A = Expirou;
				B = Strings.Left(Data, 2);
				C = Strings.Mid(Data, 4, 2);
				D = Strings.Right(Data, 4);
				E = Duracao;
				F = Serial;
				X1 = ((2*Convert.ToInt16(A))+Convert.ToInt32(B)+Convert.ToInt32(C)+Convert.ToInt32(D)+Convert.ToInt32(E)+Convert.ToInt32(F)).ToString();
				X2 = ((Convert.ToInt16(B)*Convert.ToInt16(B))+Convert.ToInt32(X1)).ToString();
				X3 = ((Convert.ToInt16(C)*(Convert.ToInt16(C)-Convert.ToInt16(A))*(Convert.ToInt16(C)-(2*Convert.ToInt16(A))))+Convert.ToInt32(X2)).ToString();
				X4 = (Convert.ToInt32(X3)+Convert.ToInt32(X1)-Convert.ToInt16(A)-Convert.ToInt32(F)).ToString();
				Resp = (Convert.ToString(System.Convert.ToInt64(Convert.ToInt32(X1)),16).ToUpper()+Convert.ToString(System.Convert.ToInt64(Convert.ToInt32(X2)),16).ToUpper()+Convert.ToString(System.Convert.ToInt64(Convert.ToInt32(X3)),16).ToUpper()+Convert.ToString(System.Convert.ToInt64(Convert.ToInt32(X4)),16).ToUpper()).ToString();
				GeraControle = Resp;
			}
			return GeraControle;
		}

		public bool AchouAtualizacao(object Bdados, string Caminho)
		{
			bool AchouAtualizacao = false;
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string dat = "";
					 string Dur = "";
					 string Serial = "";
					if (FileSystem.Dir(Caminho+ArqAtualizacao, 0)=="") {
						AchouAtualizacao = false;
						return AchouAtualizacao;
					} else {
						if (PegaDadosDeAtualizacao(Bdados, Caminho, ArqAtualizacao, ref dat, ref Dur, ref Serial)) {
							AtualizaExpiro(Bdados, "0", dat, Dur, Serial);
							Util.ApagarArquivo(Caminho+ArqAtualizacao);
							AchouAtualizacao = true;
						} else {
							AchouAtualizacao = false;
						}
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number==68) {
					Resume Next;
				} else if (Information.Err().Number==52) {
					Resume Next;
				} else if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return AchouAtualizacao;
		}

		private bool PegaDadosDeAtualizacao(object Bdados, string Cam, string Arquivo, ref string DataDeAtualizacao, ref string Duracao, ref string Serial)
		{
			bool PegaDadosDeAtualizacao = false;
			if (CMOD.bRegistrado) {
				// VBto upgrade warning: fileFile As int	OnWrite(int)
				 int fileFile;
				 string Linha = "";
				 bool OK;
				 int i;

				fileFile = FileSystem.FreeFile()(0);
				if (FileSystem.Dir(Cam+Arquivo, 0)!="") {
					FileSystem.FileOpen(fileFile, Cam+Arquivo, OpenMode.Input, (OpenAccess)(-1), (OpenShare)(-1), -1);
				} else {
					return PegaDadosDeAtualizacao;
				}

				Linha = FileSystem.LineInput(fileFile);
				if (Linha=="FUTURO") {
					Linha = FileSystem.LineInput(fileFile);
					Linha = Seguranca.DesCriptografa(Linha);
					if (Information.IsNumeric(Linha)) {
						if (ConfereSerialArq(Bdados, Linha)) {
							Serial = Linha;
							Linha = FileSystem.LineInput(fileFile);
							Linha = Seguranca.DesCriptografa(Linha);
							if (IsDate(Linha)) {
								DataDeAtualizacao = Linha;
								Linha = FileSystem.LineInput(fileFile);
								Linha = Seguranca.DesCriptografa(Linha);
								if (Information.IsNumeric(Linha)) {
									Duracao = Linha;
									PegaDadosDeAtualizacao = true;
								} else {
									PegaDadosDeAtualizacao = false;
								}
							} else {
								PegaDadosDeAtualizacao = false;
							}
						} else {
							PegaDadosDeAtualizacao = false;
						}
					} else {
						PegaDadosDeAtualizacao = false;
					}
				} else {
					PegaDadosDeAtualizacao = false;
				}
				FileSystem.FileClose(fileFile);
			}
			return PegaDadosDeAtualizacao;
		}

		public bool GeraArquivoAtualizador(string DataDeAtualizacao, string Duracao, string Serial, string Mun, string Cam)
		{
			bool GeraArquivoAtualizador = false;
			// RETIRAR PARAMETRO MUN, MAS QUEBRARA COMPATIBILIDADE...

			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					// VBto upgrade warning: fileFile As int	OnWrite(int)
					 int fileFile;
					 string Linha = "";
					 bool OK;
					 int i;
					 int J;
					// VBto upgrade warning: K As int	OnWrite(double)
					 int K;

					GeraArquivoAtualizador = false;

					fileFile = FileSystem.FreeFile()(0);

					FileSystem.FileOpen(fileFile, Cam+ArqAtualizacao, OpenMode.Output, (OpenAccess)(-1), (OpenShare)(-1), -1);

					Linha = "FUTURO";
					FileSystem.PrintLine(fileFile, Linha);

					Linha = (Serial).ToString("00000");
					Linha = Seguranca.Criptografa(Linha);
					FileSystem.PrintLine(fileFile, Linha);

					Linha = DataDeAtualizacao;
					Linha = Seguranca.Criptografa(Linha);
					FileSystem.PrintLine(fileFile, Linha);

					Linha = (Duracao).ToString("000");
					Linha = Seguranca.Criptografa(Linha);
					FileSystem.PrintLine(fileFile, Linha);


					for(i=1; i<=10; i++) {
						Linha = "";
						K = Convert.ToInt32(VBMath.Rnd()*10);
						for(J=1; J<=K; J++) {
							Linha = Linha+Chr(Convert.ToInt16(250*VBMath.Rnd()));
						}
						FileSystem.PrintLine(fileFile, Linha);
					}
					Linha = "FUTURO.";
					FileSystem.PrintLine(fileFile, Linha);

					FileSystem.FileClose(fileFile);
					GeraArquivoAtualizador = true;
				}
				return GeraArquivoAtualizador;
			}
			catch
			{	// Trata:
				if (Information.Err().Number==71) {
					Util.Avisa("Disco não encontrado.");
					GeraArquivoAtualizador = false;
					FileSystem.FileClose(fileFile);
				} else if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
					GeraArquivoAtualizador = false;
					FileSystem.FileClose(fileFile);
				}
			}
			return GeraArquivoAtualizador;
		}

		public string PegaSerialDisco(object Bdados)
		{
			string PegaSerialDisco = "";
			if (CMOD.bRegistrado) {
				 string Sql = "";
				 VSRecordset RsAux = new VSRecordset();
				Sql = "SELECT * FROM TAB_ATUALIZACAO";
				if (Bdados.AbreTabela(Sql, RsAux)) {
					PegaSerialDisco = RsAux(4);
				} else {
					PegaSerialDisco = Convert.ToString(0);
				}

				Bdados.FechaTabela(RsAux);
			}
			return PegaSerialDisco;
		}

		private bool ConfereSerialArq(object Bdados, string Serial)
		{
			bool ConfereSerialArq = false;
			if (CMOD.bRegistrado) {
				if (Serial>(PegaSerialDisco(Bdados)).ToString("00000")) {
					ConfereSerialArq = true;
				} else {
					ConfereSerialArq = false;
				}
			}
			return ConfereSerialArq;
		}

		public void NovoPerfil(ref object Formulario, ref object Cabecalho, string Cod_Sistema, string Sistema, string Desc_Formulario, string Caminho)
		{
			if (CMOD.bRegistrado) {
				Formulario.Caption = Formulario.Name;
				Cabecalho.EXIBE(Sistema, Desc_Formulario, Caminho+Cod_Sistema+".gif");
			}
		}

		public void Perfil(object Formulario, object lblForm, object lblModulo, string Usuario, string Sistema, string Desc_Formulario)
		{
			if (CMOD.bRegistrado) {
				Formulario.Caption = "Futuro Tecnologia - "+Sistema;
				Formulario.lblForm.Caption = Desc_Formulario;
				Formulario.lblModulo = Formulario.Name;
				Formulario.lblForm.ToolTipText = "Usuário: "+Usuario;
			}
		}

		public string PegaConfig(string ArqivoConfig, string Banco, TipoConfig Parametro)
		{
			string PegaConfig = "";
			if (CMOD.bRegistrado) {
				 string Val = "";
				 string B = "";
				// TipoBancot
				// dsnd
				// useru
				// catalogoc
				// senhap
				if (FileSystem.Dir(ArqivoConfig, 0)!="") {
					FileSystem.FileOpen(1, ArqivoConfig, OpenMode.Input, (OpenAccess)(-1), (OpenShare)(-1), -1);
					while (!FileSystem.EOF(1)) {
						Val = FileSystem.LineInput(1);
						if (Strings.Right(Val, 1)=="t")
						{

							// APLICACAO
							B = Seguranca.DesCriptografa(Strings.Mid(Val, 2, Val.Length-2));
							if (Banco==B) {
								PegaConfig = Strings.Left(Val, 1);
								if (Parametro==TipoConfig.tcTipo) { FileSystem.FileClose(1); /*? : */ return PegaConfig; }
							}
						}
						else if (Strings.Right(Val, 1)=="d")
						{
							// SERVIDOR
							if (Parametro==TipoConfig.tcDsn) {
								if (Banco==B) {
									PegaConfig = Strings.Left(Val, Val.Length-1);
									FileSystem.FileClose(1);
									return PegaConfig;
								}
							}
						}
						else if (Strings.Right(Val, 1)=="u")
						{
							// USUARIO
							if (Parametro==TipoConfig.tcUser) {
								if (Banco==B) {
									PegaConfig = Seguranca.DesCriptografa(Strings.Left(Val, Val.Length-1));
									FileSystem.FileClose(1);
									return PegaConfig;
								}
							}
						}
						else if (Strings.Right(Val, 1)=="c")
						{
							// SENHA
							if (Parametro==TipoConfig.tcCatalog) {
								if (Banco==B) {
									PegaConfig = Seguranca.DesCriptografa(Strings.Left(Val, Val.Length-1));
									FileSystem.FileClose(1);
									return PegaConfig;
								}
							}
						}
						else if (Strings.Right(Val, 1)=="p")
						{
							// BANCO
							if (Parametro==TipoConfig.tcPassword) {
								if (Banco==B) {
									PegaConfig = Seguranca.DesCriptografa(Strings.Left(Val, Val.Length-1));
									FileSystem.FileClose(1);
									return PegaConfig;
								}
							}
						}
					}
					FileSystem.FileClose(1);
				} else {
					Util.Avisa("Arquivo '"+ArqivoConfig+"' não encontrado.");
				}

				PegaConfig = "";
			}
			return PegaConfig;
		}

		public string PegaConfiguracao(string Sistema, TipoConfig Parametro)
		{
			string PegaConfiguracao = "";
			if (CMOD.bRegistrado) {
				 string Linha = "";
				 string LinhaSistema = "";
				// TipoBancot
				// dsnd
				// useru
				// catalogoc
				// senhap
				 string ArqivoConfiguracao = "";
				ArqivoConfiguracao = App.Path+"\\Conexao.dbc";
				if (FileSystem.Dir(ArqivoConfiguracao, 0)!="") {
					FileSystem.FileOpen(1, ArqivoConfiguracao, OpenMode.Input, (OpenAccess)(-1), (OpenShare)(-1), -1);
					while (!FileSystem.EOF(1)) {
						Linha = FileSystem.LineInput(1);
						if (Strings.Right(Linha, 1)=="A")
						{

							// SISTEMA(APLICACAO)
							LinhaSistema = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Strings.Left(Linha, Linha.Length-1))));
							// If Sistema <> LinhaSistema Then Close #1: Exit Function
						}
						else if (Strings.Right(Linha, 1)=="D")
						{
							// SGBD
							if (Parametro==TipoConfig.tcTipo) {
								if (Sistema==LinhaSistema) {
									PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Strings.Left(Linha, Linha.Length-1))));
									if (Parametro==TipoConfig.tcTipo) { FileSystem.FileClose(1); /*? : */ return PegaConfiguracao; }
								}
							}
						}
						else if (Strings.Right(Linha, 1)=="S")
						{
							// SERVIDOR
							if (Parametro==TipoConfig.tcDsn) {
								if (Sistema==LinhaSistema) {
									PegaConfiguracao = Strings.Left(Linha, Linha.Length-1);
									FileSystem.FileClose(1);
									return PegaConfiguracao;
								}
							}
						}
						else if (Strings.Right(Linha, 1)=="B")
						{
							// BANCO
							// If Sistema = LinhaSistema Then
							// If Parametro = tcCatalog Then
							// PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
							// Close #1
							// Exit Function
							// End If
							// End If

						}
						else if (Strings.Right(Linha, 1)=="U")
						{
							// USUARIO
							if (Parametro==TipoConfig.tcUser) {
								if (Sistema==LinhaSistema) {
									PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Strings.Left(Linha, Linha.Length-1))));
									FileSystem.FileClose(1);
									return PegaConfiguracao;
								}
							}
						}
						else if (Strings.Right(Linha, 1)=="P")
						{
							// SENHA
							if (Parametro==TipoConfig.tcPassword) {
								if (Sistema==LinhaSistema) {
									PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Strings.Left(Linha, Linha.Length-1))));
									FileSystem.FileClose(1);
									return PegaConfiguracao;
								}
							}
						}
						else if (Strings.Right(Linha, 1)=="C")
						{
							// CATALOGO
							if (Parametro==TipoConfig.tcCatalog) {
								if (Sistema==LinhaSistema) {
									PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Strings.Left(Linha, Linha.Length-1))));
									FileSystem.FileClose(1);
									return PegaConfiguracao;
								}
							}
						}
					}
					FileSystem.FileClose(1);
				} else {
					Util.Avisa("Arquivo de Configuração não encontrado.");
				}

				PegaConfiguracao = "";
			}
			return PegaConfiguracao;
		}

		public string GeraConfig(string Tipo, string Sistema, string Dsn, string User, string Password, string Catalog)
		{
			string GeraConfig = "";
			if (CMOD.bRegistrado) {
				// TipoBancot
				// dsnd
				// useru
				// catalogoc
				// senhap
				GeraConfig = Tipo+Seguranca.Criptografa(Sistema)+"t";
				GeraConfig = GeraConfig+"\r\n"+Dsn+"d";
				if (User!="") GeraConfig = GeraConfig+"\r\n"+Seguranca.Criptografa(User)+"u";
				if (Password!="") GeraConfig = GeraConfig+"\r\n"+Seguranca.Criptografa(Password)+"p";
				if (Catalog!="") GeraConfig = GeraConfig+"\r\n"+Seguranca.Criptografa(Catalog)+"c";
			}
			return GeraConfig;
		}

		private VSInstala() : base()
		{
			CMOD.ValidaComponente("CLASS");
		}

	}
}
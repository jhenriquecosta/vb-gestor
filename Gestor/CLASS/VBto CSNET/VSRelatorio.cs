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
	public class VSRelatorio
	{

		//=========================================================

		 private VSClass.VSUtil Util = new VSClass.VSUtil();
		 private VSTexto Edita = new VSTexto();
		 private CRAXDRT.Report Relatorio = new CRAXDRT.Report();
		 private CRAXDRT.Application Aplica = new CRAXDRT.Application();
		 private string Arq;
		 private string Diretorio;

		 private string SubRel;
		 public string Titulo;
		 public bool Arvore;
		 public bool Detalhamento;
		 public bool Exportacao;

		enum TipoRpt {
			Vertical = 0,
			Horizontal = 1
		};

		enum AlinhamentoTexto {
			Centralizado = 0,
			Esquerdo = 1
		};

		/*? Public Enum */		TipoFormulas();
 /*? Normal */
		Especial();
 /*? End Enum */

		public int CopiasDetalhes
		{
			set
			{
				if (CMOD.bRegistrado) Relatorio.Areas.item("D").CopiesToPrint = value;
			}
		}

		public string Arquivo
		{
			get
			{
				string Arquivo = "";
				if (CMOD.bRegistrado) Arquivo = Arq;
				return Arquivo;
			}
		}

		public bool DefinirArquivo(object Bdados, string ArquivoRPT)
		{
			bool DefinirArquivo = false;
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					Arq = ArquivoRPT;

					 int i;
					i = 1;
					i = Strings.InStr(i, ArquivoRPT, "/", CompareMethod.Text);
					Diretorio = "";
					do {
						Diretorio = Strings.Mid(ArquivoRPT, 1, i);
						i = Strings.InStr(i+1, ArquivoRPT, "\\", CompareMethod.Text);
					} while (!(i==0));
					if (FileSystem.Dir(Arq, 0)!="") {
						Relatorio = Aplica.OpenReport(Arq, 1);
						AtualizaBanco(Bdados, Relatorio);
						AtualizaSubRelatorios(Bdados);
						LimparFormulas();
						DefinirArquivo = true;
					} else {
						// Util.Informa "Arquivo " & Mid(ArquivoRPT, InStrRev(ArquivoRPT, "\") + 1) & " não encontrado."
						Util.Informa("Arquivo "+Arq+" não encontrado.");
						DefinirArquivo = false;
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
			return DefinirArquivo;
		}

		public void LimparFormulas()
		{
			LimparFormulas(false);
		}
		public void LimparFormulas(bool Todas)
		{
			if (CMOD.bRegistrado) {
				 int i;
				for(i=1; i<=Relatorio.FormulaFields.Count; i++) {
					if (Todas || Strings.UCase(Strings.Mid((Relatorio.FormulaFields.item(i).Name).ToString(), 3, 2))!="VT") {
						Relatorio.FormulaFields.item(i).Text = "";
					}
				}
			}
		}

		public void Formulas(string Formula, string Valor)
		{
			Formulas(Formula, Valor, VBtoConverter.Normal);
		}
		public void Formulas(string Formula, string Valor, int TipoFormula)
		{
			if (CMOD.bRegistrado) Relatorio.FormulaFields.GetItemByName(Strings.Trim(Formula)).Text = (TipoFormula==VBtoConverter.Normal ? "'"+Edita.TiraPic(Valor, "'")+"'" : Valor);
		}

		public void Parametros(string Parametro, object Valor)
		{
			if (CMOD.bRegistrado) {
				 CRAXDRT.ParameterFieldDefinitions objParametros = new CRAXDRT.ParameterFieldDefinitions();
				 CRAXDRT.ParameterFieldDefinition objParametro = new CRAXDRT.ParameterFieldDefinition();

				objParametros = Relatorio.ParameterFields;
				foreach (int objParametro in  objParametros) {
					if (objParametro.ParameterFieldName==Parametro) {
						objParametro.SetCurrentValue(Valor);
						break;
					}
				} // objParametro

				Relatorio.EnableParameterPrompting = false;
			}
		}

		public bool Imprimir()
		{
			return Imprimir(false);
		}
		public bool Imprimir(bool Confirmacao)
		{
			return Imprimir(Confirmacao, 1);
		}
		public bool Imprimir(bool Confirmacao, int Copias)
		{
			return Imprimir(Confirmacao, Copias, false);
		}
		public bool Imprimir(bool Confirmacao, int Copias, bool Agrupar)
		{
			return Imprimir(Confirmacao, Copias, Agrupar, 1);
		}
		public bool Imprimir(bool Confirmacao, int Copias, bool Agrupar, int InicioPag)
		{
			return Imprimir(Confirmacao, Copias, Agrupar, InicioPag, 1);
		}
		public bool Imprimir(bool Confirmacao, int Copias, bool Agrupar, int InicioPag, int FimPag)
		{
			bool Imprimir = false;
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					// Preparar
					Relatorio.PrintOut(Confirmacao, Copias, Agrupar, InicioPag, FimPag);

					Imprimir = true;
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
				}

			}
			return Imprimir;
		}

		public bool Visualizar()
		{
			bool Visualizar = false;
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					if (!VSVisualizador.Visible) {
						Preparar();
						VSVisualizador.InstancePtr.Rpt.ReportSource = Relatorio;
						VSVisualizador.InstancePtr.Rpt.DisplayGroupTree = Arvore;
						VSVisualizador.InstancePtr.Rpt.EnableDrillDown = Detalhamento;
						VSVisualizador.InstancePtr.Rpt.EnableExportButton = Exportacao;
						VSVisualizador.Text = Titulo;

						VSVisualizador.InstancePtr.Imprimir();

						Visualizar = true;
					} else {
						Util.Avisa("Já existe um relatório sendo visualizado.");
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
				}

			}
			return Visualizar;
		}

		private void Preparar()
		{
			if (CMOD.bRegistrado) {
				Relatorio.DiscardSavedData();
				Relatorio.VerifyOnEveryPrint = true;
			}
		}

		private VSRelatorio() : base()
		{
			CMOD.ValidaComponente("CLASS");
			if (CMOD.bRegistrado) {
				Arvore = false;
				Detalhamento = true;
				Exportacao = true;
				Titulo = "Impressão de Relatório";
			}
		}

		public string SubRelatorio
		{
			get
			{
				string SubRelatorio = "";
				if (CMOD.bRegistrado) SubRelatorio = SubRel;
				return SubRelatorio;
			}

			set
			{
			const int curOnErrorGoToLabel_Default = 0;
			const int curOnErrorGoToLabel_Trata = 1;
			int vOnErrorGoToLabel = curOnErrorGoToLabel_Default;
			try
			{
					vOnErrorGoToLabel = curOnErrorGoToLabel_Trata; /* On Error GoTo Trata */
					if (CMOD.bRegistrado) {
						// VBto upgrade warning: Val As byte	OnWrite(int)
						 byte Val;

						Val = 0;

						SubRel = value;
						if (SubRel!="") {
							Relatorio = Relatorio.OpenSubreport(SubRel);
						pos: ;
							LimparFormulas();
						} else {
							if (!Relatorio.Parent==null) Relatorio = Relatorio.Parent;
						}

					Trata: ;
						if (Information.Err().Number==-2147190528) {
							if (Val==2) {
								Relatorio = Relatorio.Parent;
 /*? Resume */
							} else if (Val==1) {
								Val = 2;
								Relatorio = Relatorio.OpenSubreport(Strings.UCase(SubRel));
								Resume pos;
							} else if (Val==0) {
								Val = 1;
								Relatorio = Relatorio.OpenSubreport(Strings.LCase(SubRel));
								Resume pos;
							}
						} else if (Information.Err().Number!=0) {
							Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
						}
					}

			}
			catch
			{
				switch(vOnErrorGoToLabel) {
					default:
					case curOnErrorGoToLabel_Default:
						// ...
						break;
					case curOnErrorGoToLabel_Trata:
						//? goto Trata;
						break;
				}
			}
			}
		}



		private void AtualizaBanco(object Bdados /* , CRAXDRT.Report R */ )
		{
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					 CRAXDRT.DatabaseTable X = new CRAXDRT.DatabaseTable();
					foreach (int X in  R.Database.Tables) {
						X.SetLogOnInfo(Bdados.Conexao.Dsn, (Bdados.Conexao.FormatoBanco!=SQLServer ? "" : Bdados.Conexao.Catalog), Bdados.Conexao.User, Bdados.Conexao.Password);
					}
					VBtoVar = Bdados.Conexao.FormatoBanco;
					if (VBtoVar==Access)
					{

						R.Database.LogOnServerEx("p2soledb.dll", Bdados.Conexao.Dsn, "", "", "", "OLE DB", Bdados.Conexao.ConnectionString);
					}
					else if ((VBtoVar==SQLServer) || (VBtoVar==oracle) || (VBtoVar==interbase))
					{
						R.Database.LogOnServerEx("p2soledb.dll", Bdados.Conexao.Dsn, Bdados.Conexao.Catalog, Bdados.Conexao.User, Bdados.Conexao.Password, "OLE DB", Bdados.Conexao.ConnectionString);
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
				}

			}
		}

		private void AtualizaSubRelatorios(object Bdados)
		{
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					 CRAXDRT.Report SubRel = new CRAXDRT.Report();
					 object Objeto;
					 Section Secao = new Section();

					foreach (int Secao in  Relatorio.Sections) {
						foreach (int Objeto in  Secao.ReportObjects) {
							VBtoVar = Objeto.Kind;
							if (VBtoVar==crSubreportObject)
							{

								SubRel = Relatorio.OpenSubreport(Objeto.SubreportName);
								AtualizaBanco(Bdados, SubRel);
							}
						} // Objeto
					} // Secao
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
				}

			}
		}

		public void Cabecalho(string Estado, string Cliente, string Secretaria, string Departamento)
		{
			Cabecalho(Estado, Cliente, Secretaria, Departamento, AlinhamentoTexto.Centralizado);
		}
		public void Cabecalho(string Estado, string Cliente, string Secretaria, string Departamento, AlinhamentoTexto Alinhamento)
		{
			const int curOnErrorGoToLabel_Default = 0;
			const int curOnErrorGoToLabel_Trata = 1;
			int vOnErrorGoToLabel = curOnErrorGoToLabel_Default;
			try
			{
				vOnErrorGoToLabel = curOnErrorGoToLabel_Trata; /* On Error GoTo Trata */
				if (CMOD.bRegistrado) {
					// VBto upgrade warning: Val As byte	OnWrite(int)
					 byte Val;

					Val = 0;


					Val = 0;
					if (Alinhamento==AlinhamentoTexto.Esquerdo) {
						Relatorio = Relatorio.OpenSubreport("VSCabEsq.rpt");
					} else {
						Relatorio = Relatorio.OpenSubreport("VSCab.rpt");
					}

				pos: ;

					Formulas("Dep_1", "'"+Estado+"'");
					Formulas("Dep_2", "'"+Cliente+"'");
					Formulas("Dep_3", "'"+Secretaria+"'");
					Formulas("Dep_4", "'"+Departamento+"'");
					Relatorio = Relatorio.Parent;

					return;
				Trata: ;
					if (Information.Err().Number==-2147190528) {
						if (Val==2) {
							Util.Erro("Rodapé não existente.");
						} else if (Val==0) {
							Val = 1;
							Relatorio = Relatorio.OpenSubreport("VSCAB.RPT");
							Resume pos;
						} else if (Val==1) {
							Val = 2;
							Relatorio = Relatorio.OpenSubreport("vscab.rpt");
							Resume pos;
						}
					} else if (Information.Err().Number!=0) {
						Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
					}
				}

			}
			catch
			{
				switch(vOnErrorGoToLabel) {
					default:
					case curOnErrorGoToLabel_Default:
						// ...
						break;
					case curOnErrorGoToLabel_Trata:
						//? goto Trata;
						break;
				}
			}
		}

		public void Rodape(string Administrador, string Cliente, string Endereco, string Cod_Relatorio, string Cod_Usuario)
		{
			Rodape(Administrador, Cliente, Endereco, Cod_Relatorio, Cod_Usuario, TipoRpt.Vertical);
		}
		public void Rodape(string Administrador, string Cliente, string Endereco, string Cod_Relatorio, string Cod_Usuario, TipoRpt Modo)
		{
			const int curOnErrorGoToLabel_Default = 0;
			const int curOnErrorGoToLabel_Trata = 1;
			int vOnErrorGoToLabel = curOnErrorGoToLabel_Default;
			try
			{
				vOnErrorGoToLabel = curOnErrorGoToLabel_Trata; /* On Error GoTo Trata */
				if (CMOD.bRegistrado) {
					// VBto upgrade warning: Val As byte	OnWrite(int)
					 byte Val;

					Val = 0;
					if (Modo==TipoRpt.Horizontal) {
						Relatorio = Relatorio.OpenSubreport("VSRodHor.rpt");
					} else {
						Relatorio = Relatorio.OpenSubreport("VSRod.rpt");
					}

				pos: ;

					Formulas("Administrador", "'"+Administrador+"'");
					Formulas("Cliente", "'"+Cliente+"'");
					Formulas("CodUsuario", "'"+Cod_Usuario+"'");
					Formulas("Endereco", "'"+Endereco+"'");
					Formulas("CodRelatorio", "'"+Cod_Relatorio+"'");
					Relatorio = Relatorio.Parent;

					return;
				Trata: ;
					if (Information.Err().Number==-2147190528) {
						if (Val==2) {
							Util.Erro("Rodapé não existente.");
						} else if (Val==0) {
							Val = 1;
							Relatorio = Relatorio.OpenSubreport("VSROD.RPT");
							Resume pos;
						} else if (Val==1) {
							Val = 2;
							Relatorio = Relatorio.OpenSubreport("vsrod.rpt");
							Resume pos;
						}
					} else if (Information.Err().Number!=0) {
						Util.Erro(Information.Err().Number+" - "+Information.Err().Description+".");
					}
				}

			}
			catch
			{
				switch(vOnErrorGoToLabel) {
					default:
					case curOnErrorGoToLabel_Default:
						// ...
						break;
					case curOnErrorGoToLabel_Trata:
						//? goto Trata;
						break;
				}
			}
		}

		public string Selecao
		{
			set
			{
				if (CMOD.bRegistrado) {
					Relatorio.DiscardSavedData();
					Relatorio.RecordSelectionFormula = value;
					Relatorio.VerifyOnEveryPrint = true;
				}
			}
		}

		public string PreparaSelecaoData(string Campo, string DataInicial)
		{
			return PreparaSelecaoData(Campo, DataInicial, "");
		}
		public string PreparaSelecaoData(string Campo, string DataInicial, string DataFinal)
		{
			string PreparaSelecaoData = "";
			PreparaSelecaoData = Campo+(Strings.Trim(DataFinal)=="" ? " = " : " in ");
			PreparaSelecaoData = PreparaSelecaoData+"Date ("+DataInicial.Year+","+DataInicial.Month+","+DataInicial.Day+") ";
			if (Strings.Trim(DataFinal)!="") {
				PreparaSelecaoData = PreparaSelecaoData+" to Date ("+DataFinal.Year+","+DataFinal.Month+","+DataFinal.Day+")";
			}
			return PreparaSelecaoData;
		}

	}
}
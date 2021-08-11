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
	public class VSTemp
	{

		//=========================================================

		enum TipoOBJ {
			xGrid,
			XTree
		};

		 private VSUtil Util = new VSUtil();

		public string PegaParametro(object Bdados, string Tipo)
		{
			string PegaParametro = "";
			if (CMOD.bRegistrado) {
				 string Sql = "";
				 VSRecordset RS = new VSRecordset();

				Sql = "SELECT TPR_DESCRICAO FROM TAB_PARAMETRO WHERE TPR_TIPO = '"+Tipo+"'";
				if (Bdados.AbreTabela(Sql, RS)) {
					PegaParametro = ""+RS(0);
				}
				Bdados.FechaTabela(RS);
			}
			return PegaParametro;
		}


		public void CarregaTree(object Bdados, object oTree, string Tipo)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS1 = new VSRecordset();
					 VSRecordset Rs2 = new VSRecordset();
					 string REQ = "";
					oTree.NodesCollection.Clear();

					Sql = "SELECT * FROM TAB_GRUPO_COMPONENTE where (TGC_MODELO=1 OR TGC_MODELO=2) AND "+" TGC_TTC_COD_TIPO="+Tipo+" ORDER BY TGC_COD_GRUPO";

					if (Bdados.AbreTabela(Sql, RS1)) {
						RS1.MoveFirst();

						while (!RS1.Eof) {

							REQ = (RS1(4)==1 ? "R" : "G")+RS1(0);

							oTree.AddFolder(REQ, , RS1(1), 1, , true);

							if (RS1(2)==1) oTree.AddOption("N"+RS1(0), oTree.Nodes(REQ), "NDA", false);

							Sql = "SELECT * FROM TAB_COMPONENTE WHERE TCO_TGC_COD_GRUPO="+RS1(0)+" ORDER BY TCO_COD_COMPONENTE";
							if (Bdados.AbreTabela(Sql, Rs2)) {

								while (!Rs2.Eof) {
									if (RS1(2)!=1) {
										oTree.AddCheck("C"+Rs2(0), oTree.Nodes(REQ), (Rs2(2)).ToString(), 0, false);
										oTree.Nodes("C"+Rs2(0)).Tag = "CHECK";
									} else {
										oTree.AddOption("C"+Rs2(0), oTree.Nodes(REQ), (Rs2(2)).ToString(), false);
									}
									Rs2.MoveNext();
								}

							} else {
								Util.Erro("Componentes do Grupo "+RS1(1)+" não encontrado.");
							}
							Bdados.FechaTabela(Rs2);

							RS1.MoveNext();

						}

					}
					Bdados.FechaTabela(RS1);
				}
				return;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
		}

		public void CarregaValoresTree(object Bdados, object oTree, string Tabela, string Condicao)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS = new VSRecordset();
					 int i;

					for(i=1; i<=oTree.NodesCollection.Count; i++) {
						if (oTree.Nodes(i).Tag=="CHECK") {
							oTree.Value(i) = 0;
						} else if (oTree.Nodes(i).Children>0) {
							if (oTree.Nodes(i+1).Tag=="") oTree.Value(i+1) = 1;
						}
					}

					Sql = "SELECT * FROM "+Tabela+" where "+Condicao;
					if (Bdados.AbreTabela(Sql, RS)) {
						RS.MoveFirst();

						while (!RS.Eof) {
							if (""+RS(2)=="") {
								oTree.Value("C"+RS(1)) = 1;
							}
							RS.MoveNext();
						}
					}
					Bdados.FechaTabela(RS);
				}
				return;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
		}

		public void CarregaGrid(object Bdados, object oGrid, string Tipo)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS = new VSRecordset();
					 object ItmX;
					// VBto upgrade warning: i As byte	OnWrite(int)
					 byte i;
					 string REQ = "";

					Sql = "SELECT tco_cod_componente ,tgc_nome ,tco_nome, tgc_modelo,TGC_REQUERIDO FROM "+"TAB_GRUPO_COMPONENTE,TAB_COMPONENTE where TGC_MODELO>=3  AND "+"TCO_TGC_COD_GRUPO=TGC_COD_GRUPO  AND "+" TGC_TTC_COD_TIPO="+Tipo+" ORDER BY TGC_COD_GRUPO,TCO_COD_COMPONENTE ";

					oGrid.AllowColumnReorder = true;
					oGrid.Arrange = 2; // lvwAutoTop
					oGrid.GridLines = true;
					oGrid.View = 3; // lvwReport
					oGrid.FullRowSelect = true;
					oGrid.ListItems.Clear();
					oGrid.ColumnHeaders.Clear();


					if (Bdados.AbreTabela(Sql, RS)) {

						for(i=1; i<=3; i++) {
							oGrid.ColumnHeaders.Add(, , , (oGrid.Width/3)-100);
						}

						while (!RS.Eof) {

							REQ = (RS(4)==1 ? "R" : "G")+RS(0);

							ItmX = oGrid.ListItems.Add(, REQ, (RS(1)).ToString());
							ItmX.SubItems(1) = (RS(2)).ToString();
							ItmX.SubItems(2) = "";

							oGrid.ListItems(REQ).Tag = RS(3);

							RS.MoveNext();
						}
					}
					Bdados.FechaTabela(RS);
				}
				return;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
		}

		public void CarregaValoresGrid(object Bdados, object oGrid, string Tabela, string Condicao)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS = new VSRecordset();
					 int i;
					 string K = "";
					 bool Nao;

					for(i=1; i<=oGrid.ListItems.Count; i++) {
						oGrid.ListItems(i).ListSubItems.item(2).Text = "";
					}


					Sql = "SELECT * FROM "+Tabela+" where "+Condicao;

					if (Bdados.AbreTabela(Sql, RS)) {
						RS.MoveFirst();

						while (!RS.Eof) {
							if (RS(2)!="") {
								Nao = false;
								K = "R"+RS(1);
								oGrid.ListItems(K).ListSubItems.item(2).Text = RS(2);
							}
							RS.MoveNext();
						}
					}
					Bdados.FechaTabela(RS);
				}
				return;
			}
			catch
			{	// Trata:
				if (Information.Err().Number==35601) {
					if (!Nao) {
						Nao = true;
						K = "G"+RS(1);
 /*? Resume */
					} else {
						Util.Erro("Componentes '"+RS(1)+"' com erro.");
					}
				} else {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
		}

		public void PegaValor(object oGrid)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					if (oGrid.ListItems.Count==0) return;

					oGrid.SelectedItem.ListSubItems.item(2).Text = Strings.Trim(Strings.UCase(Interaction.InputBox("Digite o novo valor para '"+oGrid.SelectedItem.ListSubItems.item(1).Text+"':", oGrid.SelectedItem.Text, oGrid.SelectedItem.ListSubItems.item(2).Text, -1, -1)));
				}
				return;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
		}

		public bool CamposGridOK(object oGrid)
		{
			bool CamposGridOK = false;
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 int i;

					for(i=1; i<=oGrid.ListItems.Count; i++) {
						if ((oGrid.ListItems(i).Tag>=4) && (oGrid.ListItems(i).ListSubItems.item(2).Text!="")) {
							if (!Information.IsNumeric(oGrid.ListItems(i).ListSubItems.item(2).Text)) {
								Util.Avisa("Valor numérico no campo "+oGrid.ListItems(i).Text+" é inválido.");
								return CamposGridOK;
							}
						}
						if (Strings.Mid(oGrid.ListItems(i).Key, 1, 1)=="R" && Strings.Trim(oGrid.ListItems(i).ListSubItems.item(2).Text)=="") {
							Util.Avisa("Campo "+oGrid.ListItems(i).Text+" é requerido.");
							return CamposGridOK;
						}
					}
					CamposGridOK = true;
				}
				return CamposGridOK;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
			return CamposGridOK;
		}

		public bool CamposTreeOK(object oTree)
		{
			bool CamposTreeOK = false;
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 int i;
					 int J;
					 int K;
					 bool OK;

					for(i=1; i<=oTree.NodesCollection.Count; i++) {
						if (Strings.Mid(oTree.Nodes(i).Key, 1, 1)=="R") {

							J = oTree.Nodes(i).Children;
							OK = false;
							for(K=i+1; K<=i+J; K++) {
								if (Strings.Mid(oTree.NodesCollection(K).Key, 1, 1)=="N" && oTree.Value(K)==1) break;

								if (oTree.Value(K)==1) {
									OK = true;
									break;
								}
							}
							if (!OK) {
								CamposTreeOK = false;
								Util.Avisa("Campo "+oTree.NodesCollection(i).Text+" é requerido.");
								return CamposTreeOK;
							}

						}
					}
					CamposTreeOK = true;
				}
				return CamposTreeOK;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
			return CamposTreeOK;
		}

		public bool GravarDetalhes(object Bdados, object Obj, TipoOBJ Tipo, string Chave, string Tabela, string CampoChave, string CampoComponente, string CampoValor)
		{
			bool GravarDetalhes = false;
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 int i;

					Bdados.AbreTrans();

					if (Tipo==TipoOBJ.xGrid) {
						for(i=1; i<=Obj.ListItems.Count; i++) {
							if (Strings.Trim(Obj.ListItems(i).ListSubItems.item(2).Text)!="") {
								if (!Bdados.InsereDados(Tabela, Bdados.PreparaValor(Chave, Strings.Mid(Obj.ListItems(i).Key, 2), Obj.ListItems(i).ListSubItems.item(2).Text), CampoChave+","+CampoComponente+","+CampoValor)) {
									Bdados.CancelaTrans();
									GravarDetalhes = false;
									return GravarDetalhes;
								}
							}
						}

					} else if (Tipo==TipoOBJ.XTree) {

						for(i=1; i<=Obj.NodesCollection.Count; i++) {
							if (Strings.Left(Obj.Nodes(i).Key, 1)=="C") {
								if (Obj.Value(i)==1) {
									if (!Bdados.InsereDados(Tabela, Bdados.PreparaValor(Chave, Strings.Mid(Obj.Nodes(i).Key, 2)), CampoChave+","+CampoComponente)) {
										Bdados.CancelaTrans();
										GravarDetalhes = false;
										return GravarDetalhes;
									}
								}
							}
						}
					}

					Bdados.GravaTrans();
					GravarDetalhes = true;
				}
				return GravarDetalhes;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Bdados.CancelaTrans();
					GravarDetalhes = false;
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
					return GravarDetalhes;
				}
			}
			return GravarDetalhes;
		}

		public bool ApagaDetalhes(object Bdados, string Tabela, string Condicao)
		{
			bool ApagaDetalhes = false;
			if (CMOD.bRegistrado) ApagaDetalhes = Bdados.DeletaDados(Tabela, Condicao);
			return ApagaDetalhes;
		}

		public string PegaTabGeral(object Bdados, string Tabela)
		{
			return PegaTabGeral(Bdados, Tabela, null);
		}
		public string PegaTabGeral(object Bdados, string Tabela, object Objeto)
		{
			return PegaTabGeral(Bdados, Tabela, Objeto, "");
		}
		public string PegaTabGeral(object Bdados, string Tabela, object Objeto, string Condicao)
		{
			string PegaTabGeral = "";
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS = new VSRecordset();


					Sql = "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_TIPO = "+" (SELECT DISTINCT  TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME = '"+Tabela+"') AND TGE_CODIGO <> 0 ";

					if (Condicao!="") {
						Sql = Sql+" AND "+Condicao;
					}

					Sql = Sql+" ORDER BY TGE_CODIGO";

					if (Bdados.AbreTabela(Sql, RS)) {
						RS.MoveFirst();
						PegaTabGeral = RS(0);
						if (!Objeto==null) {

							if (/*? TypeOf Objeto Is TextBox */) {
								Objeto = RS(0);
							} else if (/*? TypeOf Objeto Is ComboBox */) {
								Objeto.Clear();
								while (!RS.Eof) {
									Objeto.AddItem(RS(0));
									RS.MoveNext();
								}
							}
						}
					}
					Bdados.FechaTabela(RS);
				}
				return PegaTabGeral;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
			return PegaTabGeral;
		}

		public string PegaCodigoNaGeral(object Bdados, string Tabela, string Nome)
		{
			string PegaCodigoNaGeral = "";
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS = new VSRecordset();

					Sql = "SELECT TGE_CODIGO FROM TAB_GERAL WHERE TGE_TIPO = "+" (SELECT DISTINCT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME = '"+Tabela+"') AND TGE_NOME = '"+Nome+"'";


					if (Bdados.AbreTabela(Sql, RS)) {
						RS.MoveFirst();
						PegaCodigoNaGeral = RS(0);
					}
					Bdados.FechaTabela(RS);
				}
				return PegaCodigoNaGeral;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro("Erro: "+Information.Err().Number+" - "+Information.Err().Description+".");
				}
			}
			return PegaCodigoNaGeral;
		}

		private VSTemp() : base()
		{
			CMOD.ValidaComponente("CLASS");
		}

	}
}
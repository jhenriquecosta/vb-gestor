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
	public class VSDocumento
	{

		//=========================================================

		// VBto upgrade warning: Aplicacao As Word.Application	OnWrite(int)
		 private Word.Application Aplicacao = new Word.Application(); // Object '
		// VBto upgrade warning: Documento As Word.Document	OnWrite(int)
		 private Word.Document Documento = new Word.Document(); // Object '
		// VBto upgrade warning: Faixa As Word.Range	OnWrite(int)
		 private Word.Range Faixa = new Word.Range(); // Object '

		private VSDocumento() : base()
		{
			Aplicacao = CreateObject("Word.Application");
		}

		~VSDocumento()
		{
			Aplicacao = null;
			Documento = null;
			Faixa = null;
		}

		public bool Novo()
		{
			return Novo("");
		}
		public bool Novo(string strModelo)
		{
			bool Novo = false;
			if (!Aplicacao==null) {
				Aplicacao.Visible = true;
				Documento = Aplicacao.Documents.Add(strModelo+(Strings.InStr(1, strModelo, ".dot", CompareMethod.Text)>0 ? "" : ".dot"), true);
				Novo = !Documento==null;
			}
			return Novo;
		}

		public bool Selecionar()
		{
			return Selecionar(0);
		}
		public bool Selecionar(int intInicio)
		{
			return Selecionar(intInicio, -1);
		}
		public bool Selecionar(int intInicio, int intFim)
		{
			bool Selecionar = false;
			if (!Documento==null) {
				if (intFim<0) intFim = Documento.Characters.Count;
				Faixa = Documento.Range(intInicio, intFim);
				Selecionar = !Faixa==null;
			}
			return Selecionar;
		}

		public void Substituir(string strLocalizar, string strSubstituir)
		{
			if (Faixa==null) Selecionar();
			if (!Faixa==null) {
				Faixa.Find.Execute(strLocalizar, strSubstituir, wdReplaceAll);
			}
		}

		public void salvarComo(string strArquivo)
		{
			if (!Documento==null) {
				Documento.SaveAs(strArquivo+(Strings.InStr(1, strArquivo, ".doc", CompareMethod.Text)>0 ? "" : ".doc"));
			}
		}

		public void Celula(int intTabela, int intLinha, int intColuna, string strValor)
		{
			// If Faixa Is Nothing Then Selecionar
			if (intTabela<=Documento.Tables.Count) {
				if (intColuna<=Documento.Tables(intTabela).Columns.Count) {
					if (intLinha>Documento.Tables(intTabela).Rows.Count) {
						inserirLinhas(xz, xz);
					}
					Documento.Tables(intTabela).Cell(intLinha, intColuna).Range.Text = strValor;
				}
			}
		}

		public void inserirLinhas(int intTabela)
		{
			inserirLinhas(intTabela, 0);
		}
		public void inserirLinhas(int intTabela, int intAntesLinha)
		{
			inserirLinhas(intTabela, intAntesLinha, 1);
		}
		public void inserirLinhas(int intTabela, int intAntesLinha, int intQuantidade)
		{
			 int i;
			// VBto upgrade warning: Linha As Row	OnWrite(int)
			 Row Linha = new Row();

			Linha = null;
			if (intTabela<=Documento.Tables.Count) {
				if (intAntesLinha>0) {
					if (intAntesLinha<=Documento.Tables(intTabela).Rows.Count) {
						Linha = Documento.Tables(intTabela).Rows(intAntesLinha);
					}
				}
				do {
					if (!(Linha==null)) {
						Documento.Tables(intTabela).Rows.Add(Linha);
					} else {
						Documento.Tables(intTabela).Rows.Add();
					}
					i += 1;
				} while (i<intQuantidade);
			}
		}

		public void Ativar()
		{
			Documento.Activate();
			Aplicacao.Activate();
		}

		public void Cabecalho(string Estado, string Prefeitura, string Secretaria, string Departamento)
		{
			 object Secao;

			Secao = Documento.Sections(1);
			if (Secao.Headers(1).Exists==true) { // 1= wdHeaderFooterPrimary
				Faixa = Secao.Headers(1).Range;
				Substituir("@Estado", Estado);
				Substituir("@Prefeitura", Prefeitura);
				Substituir("@Secretaria", Secretaria);
				Substituir("@Departamento", Departamento);
			}
		}

		public void textoObjeto(string strLocalizar, string strSubstituir)
		{
			 int i,  J;
			 int qtdShapes;
			 int qtdItens;

			qtdShapes = Documento.Shapes.Count;
			for(i=1; i<=qtdShapes; i++) {
				qtdItens = Documento.Shapes(i).GroupItems.Count;
				for(J=1; J<=qtdItens; J++) {
					if (Documento.Shapes(i).GroupItems(J).TextFrame.HasText) {
						if (Strings.Left(Documento.Shapes(i).GroupItems(J).TextFrame.TextRange.Text, Documento.Shapes(i).GroupItems(J).TextFrame.TextRange.Text.Length-1)==strLocalizar) Documento.Shapes(i).GroupItems(J).TextFrame.TextRange.Text = strSubstituir;
					}
				} // J
			} // i
		}

	}
}
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
	public class VSTexto
	{

		//=========================================================

		 private VSUtil Util = new VSUtil();

		 const string Ponto = ".";
		 const string Virgula = ",";

		[System.Runtime.InteropServices.DllImport("user32", EntryPoint = "SendMessageA")] private static extern unsafe int SendMessage(int hWnd, int wMsg, int wParam, void** lParam);
		private unsafe int SendMessageWrp(int hWnd, int wMsg, int wParam, ref void* lParam)
		{
			int ret;
			fixed (void** plParam = &lParam)
			{
				ret = SendMessage(hWnd, wMsg, wParam, plParam);
			}

			return ret;
		}

		enum TipoChar {
			Letra = 0,
			Numero = 1,
			Valores = 2
		};

		enum TipoFormato {
			Data = 0,
			Cpf = 1,
			Cgc = 2,
			Telefone = 3,
			CEP = 4,
			Monetario = 5,
			Hora = 6,
			UmDV = 7,
			DoisDV = 8,
			PASEP = 9,
			Documento = 10
		};

		public int ListIndexDe(object Combo, string Texto)
		{
			int ListIndexDe = 0;
			if (CMOD.bRegistrado) {
				 int i;
				for(i=0; i<=Combo.ListCount; i++) {
					if (Combo.List(i)==Texto) {
						ListIndexDe = (short)(i);
						return ListIndexDe;
					}
				}
				ListIndexDe = -1;
			}
			return ListIndexDe;
		}

		public int Minuscula(int nKeyAscii)
		{
			int Minuscula = 0;
			if (CMOD.bRegistrado) Minuscula = Convert.ToInt32(Strings.LCase(Chr(nKeyAscii))[0]);
			return Minuscula;
		}

		public int Maiuscula(int nKeyAscii)
		{
			int Maiuscula = 0;
			if (CMOD.bRegistrado) Maiuscula = Convert.ToInt32(Strings.UCase(Chr(nKeyAscii))[0]);
			return Maiuscula;
		}

		public int AceitaDig(int Dig_KeyAscii, TipoChar Dig_Tipo)
		{
			int AceitaDig = 0;
			if (CMOD.bRegistrado) {
				// Permite ou nao a digitacao de um caracter
				AceitaDig = (short)(Dig_KeyAscii);
				if (Dig_KeyAscii==Keys.Back | Dig_KeyAscii==Keys.Space) return AceitaDig;
				if (Dig_Tipo==TipoChar.Letra)
				{

					// Permite aceitação de caracter com ou sem acentuação
					if ((Dig_KeyAscii<65 |  Dig_KeyAscii>90) && (Dig_KeyAscii<97 |  Dig_KeyAscii>122)) {
						if ((Dig_KeyAscii<192 |  Dig_KeyAscii>197) && (Dig_KeyAscii<200 |  Dig_KeyAscii>207) && (Dig_KeyAscii<210 |  Dig_KeyAscii>214) && (Dig_KeyAscii<217 |  Dig_KeyAscii>220) && (Dig_KeyAscii<199)) {
							AceitaDig = 0;
						}
					}
				}
				else if (Dig_Tipo==TipoChar.Numero)
				{
					if (Dig_KeyAscii<48 |  Dig_KeyAscii>57) {
						AceitaDig = 0;
					}
				}
				else if (Dig_Tipo==TipoChar.Valores)
				{
					if ((Dig_KeyAscii<48 |  Dig_KeyAscii>57) &&  Dig_KeyAscii!=44) {
						AceitaDig = 0;
					}
				}
			}
			return AceitaDig;
		}

		public void BuscaItemNaLista(object EditControl, ref int KeyAscii)
		{
			if (CMOD.bRegistrado) {
				 /*? On Error Resume Next  */
				 string buffer = "";
				 int RetVal;
				buffer = Strings.Left(EditControl.Text, EditControl.SelStart)+Chr(KeyAscii);
				RetVal = SendMessageWrp((EditControl.hWnd), 0x14C, -1, ref buffer);
				if (RetVal!=-1) {
					EditControl.ListIndex = RetVal;
					EditControl.Text = EditControl.List(RetVal);
					EditControl.SelStart = buffer.Length;
					EditControl.SelLength = EditControl.Text.Length;
					KeyAscii = 0;
				}
			}
		}

		public bool CriticaCampos(object Form)
		{
			bool CriticaCampos = false;
			if (CMOD.bRegistrado) {
				 /*? On Error Resume Next  */
				 Control Controle = new Control();

 /*? For Each */				Controle(In); /*? Form. */ Controls();
				if (Strings.Trim(Controle.Tag)!="" && Strings.Trim(Controle.Text)=="") {
					if (Information.Err().Number==0) {
						Util.Avisa("Campo '"+Controle.Tag+"' deve ser informado.");
						Controle.Focus();
						CriticaCampos = false;
						return CriticaCampos;
					}
					Err.Clear();
				}
 /*? Next */
				CriticaCampos = true;
			}
			return CriticaCampos;
		}

		public void DestacaCaixa(object EditControl, bool Status)
		{
			if (CMOD.bRegistrado) {
				// Destaca caixa de texto e combo ao receber o foco
				if (!Status) {
					EditControl.ForeColor = 0x800000;
					EditControl.BackColor = 0x80000005;
				} else {
					EditControl.SelStart = 0;
					EditControl.SelLength = Len(Strings.Trim(EditControl));
					EditControl.ForeColor = 0xFFFFFF;
					EditControl.BackColor = 0x800000;
				}
			}
		}

		// VBto upgrade warning: EditControl As object	OnWrite(string, object)
		public string FormataTexto(ref object EditControl, TipoFormato Tipo)
		{
			return FormataTexto(EditControl, Tipo, true);
		}
		public string FormataTexto(ref object EditControl, TipoFormato Tipo, ref bool Agrupar)
		{
			string FormataTexto = "";
			if (CMOD.bRegistrado) {
				// Mascara texto de acordo com seu tipo
				if (Strings.Trim(EditControl)=="") return FormataTexto;
				 int i; string Texto = "";

				if (Tipo==TipoFormato.Data)
				{

					if (Strings.Mid(EditControl, 3, 1)=="/") {
						FormataTexto = EditControl;
						return FormataTexto;
					}
					if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==6) {
						FormataTexto = Strings.Left(EditControl, 2)+"/"+Strings.Mid(EditControl, 3, 2)+"/"+Strings.Right(EditControl, 2);
					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==8) {
						FormataTexto = Strings.Left(EditControl, 2)+"/"+Strings.Mid(EditControl, 3, 2)+"/"+Strings.Right(EditControl, 4);
					} else {
						FormataTexto = "";
					}
					if (IsDate(FormataTexto)) {
						EditControl = FormataTexto;
					} else {
						FormataTexto = "";
						EditControl = "";
					}

				}
				else if (Tipo==TipoFormato.Cpf)
				{
					if (Util.ValidaCpf(EditControl.Text)) {
						if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==11 & Strings.Mid(EditControl, 4, 1)!=".") {
							FormataTexto = Strings.Left(EditControl, 3)+"."+Strings.Mid(EditControl, 4, 3)+"."+Strings.Mid(EditControl, 7, 3)+"-"+Strings.Right(EditControl, 2);
						}
						EditControl = FormataTexto;
					} else {
						FormataTexto = "";
						EditControl = "";
					}

				}
				else if (Tipo==TipoFormato.CEP)
				{
					if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==8) {
						FormataTexto = Strings.Mid(EditControl, 1, 5)+"-";
						FormataTexto += Strings.Mid(EditControl, 6, 3);
					}
					EditControl = FormataTexto;

				}
				else if (Tipo==TipoFormato.Monetario)
				{
					Texto = "";
					if (Information.IsNumeric(EditControl.Text)) {
						i = Strings.InStr(1, EditControl.Text, ",", CompareMethod.Text);
						if (i>0) {
							Texto = Util.Nvl(Strings.Mid(EditControl.Text, 1, i-1), 0);
							Texto = (Texto).ToString("#,##0");
							Texto = Texto+","+Strings.Mid(EditControl.Text, i+1);
						}
						Texto = (EditControl.Text).ToString("Standard");
					}
					FormataTexto = Texto;
					EditControl = Texto;

				}
				else if (Tipo==TipoFormato.Cgc)
				{
					if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==14 & Strings.Mid(EditControl, 4, 1)!=".") {
						if (Strings.Mid(EditControl, 3, 1)!=".") {
							FormataTexto = Strings.Left(EditControl, 2)+"."+Strings.Mid(EditControl, 3, 3)+"."+Strings.Mid(EditControl, 6, 3)+"/"+Strings.Mid(EditControl, 9, 4)+"-"+Strings.Right(EditControl, 2);
						}
					}
					EditControl = FormataTexto;
					// If Not Util.ValidaCgc(EditControl.Text) Then
					// FormataTexto = ""
					// EditControl = ""
					// End If

				}
				else if (Tipo==TipoFormato.Telefone)
				{
					if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==7) {
						FormataTexto = Strings.Left(EditControl, 3)+"-"+Strings.Mid(EditControl, 4);
						EditControl = FormataTexto;

					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==8) {
						FormataTexto = Strings.Left(EditControl, 4)+"-"+Strings.Mid(EditControl, 5);
						EditControl = FormataTexto;

					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==10) {
						FormataTexto = "("+Strings.Trim(Strings.Left(EditControl, 3))+") "+Strings.Mid(EditControl, 4, 3)+"-"+Strings.Mid(EditControl, 7, 4);
						EditControl = FormataTexto;

					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==11) {
						FormataTexto = "("+Strings.Trim(Strings.Left(EditControl, 3))+") "+Strings.Mid(EditControl, 4, 4)+"-"+Strings.Mid(EditControl, 8, 4);
						EditControl = FormataTexto;
					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==13) {
						FormataTexto = EditControl;
					}

				}
				else if (Tipo==TipoFormato.Hora)
				{
					if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)>2 & (Strings.Mid(EditControl, 3, 1)==":" || Strings.Mid(EditControl, 3, 1)>5)) {
						FormataTexto = EditControl;
						return FormataTexto;
					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==1) {
						FormataTexto = "0"+Strings.Left(EditControl, 2)+":00";
						EditControl = FormataTexto;
					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==2) {
						FormataTexto = Strings.Left(EditControl, 2)+":00";
						EditControl = FormataTexto;
					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==4) {
						FormataTexto = Strings.Left(EditControl, 2)+":"+Strings.Right(EditControl, 2);
						EditControl = FormataTexto;
					} else if (System.Runtime.InteropServices.Marshal.SizeOf(EditControl)==6) {
						FormataTexto = Strings.Left(EditControl, 2)+":"+Strings.Mid(EditControl, 3, 2)+":"+Strings.Right(EditControl, 4);
						EditControl = FormataTexto;
					} else {
						FormataTexto = EditControl;
					}

				}
				else if (Tipo==TipoFormato.UmDV)
				{
					 string Str = "";
					if (Agrupar) {
						Texto = "#,##0";
					} else {
						Texto = "0";
					}
					Str = Strings.Mid(EditControl, 1, System.Runtime.InteropServices.Marshal.SizeOf(EditControl)-1);
					Str = (Str).ToString(Texto);
					EditControl = Str+"-"+Strings.Mid(EditControl, System.Runtime.InteropServices.Marshal.SizeOf(EditControl));
					FormataTexto = EditControl;

				}
				else if (Tipo==TipoFormato.DoisDV)
				{
					if (Agrupar) {
						Texto = "#,##0";
					} else {
						Texto = "0";
					}
					Str = Strings.Mid(EditControl, 1, System.Runtime.InteropServices.Marshal.SizeOf(EditControl)-2);
					Str = (Str).ToString(Texto);
					EditControl = Str+"-"+Strings.Mid(EditControl, System.Runtime.InteropServices.Marshal.SizeOf(EditControl)-1);
					FormataTexto = EditControl;

				}
				else if (Tipo==TipoFormato.PASEP)
				{
					Str = BotaPic(EditControl.Text, ".", 3);
					Str = BotaPic(Str, ".", 9);
					Str = BotaPic(Str, ".", 12);
					EditControl = Str;
					FormataTexto = EditControl;

				}
				else if (Tipo==TipoFormato.Documento)
				{
					switch (EditControl.Text.Length) {
						case 14:
						{

							EditControl.Text = FormataTexto(ref EditControl, TipoFormato.Cgc, ref Agrupar);
							break;
						}
						case 11:
						{
							EditControl.Text = FormataTexto(ref EditControl, TipoFormato.Cpf, ref Agrupar);
							break;
						}
						default: {
							EditControl.Text = FormataTexto(ref EditControl, TipoFormato.UmDV, ref Agrupar);
							break;
						}
					} //end switch
					FormataTexto = EditControl.Text;

				}
			}
			return FormataTexto;
		}

		public void LimpaCampos(object Form_Name)
		{
			if (CMOD.bRegistrado) {
				 Control Controle = new Control();
				 /*? On Error Resume Next  */
				foreach (int Controle in  Form_Name.Controls) {
					Controle.Text = "";
					if (Controle.Style!=0) Controle.ListIndex = -1;
				}
			}
		}

		public void HabilitaCampos(object Form_Name, bool Valor)
		{
			if (CMOD.bRegistrado) {
				 Control Controle = new Control();
				 /*? On Error Resume Next  */
				foreach (int Controle in  Form_Name.Controls) {
					if (!/*? TypeOf Controle Is Label */) {
						Controle.Enabled = Valor;
					}
				}
			}
		}

		public void SelecionaTexto(object txt)
		{
			if (CMOD.bRegistrado) {
				txt.SelStart = 0;
				txt.SelLength = System.Runtime.InteropServices.Marshal.SizeOf(txt);
			}
		}

		public void FocalizaCaixa(object Form)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 static object Objeto;
					if (!IsMissing) {
						if (Objeto==null) {
							Objeto = ; /*? Form. */ ActiveControl();
							DestacaCaixa(Objeto, true);
						} else if () {							/*? Form. */ ActiveControl.Name(); /*? <> */ Objeto.Name(); /*? Then */
							DestacaCaixa(Objeto, false);
 /*? Call */							DestacaCaixa(()); /*? Form. */ ActiveControl(, true); /*? ) */
							Objeto = ; /*? Form. */ ActiveControl();
						}
					} else if (!(Objeto==null)) {
						if () {
							DestacaCaixa(Objeto, false);
							Objeto = ; /*? Form. */ ActiveControl();
						}
					}
				}
			}
			catch
			{	// Trata:
				return;
			}
		}

		public void AtualizaCombo(object Bdados, object Combo, string Tabela)
		{
			if (CMOD.bRegistrado) {
				 VSRecordset RS = new VSRecordset();
				Combo.Clear();
				if (Bdados.AbreTabela(Tabela, RS)) {
					while (!RS.Eof) {
						if (!IsNull(RS(0))) {
							if (Strings.Trim(RS(0))!="") {
								Combo.AddItem(RS(0));
							}
						}
						RS.MoveNext();
					}
				}
			}
		}

		public void AtualizaComboGeral(object Bdados, object Combo, string Tabela)
		{
			if (CMOD.bRegistrado) {
				 VSRecordset RS = new VSRecordset();
				 string strSql = "";

				Combo.Clear();
				strSql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO>0 AND TGE_TIPO="+"(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='"+Tabela+"')";
				if (Bdados.AbreTabela(strSql, RS)) {
					while (!RS.Eof) {
						if (!IsNull(RS!TGE_NOME)) {
							if (Strings.Trim(RS!TGE_NOME)!="") {
								Combo.AddItem(RS!TGE_NOME);
								Combo.ItemData(Combo.NewIndex) = RS!TGE_CODIGO;
							}
						}
						RS.MoveNext();
					}
				}
				Bdados.FechaTabela(RS);
			}
		}

		public bool PassaTamanhoCombo(object Combo, int MaxLength)
		{
			bool PassaTamanhoCombo = false;
			if (CMOD.bRegistrado) {
				if (Combo.Text.Length>MaxLength) {
					Util.Avisa("Campo com tamanho máximo de "+MaxLength+" caracteres.");
					Combo.Focus();
					PassaTamanhoCombo = true;
				}
			}
			return PassaTamanhoCombo;
		}

		public string TiraPic(ref string cString, string cChar)
		{
			string TiraPic = "";
			if (CMOD.bRegistrado) {
				 int nPoint;
				if (cChar.Length==1) {
					nPoint = Strings.InStr(cString, cChar, CompareMethod.Text);
					while (nPoint) {
						cString = Strings.Left(cString, nPoint-1)+Strings.Mid(cString, nPoint+1, cString.Length);
						nPoint = Strings.InStr(nPoint, cString, cChar, CompareMethod.Text);
					}
				}
				TiraPic = cString;
			}
			return TiraPic;
		}

		public string TiraTudo(ref string strValor)
		{
			string TiraTudo = "";
			 string strRetorno = "";

			strRetorno = TiraPic(ref strValor, ".");
			strRetorno = TiraPic(ref strRetorno, ",");
			strRetorno = TiraPic(ref strRetorno, "/");
			strRetorno = TiraPic(ref strRetorno, "\\");
			strRetorno = TiraPic(ref strRetorno, "-");
			strRetorno = TiraPic(ref strRetorno, ":");
			strRetorno = TiraPic(ref strRetorno, ";");
			strRetorno = TiraPic(ref strRetorno, "(");
			strRetorno = TiraPic(ref strRetorno, ")");
			TiraTudo = strRetorno;
			return TiraTudo;
		}

		public string BotaPic(string cString, string cChar, int nPos)
		{
			string BotaPic = "";
			if (CMOD.bRegistrado) {
				if (cString.Length>nPos) {
					BotaPic = Strings.Mid(cString, 1, nPos)+cChar+Strings.Mid(cString, nPos+1);
				} else {
					BotaPic = cString;
				}
			}
			return BotaPic;
		}

		public int PosPic(string cString, string cChar)
		{
			int PosPic = 0;
			if (CMOD.bRegistrado) PosPic = Strings.InStr(1, cString, cChar, CompareMethod.Text);
			return PosPic;
		}

		public string TrocaPic(ref string cString, string cChar, string Troca)
		{
			string TrocaPic = "";
			if (CMOD.bRegistrado) {
				 int nPoint;
				if (cChar.Length==1 & Troca.Length==1) {
					nPoint = Strings.InStr(cString, cChar, CompareMethod.Text);
					while (nPoint) {
						cString = Strings.Left(cString, nPoint-1)+Troca+Strings.Mid(cString, nPoint+1, cString.Length);
						nPoint = Strings.InStr(nPoint, cString, cChar, CompareMethod.Text);
					}
				}
				TrocaPic = cString;
			}
			return TrocaPic;
		}

		public int BuscaItemListView(object EditControl, int Coluna, string item)
		{
			int BuscaItemListView = 0;
			try
			{	// On Error GoTo Retorno
				if (CMOD.bRegistrado) {
					 int Linha;

					for(Linha=1; Linha<=EditControl.ListItems.Count; Linha++) {
						EditControl.SelectedItem = EditControl.ListItems(Linha);
						if (Coluna!=0) {
							if (EditControl.SelectedItem.SubItems(Coluna)==item) {
								BuscaItemListView = (short)(Linha);
								break;
							}
						} else {
							if (EditControl.SelectedItem==item) {
								BuscaItemListView = (short)(Linha);
								break;
							}
						}
					}
				}
				return BuscaItemListView;
			}
			catch
			{	// Retorno:
				BuscaItemListView = 0;
			}
			return BuscaItemListView;
		}

		private VSTexto() : base()
		{
			CMOD.ValidaComponente("CLASS");
		}

	}
}
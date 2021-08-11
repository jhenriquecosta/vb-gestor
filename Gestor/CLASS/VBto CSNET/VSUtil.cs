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
	public class VSUtil
	{

		//=========================================================
		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "GetSystemDirectoryA")] private static extern unsafe int apiSysDir(IntPtr lpBuffer, int nSize);
		private unsafe int apiSysDirWrp(ref string lpBuffer, int nSize)
		{
			int ret;
			IntPtr plpBuffer = VBtoConverter.GetByteFromString(lpBuffer);

			ret = apiSysDir(plpBuffer, nSize);

			VBtoConverter.GetStringFromByte(ref lpBuffer, plpBuffer);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = " LogonUserA")] private static extern unsafe int LogonUser(IntPtr lpszUsername, IntPtr lpszDomain, IntPtr lpszPassword, int dwLogonType, int dwLogonProvider, int* phToken);
		private unsafe int LogonUserWrp(ref string lpszUsername, ref string lpszDomain, ref string lpszPassword, int dwLogonType, int dwLogonProvider, ref int phToken)
		{
			int ret;
			IntPtr plpszUsername = VBtoConverter.GetByteFromString(lpszUsername);
			IntPtr plpszDomain = VBtoConverter.GetByteFromString(lpszDomain);
			IntPtr plpszPassword = VBtoConverter.GetByteFromString(lpszPassword);

			fixed (int* pphToken = &phToken)
			{
				ret = LogonUser(plpszUsername, plpszDomain, plpszPassword, dwLogonType, dwLogonProvider, pphToken);
			}

			VBtoConverter.GetStringFromByte(ref lpszUsername, plpszUsername);
			VBtoConverter.GetStringFromByte(ref lpszDomain, plpszDomain);
			VBtoConverter.GetStringFromByte(ref lpszPassword, plpszPassword);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "GetUserNameA")] private static extern unsafe int apiUserName(IntPtr lpBuffer, int* nSize);
		private unsafe int apiUserNameWrp(ref string lpBuffer, ref int nSize)
		{
			int ret;
			IntPtr plpBuffer = VBtoConverter.GetByteFromString(lpBuffer);

			fixed (int* pnSize = &nSize)
			{
				ret = apiUserName(plpBuffer, pnSize);
			}

			VBtoConverter.GetStringFromByte(ref lpBuffer, plpBuffer);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("user32")] private static extern int ExitWindowsEx(int uFlags, int dwReserved);
		[System.Runtime.InteropServices.DllImport("mpr.dll", EntryPoint = "WNetAddConnectionA")] private static extern unsafe int WNetAddConnection(IntPtr lpszNetPath, IntPtr lpszPassword, IntPtr lpszLocalName);
		private unsafe int WNetAddConnectionWrp(ref string lpszNetPath, ref string lpszPassword, ref string lpszLocalName)
		{
			int ret;
			IntPtr plpszNetPath = VBtoConverter.GetByteFromString(lpszNetPath);
			IntPtr plpszPassword = VBtoConverter.GetByteFromString(lpszPassword);
			IntPtr plpszLocalName = VBtoConverter.GetByteFromString(lpszLocalName);

			ret = WNetAddConnection(plpszNetPath, plpszPassword, plpszLocalName);

			VBtoConverter.GetStringFromByte(ref lpszNetPath, plpszNetPath);
			VBtoConverter.GetStringFromByte(ref lpszPassword, plpszPassword);
			VBtoConverter.GetStringFromByte(ref lpszLocalName, plpszLocalName);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "SetComputerNameA")] private static extern unsafe int SetComputerName(IntPtr lpComputerName);
		private unsafe int SetComputerNameWrp(ref string lpComputerName)
		{
			int ret;
			IntPtr plpComputerName = VBtoConverter.GetByteFromString(lpComputerName);

			ret = SetComputerName(plpComputerName);

			VBtoConverter.GetStringFromByte(ref lpComputerName, plpComputerName);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "GetComputerNameA")] private static extern unsafe int GetComputerName(IntPtr lpBuffer, int* nSize);
		private unsafe int GetComputerNameWrp(ref string lpBuffer, ref int nSize)
		{
			int ret;
			IntPtr plpBuffer = VBtoConverter.GetByteFromString(lpBuffer);

			fixed (int* pnSize = &nSize)
			{
				ret = GetComputerName(plpBuffer, pnSize);
			}

			VBtoConverter.GetStringFromByte(ref lpBuffer, plpBuffer);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "DeleteFileA")] private static extern unsafe int DeleteFile(IntPtr lpFileName);
		private unsafe int DeleteFileWrp(ref string lpFileName)
		{
			int ret;
			IntPtr plpFileName = VBtoConverter.GetByteFromString(lpFileName);

			ret = DeleteFile(plpFileName);

			VBtoConverter.GetStringFromByte(ref lpFileName, plpFileName);

			return ret;
		}

		[System.Runtime.InteropServices.DllImport("kernel32")] private static extern void Sleep(int dwMilliseconds);
		 private VSTemp Temp = new VSTemp();


		public struct Retorno
		{
			 public byte OpcaoBotao;
			 public string Resposta;
		};

		enum DesligaComo {
			Logoff = 0,
			Desligar = 1,
			Reiniciar = 2
		};

		public void Pausa(int Tempo_Milisegundos)
		{
			if (CMOD.bRegistrado) Sleep(Tempo_Milisegundos);
		}

		public bool Confirma(string Mensagem)
		{
			return Confirma(Mensagem, "Confirmação");
		}
		public bool Confirma(string Mensagem, string Titulo)
		{
			bool Confirma = false;
			if (CMOD.bRegistrado) {
				// Cria msgbox para o usuario confirmar uma operacao
				CMOD.T.OpcaoBotao = 0;
				VSMensagem.InstancePtr.imgExclama.Visible = true;
				VSMensagem.lblTitulo.Text = Titulo;
				VSMensagem.lblMsg.Text = Mensagem;
				Load(VSMensagem.InstancePtr.cmdBotao(7));
				VSMensagem.InstancePtr.cmdBotao(7).Caption = "&Não";
				VSMensagem.InstancePtr.cmdBotao(7).Left = VSMensagem.lblMsg.Left+VSMensagem.lblMsg.Width-VSMensagem.InstancePtr.cmdBotao(7).Width;
				VSMensagem.InstancePtr.cmdBotao(7).Top = VSMensagem.InstancePtr.cmdBotao(6).Top;
				VSMensagem.InstancePtr.cmdBotao(7).Visible = true;
				VSMensagem.InstancePtr.cmdBotao(6).Caption = "&Sim";
				VSMensagem.InstancePtr.cmdBotao(6).Left = VSMensagem.InstancePtr.cmdBotao(7).Left-VSMensagem.InstancePtr.cmdBotao(6).Width-30;
				VSMensagem.InstancePtr.cmdBotao(6).Top = VSMensagem.InstancePtr.cmdBotao(7).Top;
				VSMensagem.InstancePtr.cmdBotao(7).Cancel = true;
				 VSMensagem dlg = new VSMensagem();  dlg.ShowDialog();
				Confirma = (CMOD.T.OpcaoBotao==6 ? true : false);
			}
			return Confirma;
		}

		public void Erro(string Mensagem)
		{
			if (CMOD.bRegistrado) {
				// Cria msgbox para o usuario informando a existencia de um erro em uma operacao
				VSMensagem.InstancePtr.imgErro.Visible = true;
				MostraMensagem(Mensagem, "Erro");
			}
		}

		public void Informa(string Mensagem)
		{
			if (CMOD.bRegistrado) {
				// Cria uma msgbox contendo uma informacao de Informa ao usuario
				VSMensagem.InstancePtr.imgInforma.Visible = true;
				MostraMensagem(Mensagem, "Aviso");
			}
		}

		public void Mensagem(string Mensagem)
		{
			if (CMOD.bRegistrado) {
				// Cria uma msgbox contendo uma informacao de Informa ao usuario
				VSMensagem.InstancePtr.imgMsg.Visible = true;
				MostraMensagem(Mensagem, "Aviso");
			}
		}

		public void Avisa(string Mensagem)
		{
			if (CMOD.bRegistrado) {
				// Cria uma msgbox contendo uma informacao de Informa ao usuario
				VSMensagem.InstancePtr.imgAvisa.Visible = true;
				MostraMensagem(Mensagem, "Atenção");
			}
		}

		public string Entrada(string Pergunta, string Titulo)
		{
			return Entrada(Pergunta, Titulo, "");
		}
		public string Entrada(string Pergunta, string Titulo, string ValorPadrao)
		{
			string Entrada = "";
			if (CMOD.bRegistrado) {
				// Cria uma msgbox contendo uma informacao de Informa ao usuario
				CMOD.T.Resposta = "";
				VSMensagem.InstancePtr.lblMsg.Visible = false;
				VSMensagem.InstancePtr.txtEntrada.Visible = true;
				VSMensagem.InstancePtr.imgEntrada.Visible = true;
				VSMensagem.lblTitulo.Text = Pergunta;
				VSMensagem.Text = Titulo;
				VSMensagem.lblMsg.Text = Pergunta;
				 VSMensagem dlg = new VSMensagem();  dlg.ShowDialog();
				Entrada = (CMOD.T.Resposta!="" ? CMOD.T.Resposta : ValorPadrao);
			}
			return Entrada;
		}

		private void MostraMensagem(string Mensagem, string Titulo)
		{
			if (CMOD.bRegistrado) {
				// Mostra o conteúdo das janelas de informações ao usuário
				VSMensagem.lblTitulo.Text = Titulo;
				VSMensagem.lblMsg.Text = Mensagem;
				VSMensagem.Text = Titulo;
				 VSMensagem dlg = new VSMensagem();  dlg.ShowDialog();
			}
		}

		public object Nvl(string Conteudo, object Valor_para_Retorno)
		{
			object Nvl = 0;
			if (CMOD.bRegistrado) Nvl = (Strings.Trim(Conteudo)!="" ? Conteudo : Valor_para_Retorno);
			return Nvl;
		}

		public object Nnl(object Conteudo, object Valor_para_Retorno)
		{
			object Nnl = 0;
			if (CMOD.bRegistrado) { Nnl = IIf(!Conteudo); /*? Is Null, */ Conteudo[, Valor_para_Retorno]; /*? ) */ }
			return Nnl;
		}

		public void Negrito(object Formulario, byte Caixa_Inicial, byte Caixa_Final)
		{
			if (CMOD.bRegistrado) {
				// VBto upgrade warning: i As int	OnWrite(byte, int)
				 int i;
				for(i=Caixa_Inicial; i<=Caixa_Final; i++) {
					Formulario.txt(i).FontBold = true;
				}
			}
		}

		public bool ValidaCartao(string Num_Cartao)
		{
			bool ValidaCartao = false;
			if (CMOD.bRegistrado) {
				// Retorna True se numero do cartao de credito for valido
				// VBto upgrade warning: Cartao As byte	OnWrite(double)
				 byte []Cartao = new byte[16+1]; // Recebe os digitos do nº do cartao
				 int Mult_Dig; // Valor da Multip. por 2 dos digitos das posicoes impares do nº
				 int Soma_Total; // Acumula total da soma de todos os digitos do nº do cartao
				 int i;

				Soma_Total = 0;
				Mult_Dig = 0;
				if (Num_Cartao.Length<16) { // Verifica se foi digitado 16 numeros
					ValidaCartao = false;
					return ValidaCartao;
				}
				for(i=1; i<=16; i++) { // Carrega o vetor com o nº do cartao
					Cartao[i] = Conversion.Val(Strings.Mid(Num_Cartao, i, 1));
				}
				for(i=1; i<=16; i++) {
					switch (i) { // Verifica a posicao do digito no numero do cartao
						
						case 1:
						case 3:
						case 5:
						case 7:
						case 9:
						case 11:
						case 13:
						case 15:
						{
							Mult_Dig = Cartao[i]*2;
							if (Mult_Dig>9) {
								Mult_Dig -= 9;
							}
							Soma_Total += Mult_Dig;
							break;
						}
						default: {							// Posicoes pares do numero do cartao
							Soma_Total += Cartao[i];
							break;
						}
					} //end switch
				}
				// Soma tem que ser menor que 150 e  multipla de 10
				if (Soma_Total<150) {
					ValidaCartao = (Soma_Total % 10==0 ? true : false);
				} else {					// Numero de cartao invalido
					ValidaCartao = false;
				}
			}
			return ValidaCartao;
		}

		public bool ValidaCgc(string Cgc)
		{
			bool ValidaCgc = false;
			try
			{	// On Error GoTo Err_CGC
				if (CMOD.bRegistrado) {
					 string strCgc = ""; // armazena a parte do CGC entre /0001- ou seja 0001
					 string Pega_Cgc = ""; // armazena do CPF que será utilizada para o cálculo
					 string Inverte_Cgc = ""; // armazena os digitos do CPF da direita para a esquerda
					// VBto upgrade warning: Dig_Cgc As int	OnWrite(int, string)
					 int Dig_Cgc; // armazena o digito separado para cálculo (uma a um)
					 int Dig_Cgc_Mult; // armazena o digito específico multiplicado pela sua base
					 int Soma_Dig_Cgc_Mult; // armazena a soma dos digitos multiplicados pela sua base(Dig_Cgc_Mult)
					 int Soma_Dig_Cgc_Mult1; // armazena a soma dos 8 primeiros digitos multiplicados pela sua base(Dig_Cgc_Mult)
					 int Soma_Dig_Cgc_Mult2; // armazena a soma dos 4 ultimos digitos multiplicados pela sua base(Dig_Cgc_Mult)
					 double Div_Dig_Cgc_Mult; // armazena a divisão dos digitos*base por 11
					// VBto upgrade warning: Result_Div As int	OnWrite(double)
					 int Result_Div; // armazena inteiro da divisão
					// VBto upgrade warning: Resto_Div As int	OnWrite(int)
					 int Resto_Div; // armazena o resto
					 int Dig_Ver1; // armazena o 1º digito verificador
					 int Dig_Ver2; // armazena o 2º digito verificador
					// VBto upgrade warning: Dig_Ver As string	OnWrite(int)
					 string Dig_Ver = ""; // armazena o digito verificador
					 int i;
					 VSTexto Edita = new VSTexto();

					Soma_Dig_Cgc_Mult = 0;
					Soma_Dig_Cgc_Mult1 = 0;
					Soma_Dig_Cgc_Mult2 = 0;
					Dig_Cgc = 0;
					Dig_Cgc_Mult = 0;
					// Inicia cálculos do 1º dígito
					// Separa os dígitos do CGC que serão multiplicados de 2 a 9.
					// Retira a "/" da máscara de entrada.
					strCgc = Strings.Right(Cgc, 7);
					strCgc = Strings.Left(strCgc, 4);
					Pega_Cgc = Strings.Left(Cgc, 8);
					Pega_Cgc = Strings.Right(Pega_Cgc, 4)+strCgc;
					for(i=2; i<=9; i++) {
						Inverte_Cgc = Strings.Right(Pega_Cgc, i-1);
						Dig_Cgc = int.Parse(Strings.Left(Inverte_Cgc, 1));
						Dig_Cgc_Mult = Dig_Cgc*i;
						Soma_Dig_Cgc_Mult1 += Dig_Cgc_Mult;
					} // i
					// Separa os 4 primeiros dígitos do CGC
					Pega_Cgc = Strings.Left(Cgc, 4);
					for(i=2; i<=5; i++) {
						Inverte_Cgc = Strings.Right(Pega_Cgc, i-1);
						Dig_Cgc = int.Parse(Strings.Left(Inverte_Cgc, 1));
						Dig_Cgc_Mult = Dig_Cgc*i;
						Soma_Dig_Cgc_Mult2 += Dig_Cgc_Mult;
					} // i
					Soma_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult1+Soma_Dig_Cgc_Mult2;
					Div_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult/11;
					Result_Div = Convert.ToInt32(Math.Floor(Convert.ToDouble(Div_Dig_Cgc_Mult))*11);
					Resto_Div = Soma_Dig_Cgc_Mult-Result_Div;
					if (Resto_Div==0 |  Resto_Div==1) {
						Dig_Ver1 = 0;
					} else {
						Dig_Ver1 = 11-Resto_Div;
					}
					Soma_Dig_Cgc_Mult = 0;
					Soma_Dig_Cgc_Mult1 = 0;
					Soma_Dig_Cgc_Mult2 = 0;
					Dig_Cgc = 0;
					Dig_Cgc_Mult = 0;
					// Inicia cálculos do 2º dígito
					strCgc = Strings.Right(Cgc, 7);
					strCgc = Strings.Left(strCgc, 4);
					Pega_Cgc = Strings.Left(Cgc, 8);
					Pega_Cgc = Convert.ToString(Strings.Right(Pega_Cgc, 3)+strCgc+Dig_Ver1);
					for(i=2; i<=9; i++) {
						Inverte_Cgc = Strings.Right(Pega_Cgc, i-1);
						Dig_Cgc = int.Parse(Strings.Left(Inverte_Cgc, 1));
						Dig_Cgc_Mult = Dig_Cgc*i;
						Soma_Dig_Cgc_Mult1 += Dig_Cgc_Mult;
					} // i
					Pega_Cgc = Strings.Left(Cgc, 5);
					for(i=2; i<=6; i++) {
						Inverte_Cgc = Strings.Right(Pega_Cgc, i-1);
						Dig_Cgc = int.Parse(Strings.Left(Inverte_Cgc, 1));
						Dig_Cgc_Mult = Dig_Cgc*i;
						Soma_Dig_Cgc_Mult2 += Dig_Cgc_Mult;
					} // i
					Soma_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult1+Soma_Dig_Cgc_Mult2;
					Div_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult/11;
					Result_Div = Convert.ToInt32(Math.Floor(Convert.ToDouble(Div_Dig_Cgc_Mult))*11);
					Resto_Div = Soma_Dig_Cgc_Mult-Result_Div;
					if (Resto_Div==0 |  Resto_Div==1) {
						Dig_Ver2 = 0;
					} else {
						Dig_Ver2 = 11-Resto_Div;
					}
					Dig_Ver = Convert.ToString(Dig_Ver1+Dig_Ver2);
					// Caso o CGC esteja errado dispara a mensagem
					if (Dig_Ver!=Strings.Right(Cgc, 2)) {
						ValidaCgc = false;
						Avisa("CNPJ inválido.");
					} else {
						ValidaCgc = true;
					}
				}
				return ValidaCgc;
			Exit_CGC: ;
				return ValidaCgc;
			}
			catch
			{	// Err_CGC:
				MessageBox.Show(Error);
				Resume Exit_CGC;
			}
			return ValidaCgc;
		}

		public bool ValidaCpf(string Cpf)
		{
			bool ValidaCpf = false;
			const int curOnErrorGoToLabel_Default = 0;
			const int curOnErrorGoToLabel_Err_CPF = 1;
			int vOnErrorGoToLabel = curOnErrorGoToLabel_Default;
			try
			{
				if (CMOD.bRegistrado) {
					 string Pega_Cpf = ""; // armazena do CPF que será utilizada para o cálculo
					 string Inverte_Cpf = ""; // armazena os digitos do CPF da direita para a esquerda
					// VBto upgrade warning: Dig_Cpf As int	OnWrite(int, string)
					 int Dig_Cpf; // armazena o digito separado para cálculo (uma a um)
					 int Dig_Cpf_Mult; // armazena o digito específico multiplicado pela sua base
					 int Soma_Dig_Cpf_Mult; // armazena a soma dos digitos multiplicados pela sua base(Dig_Cpf_Mult)
					 double Div_Dig_Cpf_Mult; // armazena a divisão dos digitos*base por 11
					// VBto upgrade warning: Result_Div As int	OnWrite(double)
					 int Result_Div; // armazena inteiro da divisão
					// VBto upgrade warning: Resto_Div As int	OnWrite(int)
					 int Resto_Div; // armazena o resto
					 int Dig_Ver1; // armazena o 1º digito verificador
					 int Dig_Ver2; // armazena o 2º digito verificador
					// VBto upgrade warning: Dig_Ver As string	OnWrite(int)
					 string Dig_Ver = ""; // armazena o digito verificador
					 int i;
					vOnErrorGoToLabel = curOnErrorGoToLabel_Err_CPF; /* On Error GoTo Err_CPF */

					Soma_Dig_Cpf_Mult = 0;
					Dig_Cpf = 0;
					Dig_Cpf_Mult = 0;
					Pega_Cpf = Strings.Left(Cpf, 9);

					// Inicia cálculos do 1º dígito
					for(i=2; i<=10; i++) {
						Inverte_Cpf = Strings.Right(Pega_Cpf, i-1);
						Dig_Cpf = int.Parse(Strings.Left(Inverte_Cpf, 1));
						Dig_Cpf_Mult = Dig_Cpf*i;
						Soma_Dig_Cpf_Mult += Dig_Cpf_Mult;
					} // i
					Div_Dig_Cpf_Mult = Soma_Dig_Cpf_Mult/11;

					Result_Div = Convert.ToInt32(Math.Floor(Convert.ToDouble(Div_Dig_Cpf_Mult))*11);
					Resto_Div = Soma_Dig_Cpf_Mult-Result_Div;
					if (Resto_Div==0 |  Resto_Div==1) {
						Dig_Ver1 = 0;
					} else {
						Dig_Ver1 = 11-Resto_Div;
					}

					Pega_Cpf = Convert.ToString(Pega_Cpf+Dig_Ver1); // concatena o CPF com o primeiro digito verificador
					Soma_Dig_Cpf_Mult = 0;
					Dig_Cpf = 0;
					Dig_Cpf_Mult = 0;
					// Inicia cálculos do 2º dígito
					for(i=2; i<=11; i++) {
						Inverte_Cpf = Strings.Right(Pega_Cpf, i-1);
						Dig_Cpf = int.Parse(Strings.Left(Inverte_Cpf, 1));
						Dig_Cpf_Mult = Dig_Cpf*i;
						Soma_Dig_Cpf_Mult += Dig_Cpf_Mult;
					} // i
					Div_Dig_Cpf_Mult = Soma_Dig_Cpf_Mult/11;
					Result_Div = Convert.ToInt32(Math.Floor(Convert.ToDouble(Div_Dig_Cpf_Mult))*11);
					Resto_Div = Soma_Dig_Cpf_Mult-Result_Div;
					if (Resto_Div==0 |  Resto_Div==1) {
						Dig_Ver2 = 0;
					} else {
						Dig_Ver2 = 11-Resto_Div;
					}
					Dig_Ver = Convert.ToString(Dig_Ver1+Dig_Ver2);
					// Caso o CPF esteja errado dispara a mensagem
					if (Dig_Ver!=Strings.Right(Cpf, 2)) {
						ValidaCpf = false;
						Avisa("CPF inválido.");
					} else {
						ValidaCpf = true;
					}
				}
				return ValidaCpf;
			Exit_CPF: ;
				return ValidaCpf;
			Err_CPF: ;
				MessageBox.Show(Error);
				Resume Exit_CPF;

			}
			catch
			{
				switch(vOnErrorGoToLabel) {
					default:
					case curOnErrorGoToLabel_Default:
						// ...
						break;
					case curOnErrorGoToLabel_Err_CPF:
						//? goto Err_CPF;
						break;
				}
			}
			return ValidaCpf;
		}

		public string SystemDir()
		{
			string SystemDir = "";

			if (CMOD.bRegistrado) {
				SystemDir = Strings.Space(50);
				if (apiSysDirWrp(ref SystemDir, 50)>0) {
					SystemDir = Strings.RTrim(SystemDir);
					if (Strings.InStr(SystemDir, Chr(0), CompareMethod.Text)>0) {
						SystemDir = Strings.Left(SystemDir, Strings.InStr(SystemDir, Chr(0), CompareMethod.Text)-1);
					}
					SystemDir = (Strings.Right(SystemDir, 1)!="\\" ? SystemDir+"\\" : SystemDir);
				} else {
					SystemDir = "";
				}
			}
			return SystemDir;
		}

		public void DesligaComputador(DesligaComo Tipo)
		{
			if (CMOD.bRegistrado) ExitWindowsEx(Tipo, 0);
		}

		public string ProcuraArquivo(string Unidade, ref string Arquivo)
		{
			string ProcuraArquivo = "";
			if (CMOD.bRegistrado) {
				 string[] dir_names;
				 int num_dirs;
				 int i;
				 string FileName = "";
				 string new_files = "";
				 int attr;
				 /*? On Error Resume Next  */


				FileName = FileSystem.Dir(Unidade+"\\"+Arquivo, FormWindowState.Normal);
				while (FileName!="") {
					new_files = new_files+Unidade+"\\"+FileName+"\r\n";
					FileName = FileSystem.Dir();
				}
				ProcuraArquivo = ProcuraArquivo+new_files;

				FileName = FileSystem.Dir(Unidade+"\\*.*", FileAttribute.Directory);
				while (FileName!="") {
					attr = 0;
					attr = GetAttr(Unidade+"\\"+FileName);

					if (FileName!="." && FileName!=".." && (attr & FileAttribute.Directory)!=0) {
						num_dirs += 1;
						  dir_names.ReDimPreserve(num_dirs);
						dir_names[num_dirs] = FileName;
					}
					FileName = FileSystem.Dir();
				}

				for(i=1; i<=num_dirs; i++) {
					ProcuraArquivo = ProcuraArquivo+ProcuraArquivo(Unidade+"\\"+dir_names[i], ref Arquivo);
				} // i
			}
			return ProcuraArquivo;
		}

		public bool MapUnidade(ref string Caminho, ref string Senha, ref string Unidade)
		{
			bool MapUnidade = false;
			if (CMOD.bRegistrado) {
				if (Strings.InStr(Unidade, ":", CompareMethod.Text)==0) Unidade = Unidade+":";
				MapUnidade = (WNetAddConnectionWrp(ref Caminho, ref Senha, ref Unidade)>0 ? false : true);
			}
			return MapUnidade;
		}

		public void MudaNomeMaquina(ref string Nome)
		{
			if (CMOD.bRegistrado) SetComputerNameWrp(ref Nome);
		}

		public string PegaIdentidadeComputador()
		{
			string PegaIdentidadeComputador = "";
			if (CMOD.bRegistrado) {
				 string NomeComputador = "";
				NomeComputador = Strings.Space(256);
				GetComputerNameWrp(ref NomeComputador, ref 256);
				PegaIdentidadeComputador = Strings.Left(Strings.Trim(NomeComputador), Len(Strings.Trim(NomeComputador))-1);
			}
			return PegaIdentidadeComputador;
		}

		public void CentralizaForm(object Form, object Mdi)
		{
			 object Left,  Top;	// - "AutoDim"

			if (CMOD.bRegistrado) {
 /*? Form. */				Left = (((Mdi.ScaleWidth-))); /*? Form. */ Width(); /*? ) \ 2) + */ Mdi.Left(); /*? )
				        Form. */				Top = (((Mdi.ScaleHeight-))); /*? Form. */ Height(); /*? ) \ 2) + */ Mdi.Top(); /*? ) */
			}
		}
		public void HabilitaForm(object Form, bool Valor)
		{
			// VBto upgrade warning: Enabled As object --> As bool
			 bool Enabled;	// - "AutoDim"

			if (CMOD.bRegistrado) {
 /*? Form. */				Enabled = Valor;
				if (!Valor) {
					System.Windows.Forms.Cursor.Current = Cursors.AppStarting;
				} else {
					System.Windows.Forms.Cursor.Current = FormWindowState.Normal;
				}
			}
		}
		public bool ApagarArquivo(string Arquivo)
		{
			bool ApagarArquivo = false;
			if (CMOD.bRegistrado) {
				 string A = "";
				A = FileSystem.Dir(Arquivo, 0);
				if (A!="") {
					FileSystem.Kill(Arquivo);
					ApagarArquivo = true;
				} else {
					ApagarArquivo = false;
				}
			}
			return ApagarArquivo;
		}

		public void OrdenaGrid(object Grid, object Coluna)
		{
			if (CMOD.bRegistrado) {
				Grid.Sorted = true;
				if (Grid.SortKey==Coluna.Index-1) {
					Grid.SortOrder = Math.Abs(Grid.SortOrder-1);
				} else {
					Grid.SortOrder = 0;
					Grid.SortKey = Coluna.Index-1;
				}
			}
		}
		public object MontaGrid(object Bdados, object Grid, string Sql)
		{
			object MontaGrid = 0; Tamanho_Colunas(()); /*? ) As Boolean */
			if (CMOD.bRegistrado) {
				 VSRecordset RS = new VSRecordset();
				 object ItmX;
				 int i;

				Grid.AllowColumnReorder = true;
				Grid.Arrange = 2; // lvwAutoTop
				Grid.Gridlines = true;
				Grid.LabelEdit = 1; // lvwManual
				Grid.View = 3; // lvwReport
				Grid.FullRowSelect = true;
				Grid.HotTracking = true;
				Grid.FlatScrollBar = false;
				Grid.HideSelection = false;
				Grid.LabelWrap = false;
				Grid.ListItems.Clear();
				Grid.ColumnHeaders.Clear();

				if (Strings.Trim(Sql)=="") {
					Grid.ListItems.Clear();
				} else {
					if (Bdados.AbreTabela(Sql, RS)) {
						for(i=0; i<=RS.Fields.Count-1; i++) {
							if (i<=UBound(Tamanho_Colunas())) {
								Grid.ColumnHeaders.Add(, , RS.Fields(i).Name, Tamanho_Colunas(i));
							} else {
								Grid.ColumnHeaders.Add(, , RS.Fields(i).Name, (Grid.Width/RS.Fields.Count));
							}
						}
						while (!RS.Eof) {
							ItmX = Grid.ListItems.Add(, , (""+RS(0)).ToString());
							for(i=1; i<=RS.Fields.Count-1; i++) {
								if (!IsNull(RS(i))) {
									ItmX.SubItems(i) = (""+RS(i)).ToString();
								}
							}
							RS.MoveNext();
						}
						MontaGrid = true;
					}
				}
				Bdados.FechaTabela(RS);
			}
			return MontaGrid;
		}

		public void Marcar(ref object Marc, object Fixo, ref bool Valor)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {

					Marc.Checked = Valor;
					if (Marc.Children>0) {
						Marcar(ref Marc.Child, Marc, ref Valor);
					}
					if (!Marc.Parent==null) {
						if (Marc.Parent.Children>1 & !Marc) {
							if (Marc!=Fixo) { Marcar(ref Marc); /*? Next, */ Marc(, Valor); }
						}
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
		}

		public void Integridade(ref object No)
		{
			Integridade(No, false);
		}
		public void Integridade(ref object No, bool Subindo)
		{
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					if (No.Children>0 & !Subindo) {
						Integridade(ref No.Child);
					} else {
						if (No) {
							if (!No.Parent==null) {
								No.Parent.Checked = MarcarPai(No.Parent.Child);
								if (!No.Parent) {
									Integridade(ref No.Parent); /*? Next */
								} else {
									Integridade(ref No.Parent, true);
								}
							}
						} else {
							Integridade(ref No); /*? Next */
						}
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
		}

		public bool MarcarPai(object No)
		{
			bool MarcarPai = false;
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					if (No.Checked) {
						MarcarPai = true;
						return MarcarPai;
					} else {
						if (!No) {
							MarcarPai = MarcarPai(No); /*? Next) */
						}
					}
				}
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return MarcarPai;
		}

		public void CarregaFig(object Img, string Cam, byte TipoMin, byte TipoMax, byte Tam)
		{
			if (CMOD.bRegistrado) {
				 /*? On Error Resume Next  */
				// Tipo = 4-7 p/ Menu
				 string Arq = "";
				 string T = "";

				Img.ListImages.Clear();
				Img.ImageWidth = Tam;
				Img.ImageHeight = Tam;

				Arq = FileSystem.Dir(Cam+"\\Imagens\\*.ico", FormWindowState.Normal);
				while (Arq!="") {
					T = Strings.Mid(Arq, 1, Arq.Length-4);
					if (T.Length>=TipoMin && T.Length<=TipoMax) {
						Img.ListImages.Add(, T, LoadPicture(Cam+"\\Imagens\\"+Arq));
					}
					Arq = FileSystem.Dir();
				}
				if (Img.ListImages.Count==0) {
					Img.ListImages.Add(, "NAOTEM", LoadPicture(Cam+"\\NAOTEM.ico"));
				}
			}
		}

		public DateTime UltimoDiaDoMes(DateTime Data)
		{
			DateTime UltimoDiaDoMes = System.DateTime.Now;
			if (CMOD.bRegistrado) UltimoDiaDoMes = DateAdd("d", -1, "01/"+Strings.Mid(DateAdd("m", 1, Data), 4));
			return UltimoDiaDoMes;
		}

		public DateTime PrimeiroDiaDoMes(DateTime Data)
		{
			DateTime PrimeiroDiaDoMes = System.DateTime.Now;
			if (CMOD.bRegistrado) PrimeiroDiaDoMes = CDate("01/"+Data.Month+"/"+Data.Year);
			return PrimeiroDiaDoMes;
		}

		public void AtualizaCount(object Label, object Grid)
		{
			if (CMOD.bRegistrado) {
				switch (Grid.ListItems.Count) {
					case 0:
					{

						Label.Caption(); /*? = "Nenhum Registro */
						break;
					}
					case 1:
					{
						Label.Caption(); /*? = "1 Registro */
						break;
					}
					default: {
						Label.Caption(); /*? = */ Grid.ListItems.Count(); /*? & " Registros */
						break;
					}
				} //end switch
			}
		}

		public void HabilitarGuias(object ActiveTab)
		{			NumerosGuias(()); /*? ) */
			if (CMOD.bRegistrado) {
				 int numguia;
				for(numguia=1; numguia<=ActiveTab.Tabs.Count; numguia++) {
					ActiveTab.Tabs(numguia).Enabled = Contem(numguia, NumerosGuias());
				} // numguia
				ActiveTab.SelectedTab = NumerosGuias(LBound(NumerosGuias));
			}
		}

		public void ExibirGuias(object ActiveTab)
		{			NumerosGuias(()); /*? ) */
			if (CMOD.bRegistrado) {
				 int numguia;
				for(numguia=1; numguia<=ActiveTab.Tabs.Count; numguia++) {
					ActiveTab.Tabs(numguia).Visible = Contem(numguia, NumerosGuias());
				} // numguia
			}
		}

		public object Contem(object Elemento)
		{
			object Contem = 0; Vetor(()); /*? ) As Boolean */
			if (CMOD.bRegistrado) {
				 int item;
				for(item=LBound(Vetor(0));item<=UBound(Vetor(0));item++) {
					if (Vetor(0)) {
						Contem = true;
						break;
					}
				}
			}
			return Contem;
		}

		public object ParseString(string vsString, string vsDelimiter, int viNumber)
		{
			object ParseString = 0;
			// VSClass.VSUtil.Function ParseString
			// ================================================================================
			// Queiroz em VTDES_01
			// 24/05/2002-11:24:25
			// 
			// Descricao  : Considerando uma string com delimitadores como um vetor, ParseString
			// retorna a substring que esta na posicao viNumber
			// 
			// Parametros : vsString (String) - String com delimitadores
			// vsDelimiter (String) - Delimitador da string
			// viNumber (Integer) - Posicao no "vetor". Comeca em 1
			// 
			// Ex: ParseString("Maranhao#Piaui#Ceara#Para", "#", 3) = "Ceara"
			// --------------------------------------------------------------------------------
			try
			{	// On Error GoTo ErroParsing
				if (CMOD.bRegistrado) {
					 int PosBusca,  PosAnterior;
					 int Elemento;
					 string Valor = "";

					if (Strings.Trim(vsString)=="") return ParseString;
					// If Trim(vsDelimiter) = "" Then Exit Function
					if (viNumber<=0) return ParseString;

					PosBusca = -1*vsDelimiter.Length+1;
					Valor = "";
					Elemento = 0;

					do {
						PosAnterior = PosBusca+vsDelimiter.Length;
						PosBusca = Strings.InStr(PosAnterior, vsString, vsDelimiter, CompareMethod.Text);
						Elemento += 1;
					} while (!((Elemento==viNumber) || (PosBusca==0)));

					if (Elemento==viNumber) {
						if (PosBusca==0) {
							Valor = Strings.Mid(vsString, PosAnterior);
						} else {
							Valor = Strings.Mid(vsString, PosAnterior, PosBusca-1);
							PosBusca = Strings.InStr(1, Valor, vsDelimiter, CompareMethod.Text);
							if (PosBusca>0) Valor = Strings.Mid(Valor, 1, PosBusca-1);
						}
					}

					ParseString = Valor;
				}
				return ParseString;
			}
			catch
			{	// ErroParsing:
				Erro(Information.Err().Description);
			}
			return ParseString;
		}

		private VSUtil() : base()
		{
			CMOD.ValidaComponente("CLASS");
		}

		public int PosicaoLista(string Valor, string Lista, string Delimitador)
		{
			int PosicaoLista = 0;
			 int i; string item = "";

			i = 1;
			do {
				item = Convert.ToString(ParseString(Lista, Delimitador, i));
				if (item==Valor) {
					PosicaoLista = (short)(i);
					break;
				}
				i += 1;
			} while (item!="");
			return PosicaoLista;
		}


	}
}
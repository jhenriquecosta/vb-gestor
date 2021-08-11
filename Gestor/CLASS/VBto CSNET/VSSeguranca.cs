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
	public class VSSeguranca
	{

		//=========================================================

		[System.Runtime.InteropServices.DllImport("kernel32", EntryPoint = "GetVolumeInformationA")] private static extern unsafe int apiSerialNumber(IntPtr lpRootPathName, IntPtr lpVolumeNameBuffer, int nVolumeNameSize, int* lpVolumeSerialNumber, int* lpMaximumComponentLength, int* lpFileSystemFlags, IntPtr lpFileSystemNameBuffer, int nFileSystemNameSize);
		private unsafe int apiSerialNumberWrp(ref string lpRootPathName, ref string lpVolumeNameBuffer, int nVolumeNameSize, ref int lpVolumeSerialNumber, ref int lpMaximumComponentLength, ref int lpFileSystemFlags, ref string lpFileSystemNameBuffer, int nFileSystemNameSize)
		{
			int ret;
			IntPtr plpRootPathName = VBtoConverter.GetByteFromString(lpRootPathName);
			IntPtr plpVolumeNameBuffer = VBtoConverter.GetByteFromString(lpVolumeNameBuffer);
			IntPtr plpFileSystemNameBuffer = VBtoConverter.GetByteFromString(lpFileSystemNameBuffer);

			fixed (int* plpVolumeSerialNumber = &lpVolumeSerialNumber, plpMaximumComponentLength = &lpMaximumComponentLength, plpFileSystemFlags = &lpFileSystemFlags)
			{
				ret = apiSerialNumber(plpRootPathName, plpVolumeNameBuffer, nVolumeNameSize, plpVolumeSerialNumber, plpMaximumComponentLength, plpFileSystemFlags, plpFileSystemNameBuffer, nFileSystemNameSize);
			}

			VBtoConverter.GetStringFromByte(ref lpRootPathName, plpRootPathName);
			VBtoConverter.GetStringFromByte(ref lpVolumeNameBuffer, plpVolumeNameBuffer);
			VBtoConverter.GetStringFromByte(ref lpFileSystemNameBuffer, plpFileSystemNameBuffer);

			return ret;
		}

		 private VSUtil Util = new VSUtil();

		public string Mascara(string Senha)
		{
			string Mascara = "";
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					// VBto upgrade warning: i As int	OnWrite(int, byte)
					 int i;
					 byte Tam;
					Tam = Senha.Length;
					for(i=1; i<=Tam; i++) {
						if (i % 2==0) {
							Mascara = Mascara+Chr(Convert.ToInt32(Strings.Mid(Senha, i, 1)[0])+6+i);
						} else {
							Mascara = Mascara+Chr(Convert.ToInt32(Strings.Mid(Senha, i, 1)[0])+2+i);
						}
					}
				}
				return Mascara;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return Mascara;
		}

		public string Desmascara(string Senha)
		{
			string Desmascara = "";
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Senha_Cripto = "";
					 byte Tam;
					// VBto upgrade warning: i As int	OnWrite(int, byte)
					 int i;
					Tam = Senha.Length;
					for(i=1; i<=Tam; i++) {
						Senha_Cripto = Strings.Mid(Senha, i, 1);
						if (i % 2==0) {
							Desmascara = Desmascara+Chr(Convert.ToInt32(Senha_Cripto[0])-(6+i));
						} else {
							Desmascara = Desmascara+Chr(Convert.ToInt32(Senha_Cripto[0])-(2+i));
						}
					}
				}
				return Desmascara;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}

			}
			return Desmascara;
		}

		public string NumSerie()
		{
			string NumSerie = "";
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Diretorio_Raiz = "";
					 string LabelVol = "";
					 int TamVol;
					 int Dec_SerNum;
					 int Oct_SerNum;
					 int MaxLen;
					 int Flags;
					 string Nome = "";
					 int TamNome;

					Diretorio_Raiz = Strings.Left(Environment.CurrentDirectory, 3);
					if (apiSerialNumberWrp(ref Diretorio_Raiz, ref LabelVol, TamVol, ref Dec_SerNum, ref MaxLen, ref Flags, ref Nome, TamNome)) {
						// Retorna um monte de informações do sistema
						NumSerie = (Convert.ToString(System.Convert.ToInt64(Dec_SerNum),16).ToUpper()).ToString("00000000"); // Formata o numero de serie
					} else {						// Erro na função. Retorna "0000-0000"
						NumSerie = "00000000";
					}
				}
				return NumSerie;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}

			}
			return NumSerie;
		}

		public string Criptografa(ref string Texto)
		{
			string Criptografa = "";
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					 int i;
					 int Tam;
					 string Temp = "";
					Texto = "CRIPTO"+Texto;
					Tam = Texto.Length;
					Temp = "";

					// Debug.Print "Texto: " & Texto
					if (Tam % 2==0) {
						for(i=1; i<=Tam; i++) {
							if (i % 2==0) {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, Tam+2-i, 1)[0]));
							} else {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, i, 1)[0]));
							}
							// Debug.Print Chr(255 - Asc(Mid(Temp, I, 1)))
						}
					} else {
						for(i=1; i<=Tam; i++) {
							if (i % 2==0) {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, i, 1)[0]));
							} else {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, Tam+1-i, 1)[0]));
							}
						}
					}

					// Debug.Print "Enbaralhado + Complementar: " & Temp

					// If Tam <= 255 Then
					// For i = 1 To Tam
					// Mid(Temp, i, 1) = Chr(IIf(Asc(Mid(Temp, i, 1)) + i * 2 > 255, (Asc(Mid(Temp, i, 1)) + i * 2) - 255, 255 - (Asc(Mid(Temp, i, 1)) + i * 2)))
					// Next
					// End If

					// Debug.Print "Incrementação: " & Temp

					Criptografa = Temp;
				}
				return Criptografa;

			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return Criptografa;
		}


		public string DesCriptografa(string Texto)
		{
			string DesCriptografa = "";
			try
			{	// On Error GoTo Trata

				if (CMOD.bRegistrado) {
					 int i;
					 int Tam;
					 string Temp = "";

					Tam = Texto.Length;

					// Debug.Print "Texto: " & Texto

					// If Tam <= 255 Then
					// For i = 1 To Tam
					// Mid(Texto, i, 1) = Chr(IIf(Asc(Mid(Texto, i, 1)) + i * 2 > 255, (Asc(Mid(Texto, i, 1)) + i * 2) - 255, 255 - (Asc(Mid(Texto, i, 1)) + i * 2)))
					// Debug.Print Chr(255 - Asc(Mid(Texto, I, 1)))
					// Next
					// End If

					// Debug.Print "Desincrementação: " & Texto

					if (Tam % 2==0) {
						for(i=1; i<=Tam; i++) {
							if (i % 2==0) {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, Tam+2-i, 1)[0]));
							} else {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, i, 1)[0]));
							}
							// Debug.Print Chr(255 - Asc(Mid(Temp, I, 1)))
						}
					} else {
						for(i=1; i<=Tam; i++) {
							if (i % 2==0) {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, i, 1)[0]));
							} else {
								Temp = Temp+Chr(255-Convert.ToInt32(Strings.Mid(Texto, Tam+1-i, 1)[0]));
							}
						}
					}

					// Debug.Print "Desenbaralhado + Complementar (Resposta): " & Temp

					DesCriptografa = Strings.Mid(Temp, 7);

				}
				return DesCriptografa;

			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Avisa(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return DesCriptografa;
		}

		public string ExisteUsuario(object Bdados, string User)
		{
			string ExisteUsuario = "";
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) {
					 string Sql = "";
					 VSRecordset RS = new VSRecordset();

					ExisteUsuario = "";
					Sql = "SELECT * FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '"+(User)+"'";
					if (Bdados.AbreTabela(Sql, RS)) {
						if (RS!TUS_ATIVO==true || RS!TUS_ATIVO==1) ExisteUsuario = (RS!TUS_NOME);
					}
					Bdados.FechaTabela(RS);
				}
				return ExisteUsuario;
			}
			catch
			{	// Trata:
				if (Information.Err().Number!=0) {
					Util.Erro(Information.Err().Description);
					System.Windows.Forms.Cursor.Current = 0;
				}
			}
			return ExisteUsuario;
		}

		private VSSeguranca() : base()
		{
			CMOD.ValidaComponente("CLASS");
		}

	}
}
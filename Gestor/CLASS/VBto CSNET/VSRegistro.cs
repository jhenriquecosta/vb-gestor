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
	public class VSRegistro
	{

		//=========================================================


		private struct SECURITY_ATTRIBUTES
		{
			 public int nLength;
			 public int lpSecurityDescriptor;
			 public int bInheritHandle;
		};

		 const int KEY_ALL_ACCESS = 0xF003F;
		 const int KEY_CREATE_LINK = 0x20;
		 const int KEY_CREATE_SUB_KEY = 0x4;
		 const int KEY_ENUMERATE_SUB_KEYS = 0x8;
		 const int KEY_EXECUTE = 0x20019;
		 const int KEY_NOTIFY = 0x10;
		 const int KEY_QUERY_VALUE = 0x1;
		 const int KEY_READ = 0x20019;
		 const int KEY_SET_VALUE = 0x2;
		 const int KEY_WRITE = 0x20006;

		// FecharChave
		[System.Runtime.InteropServices.DllImport("advapi32.dll")] private static extern int RegCloseKey(int hkey);
		// CriarChave
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegCreateKeyExA")] private static extern unsafe int RegCreateKeyEx(int hkey, IntPtr lpSubKey, int Reserved, IntPtr lpClass, int dwOptions, int samDesired, SECURITY_ATTRIBUTES* lpSecurityAttributes, int* phkResult, int* lpdwDisposition);
		private unsafe int RegCreateKeyExWrp(int hkey, ref string lpSubKey, int Reserved, string lpClass, int dwOptions, int samDesired, ref SECURITY_ATTRIBUTES lpSecurityAttributes, ref int phkResult, ref int lpdwDisposition)
		{
			int ret;
			IntPtr plpSubKey = VBtoConverter.GetByteFromString(lpSubKey);
			IntPtr plpClass = VBtoConverter.GetByteFromString(lpClass);

			fixed (SECURITY_ATTRIBUTES* plpSecurityAttributes = &lpSecurityAttributes)
			{
				fixed (int* pphkResult = &phkResult, plpdwDisposition = &lpdwDisposition)
				{
					ret = RegCreateKeyEx(hkey, plpSubKey, Reserved, plpClass, dwOptions, samDesired, plpSecurityAttributes, pphkResult, plpdwDisposition);
				}
			}

			VBtoConverter.GetStringFromByte(ref lpSubKey, plpSubKey);
			//VBtoConverter.GetStringFromByte(ref lpClass, plpClass);

			return ret;
		}

		// ApagarChave
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegDeleteKeyA")] private static extern unsafe int RegDeleteKey(int hkey, IntPtr lpSubKey);
		private unsafe int RegDeleteKeyWrp(int hkey, ref string lpSubKey)
		{
			int ret;
			IntPtr plpSubKey = VBtoConverter.GetByteFromString(lpSubKey);

			ret = RegDeleteKey(hkey, plpSubKey);

			VBtoConverter.GetStringFromByte(ref lpSubKey, plpSubKey);

			return ret;
		}

		// ApagarValor
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegDeleteValueA")] private static extern unsafe int RegDeleteValue(int hkey, IntPtr lpValueName);
		private unsafe int RegDeleteValueWrp(int hkey, ref string lpValueName)
		{
			int ret;
			IntPtr plpValueName = VBtoConverter.GetByteFromString(lpValueName);

			ret = RegDeleteValue(hkey, plpValueName);

			VBtoConverter.GetStringFromByte(ref lpValueName, plpValueName);

			return ret;
		}

		// EnumerarChaves
		// Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
		// EnumerarValores
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegEnumValueA")] private static extern unsafe int RegEnumValue(int hkey, int dwIndex, IntPtr lpValueName, int* lpcbValueName, int lpReserved, int* lpType, byte* lpData, int* lpcbData);
		private unsafe int RegEnumValueWrp(int hkey, int dwIndex, ref string lpValueName, ref int lpcbValueName, int lpReserved, ref int lpType, ref byte lpData, ref int lpcbData)
		{
			int ret;
			IntPtr plpValueName = VBtoConverter.GetByteFromString(lpValueName);

			fixed (int* plpcbValueName = &lpcbValueName, plpType = &lpType, plpcbData = &lpcbData)
			{
				fixed (byte* plpData = &lpData)
				{
					ret = RegEnumValue(hkey, dwIndex, plpValueName, plpcbValueName, lpReserved, plpType, plpData, plpcbData);
				}
			}

			VBtoConverter.GetStringFromByte(ref lpValueName, plpValueName);

			return ret;
		}

		// AbrirChave
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegOpenKeyExA")] private static extern unsafe int RegOpenKeyEx(int hkey, IntPtr lpSubKey, int ulOptions, int samDesired, int* phkResult);
		private unsafe int RegOpenKeyExWrp(int hkey, ref string lpSubKey, int ulOptions, int samDesired, ref int phkResult)
		{
			int ret;
			IntPtr plpSubKey = VBtoConverter.GetByteFromString(lpSubKey);

			fixed (int* pphkResult = &phkResult)
			{
				ret = RegOpenKeyEx(hkey, plpSubKey, ulOptions, samDesired, pphkResult);
			}

			VBtoConverter.GetStringFromByte(ref lpSubKey, plpSubKey);

			return ret;
		}

		// ConsultarValor
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA")] private static extern unsafe int RegQueryValueEx(int hkey, IntPtr lpValueName, int lpReserved, int* lpType, void** lpData, int* lpcbData);
		private unsafe int RegQueryValueExWrp(int hkey, ref string lpValueName, int lpReserved, ref int lpType, ref void* lpData, ref int lpcbData)
		{
			int ret;
			IntPtr plpValueName = VBtoConverter.GetByteFromString(lpValueName);

			fixed (int* plpType = &lpType, plpcbData = &lpcbData)
			{
				fixed (void** plpData = &lpData)
				{
					ret = RegQueryValueEx(hkey, plpValueName, lpReserved, plpType, plpData, plpcbData);
				}
			}

			VBtoConverter.GetStringFromByte(ref lpValueName, plpValueName);

			return ret;
		}

		// DefinirValor
		[System.Runtime.InteropServices.DllImport("advapi32.dll", EntryPoint = "RegSetValueExA")] private static extern unsafe int RegSetValueEx(int hkey, IntPtr lpValueName, int Reserved, int dwType, void** lpData, int cbData);
		private unsafe int RegSetValueExWrp(int hkey, ref string lpValueName, int Reserved, int dwType, ref void* lpData, int cbData)
		{
			int ret;
			IntPtr plpValueName = VBtoConverter.GetByteFromString(lpValueName);

			fixed (void** plpData = &lpData)
			{
				ret = RegSetValueEx(hkey, plpValueName, Reserved, dwType, plpData, cbData);
			}

			VBtoConverter.GetStringFromByte(ref lpValueName, plpValueName);

			return ret;
		}


		enum eChave {
			chvClasses = 0x80000000,
			chvUsuarioAtual = 0x80000001,
			chvMaquinaLocal = 0x80000002,
			chvUsuarios = 0x80000003,
			chvConfiguracaoAtual = 0x80000005,
			chvDadosDinamicos = 0x80000006,
			chvDadosPerformance = 0x80000004
		};

		enum eTipoDado {
			tipBinario = 3,
			tipDWord = 4,
			tipDWordBigEndian = 5,
			tipDWordLittleEndian = 4,
			tipStringExpandida = 7,
			tipNenhum = 0,
			tipListaRecursos = 8,
			tipString = 1
		};

		public bool Escrever(eChave Chave, ref string Subchave, ref string Parametro, string Valor, eTipoDado Tipo)
		{
			bool Escrever = false;
			if (CMOD.bRegistrado) {
				 int hChave;

				hChave = Criar(Chave, ref Subchave);
				if (hChave!=0) {
					Escrever = Definir(hChave, ref Parametro, Valor, Tipo);
				}
				Fechar(Chave);
			}
			return Escrever;
		}

		public object Ler(eChave Chave, ref string Subchave, ref string Parametro)
		{
			object Ler = 0;
			if (CMOD.bRegistrado) {
				 int hChave;
				 eTipoDado Tipo;

				hChave = Abrir(Chave, ref Subchave);
				if (hChave!=0) {
					Ler = Consultar(hChave, ref Parametro, ref Tipo);
				}
				Fechar(Chave);
			}
			return Ler;
		}

		public bool Apagar(eChave Chave)
		{
			return Apagar(Chave, "");
		}
		public bool Apagar(eChave Chave, ref string Subchave)
		{
			return Apagar(Chave, Subchave, "");
		}
		public bool Apagar(eChave Chave, ref string Subchave, ref string Parametro)
		{
			bool Apagar = false;
			if (CMOD.bRegistrado) {
				 int hChave;

				if (Parametro!="") {
					hChave = Abrir(Chave, ref Subchave);
					if (hChave!=0) {
						Apagar = ApagarParametro(hChave, ref Parametro);
					}
					Fechar(Chave);
				} else {
					if (Subchave!="") {
						Apagar = ApagarChave(Chave, ref Subchave);
					}
				}
			}
			return Apagar;
		}

		public bool Existe(eChave Chave, ref string Subchave)
		{
			bool Existe = false;
			if (CMOD.bRegistrado) {
				Existe = (Abrir(Chave, ref Subchave)!=0 ? true : false);
				Fechar(Chave);
			}
			return Existe;
		}

		private bool Fechar(eChave Chave)
		{
			bool Fechar = false;
			// PLATAFORMAS
			// --------------
			// Windows 95
			// Windows 98
			// Windows NT 3.1/+
			// Windows 2000
			// Windows CE 1.0/+

			// DESCRICAO
			// --------------
			// Fecha uma chave de registro que tenha sido previamente aberta. Esta prática libera recursos do computador

			if (CMOD.bRegistrado) Fechar = (RegCloseKey(Chave)==0 ? true : false);
			return Fechar;
		}

		private int Criar(eChave Chave, ref string Subchave)
		{
			int Criar = 0;
			// PLATAFORMAS
			// --------------
			// Windows 95
			// Windows 98
			// Windows NT 3.1/+
			// Windows 2000
			// Windows CE 1.0/+

			// DESCRICAO
			// --------------
			// Cria uma nova chave de registro. Se a chave já existir, ela será aberta. O handle da chave aberta é o retorno da função

			if (CMOD.bRegistrado) {
				 int hChave;
				 SECURITY_ATTRIBUTES typSeguranca = new SECURITY_ATTRIBUTES();
				 int lngNovoUsado;

				typSeguranca.nLength = System.Runtime.InteropServices.Marshal.SizeOf(typSeguranca);
				typSeguranca.lpSecurityDescriptor = 0;
				typSeguranca.bInheritHandle = 1;

				Criar = (RegCreateKeyExWrp(Chave, ref Subchave, 0, "", 0, KEY_ALL_ACCESS, ref typSeguranca, ref hChave, ref lngNovoUsado)==0 ? hChave : 0);
			}
			return Criar;
		}

		private bool ApagarChave(eChave Chave, ref string Subchave)
		{
			bool ApagarChave = false;
			// PLATAFORMAS
			// --------------
			// Windows 95
			// Windows 98
			// Windows NT 3.1/+
			// Windows 2000
			// Windows CE 1.0/+

			// DESCRICAO
			// --------------
			// Apaga uma chave de registro com todos os valores que ela contém.
			// No WinNT/2000 ocorrerá um erro se houver subchaves.

			if (CMOD.bRegistrado) ApagarChave = (RegDeleteKeyWrp(Chave, ref Subchave)==0 ? true : false);
			return ApagarChave;
		}

		private bool ApagarParametro(int hChave, ref string Parametro)
		{
			bool ApagarParametro = false;
			// PLATAFORMAS
			// --------------
			// Windows 95
			// Windows 98
			// Windows NT 3.1/+

			// DESCRICAO
			// --------------
			// Apaga um Parametro guardado numa chave específica do registro. Esta funcao só apaga Parametroes; não consegue
			// apagar subchaves.

			if (CMOD.bRegistrado) ApagarParametro = (RegDeleteValueWrp(hChave, ref Parametro)==0 ? true : false);
			return ApagarParametro;
		}

		private int Abrir(eChave Chave, ref string Subchave)
		{
			int Abrir = 0;
			// PLATAFORMAS
			// --------------
			// Windows 32's
			// Windows 98
			// Windows NT 3.1/+

			// DESCRICAO
			// --------------
			// Abre um chave do registro e retorna o handle da chave aberta.
			// Não consegue criar a chave, caso ela não exista.

			if (CMOD.bRegistrado) {
				 int hChave;

				Abrir = (RegOpenKeyExWrp(Chave, ref Subchave, 0, KEY_ALL_ACCESS, ref hChave)==0 ? hChave : 0);
			}
			return Abrir;
		}

		private object Consultar(int hChave, ref string Parametro, ref eTipoDado Tipo)
		{
			object Consultar = 0;
			// PLATAFORMAS
			// --------------
			// Windows 95
			// Windows 98
			// Windows NT 3.1/+
			// Windows 2000
			// Windows CE 1.0/+

			// DESCRICAO
			// --------------
			// Lê um Parametro de uma chave do registro.
			if (CMOD.bRegistrado) {
				 string strBuffer = "";
				 int lngTamanho;
				 int lngSucesso;

				strBuffer = Strings.Space(255);
				lngTamanho = 255;
				lngSucesso = RegQueryValueExWrp(hChave, ref Parametro, 0, ref Tipo, ref strBuffer, ref lngTamanho);
				if (lngSucesso==0) {
					if (Tipo==eTipoDado.tipString ||  Tipo==eTipoDado.tipStringExpandida) {
						Consultar = Strings.Left(strBuffer, lngTamanho-1);
					} else {
						Consultar = strBuffer;
					}
				}
			}
			return Consultar;
		}

		private bool Definir(int hChave, ref string Parametro, string Valor, eTipoDado Tipo)
		{
			bool Definir = false;
			// PLATAFORMAS
			// --------------
			// Windows 95
			// Windows 98
			// Windows NT 3.1/+
			// Windows 2000
			// Windows CE 1.0/+

			// DESCRICAO
			// --------------
			// Escreve um valor numa chave de registro. Se o valor não exisitir, ele será criado.
			// Ao escrever uma string ou um valor numérico simples, o parâmetro lpData deve ser passado
			// por valor. Qualquer outro valor não necessita do ByVal
			if (CMOD.bRegistrado) {
				 string strBuffer = "";

				if (Tipo==eTipoDado.tipString ||  Tipo==eTipoDado.tipStringExpandida) {
					strBuffer = Valor+"\0";
					Definir = (RegSetValueExWrp(hChave, ref Parametro, 0, Tipo, ref strBuffer, strBuffer.Length)==0 ? true : false);
				} else {
					Definir = (RegSetValueExWrp(hChave, ref Parametro, 0, Tipo, ref strBuffer, strBuffer.Length)==0 ? true : false);
				}
			}
			return Definir;
		}

		private VSRegistro() : base()
		{
			CMOD.ValidaComponente("CLASS");
		}

	}
}
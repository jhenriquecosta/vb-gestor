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
	public class CMOD
	{

		//=========================================================
		 public static Retorno T = new Retorno();
		 public static bool bRegistrado; // Indica se o controle está registrado

		[System.Runtime.InteropServices.DllImport("user32")] private static extern int EnumThreadWindows(int dwThreadId, int lpfn, int lParam);
		[System.Runtime.InteropServices.DllImport("user32", EntryPoint = "GetClassNameA")] private static extern unsafe int GetClassName(int hWnd, IntPtr lpClassName, int nMaxCount);
		private static unsafe int GetClassNameWrp(int hWnd, ref string lpClassName, int nMaxCount)
		{
			int ret;
			IntPtr plpClassName = VBtoConverter.GetByteFromString(lpClassName);

			ret = GetClassName(hWnd, plpClassName, nMaxCount);

			VBtoConverter.GetStringFromByte(ref lpClassName, plpClassName);

			return ret;
		}


		// this variable is shared between the two routines below
		 static bool m_ClientIsInterpreted;

		// return True if the client application of this DLL
		// is an interpreted Visual Basic program running in the IDE
		// 
		// NOTE: this code is meant to be inserted in a BAS module
		// inside an ActiveX DLL project

		public static bool Programando()
		{
			bool Programando = false;
			EnumThreadWindows(App.ThreadID, AddressOf EnumThreadWindows_CBK, 0);
			Programando = m_ClientIsInterpreted;
			return Programando;
		}

		// this is a callback function that is executed for each
		// window in the same thead as the DLL

		private static bool EnumThreadWindows_CBK(int hWnd, int lParam)
		{
			bool EnumThreadWindows_CBK = false;
			 char[] buffer = new char[512];
			 int length;
			 string windowClass = "";

			// get the class name of this window
			length = GetClassNameWrp(hWnd, ref buffer, System.Runtime.InteropServices.Marshal.SizeOf(buffer));
			windowClass = Strings.Left(buffer, length);

			if (windowClass=="IDEOwner") {
				// this is the main VB IDE window, therefore
				// the client application is interpreted
				m_ClientIsInterpreted = true;
				// return False to stop evaluation
				EnumThreadWindows_CBK = false;
			} else {
				// return True to continue enumeration
				EnumThreadWindows_CBK = true;
			}
			return EnumThreadWindows_CBK;
		}

		public static void ValidaComponente(string Componente)
		{

			bRegistrado = true;
			return;
			try
			{	// On Error GoTo Trata

				// Faz validacao de registro se o componente estiver sendo
				// utilizado para desenvolvimento
				if (Programando()) {
					 object Reg;
					Reg = CreateObject("VTRegistro.VTReg");
					bRegistrado = Reg.Verifica(Componente);
				} else {
					bRegistrado = true;
				}
				return;
			}
			catch
			{	// Trata:
				bRegistrado = false;
			}
		}



	}
}
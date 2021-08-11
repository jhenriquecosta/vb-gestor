using System;
using System.Drawing;
using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using Microsoft.VisualBasic.Compatibility.VB6;

namespace VSClass
{
	/// <summary>
	/// This is a part of the VBto Converter (www.vbto.net). Copyright (C) 2005-2009 StressSoft Company Ltd. All rights reserved.
	/// </summary>
	public class VBtoConverter
	{
		public static double Fix(double d)
		{
			double ret = d<0 ? Math.Ceiling(d) : Math.Floor(d);
			return ret;
		}


        public static IntPtr GetByteFromString(String Buf)
        {
            IntPtr pBuf = Marshal.StringToHGlobalAnsi(Buf);
            return pBuf;
        }
        public static void GetStringFromByte(ref String Buf, IntPtr pBuf)
        {
            Buf = System.Runtime.InteropServices.Marshal.PtrToStringAnsi(pBuf);
        }


		// === External Consts: ===
		public const int cdlPDPrintSetup = 64;
		public const int cdlPDHidePrintToFile = 1048576;
		public const int Normal = 0;

	}
}

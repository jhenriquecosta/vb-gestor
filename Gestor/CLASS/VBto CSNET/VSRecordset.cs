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
	public class VSRecordset
	{

		//=========================================================

		// VBto upgrade warning: RS As DAO.Recordset	OnWrite(int)
		 private DAO.Recordset RS = new DAO.Recordset();

		public bool Eof
		{
			get
			{
				bool Eof = false;
			try
			{	// On Error GoTo Trata
					if (CMOD.bRegistrado) Eof = RS.Eof;
			}
			catch
			{	// Trata:
			}
				return Eof;
			}
		}

		public bool Bof
		{
			get
			{
				bool Bof = false;
			try
			{	// On Error GoTo Trata
					if (CMOD.bRegistrado) Bof = RS.Bof;
			}
			catch
			{	// Trata:
			}
				return Bof;
			}
		}

		public void Fechar()
		{
			if (CMOD.bRegistrado) RS.Close();
		}

		public void Move(int NumRecords)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) RS.Move(NumRecords);
			}
			catch
			{	// Trata:
			}
		}

		public void MoveNext()
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) RS.MoveNext();
			}
			catch
			{	// Trata:
			}
		}

		public void MoveFirst()
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) RS.MoveFirst();
			}
			catch
			{	// Trata:
			}
		}

		public void MoveLast()
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) RS.MoveLast();
			}
			catch
			{	// Trata:
			}
		}

		public void MovePrevious()
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) RS.MovePrevious();
			}
			catch
			{	// Trata:
			}
		}

		public void Abrir(string Source, VSConexao ActiveConnection)
		{
			Abrir(Source, ActiveConnection, SomenteAvanco);
		}
		public void Abrir(string Source, VSConexao ActiveConnection /* , TipoCursor CursorType */ )
		{
			Abrir(Source, ActiveConnection /* , CursorType */ , SomenteLeitura);
		}
		public void Abrir(string Source, VSConexao ActiveConnection /* , TipoCursor CursorType */  /* , TipoTrava LockType */ )
		{
			Abrir(Source, ActiveConnection /* , CursorType */  /* , LockType */ , -1);
		}
		public void Abrir(string Source, VSConexao ActiveConnection /* , TipoCursor CursorType */  /* , TipoTrava LockType */ , int Options)
		{
			try
			{	// On Error GoTo Trata
				if (CMOD.bRegistrado) RS.Open(Source, ActiveConnection.DBConnection, CursorType, LockType, Options);
			}
			catch
			{	// Trata:
			}
		}

		public int RecordCount
		{
			get
			{
				int RecordCount = 0;
			try
			{	// On Error GoTo Trata
					if (CMOD.bRegistrado) RecordCount = RS.RecordCount;
			}
			catch
			{	// Trata:
			}
				return RecordCount;
			}
		}

		public int AbsolutePosition
		{
			get
			{
				int AbsolutePosition = 0;
			try
			{	// On Error GoTo Trata
					if (CMOD.bRegistrado) AbsolutePosition = RS.AbsolutePosition;
			}
			catch
			{	// Trata:
			}
				return AbsolutePosition;
			}
		}

		public int State
		{
			get
			{
				int State = 0;
			try
			{	// On Error GoTo Trata
					if (CMOD.bRegistrado) State = RS.State;
			}
			catch
			{	// Trata:
			}
				return State;
			}
		}

		private VSRecordset() : base()
		{
			CMOD.ValidaComponente("CLASS");
			if (CMOD.bRegistrado) RS = new Recordset();
		Trata: ;
		}

		public void /* adodb.Fields */ Fields
		{
			get
			{
			try
			{	// On Error GoTo Trata
					if (CMOD.bRegistrado) Fields = RS.Fields;
			}
			catch
			{	// Trata:
			}
			}
		}

		~VSRecordset()
		{
			if (CMOD.bRegistrado) RS = null;
		}

	}
}
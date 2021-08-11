using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Diagnostics;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;

namespace VSClass
{
	/// <summary>
	/// Summary description for VSVisualizador.
	/// </summary>
	public class VSVisualizador : System.Windows.Forms.Form
	{
		private System.Windows.Forms.OpenFileDialog CommonDialog1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public VSVisualizador()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			if (_InstancePtr == null) _InstancePtr = this;
		}

		/// <summary>
		/// Default instance for Form
		/// </summary>
		public static VSVisualizador InstancePtr
		{
			get
			{
				return _InstancePtr == null ? _InstancePtr = new VSVisualizador() : _InstancePtr;
			}
		}
		protected static VSVisualizador _InstancePtr = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		
		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(VSVisualizador));
			this.CommonDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.SuspendLayout();
			//
			// CommonDialog1
			//
			//
			// VSVisualizador
			//
			this.ClientSize = new System.Drawing.Size(451, 312);
			this.Name = "VSVisualizador";
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)));
			this.ForeColor = System.Drawing.SystemColors.ControlText;
			this.MinimizeBox = true;
			this.MaximizeBox = true;
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Resize += new System.EventHandler(this.VSVisualizador_Resize);
			this.Text = "Titulo";
			((System.ComponentModel.ISupportInitialize)(this.Rpt)).EndInit();
			this.ResumeLayout(false);
		}
		#endregion


		//=========================================================
		 string Cam;

		public void Imprimir()
		{
			Imprimir("");
		}
		public void Imprimir(string Caminho)
		{
			Cam = Caminho;
			Rpt.ViewReport();
			 ;
			Application.DoEvents();
		}

		private void VSVisualizador_Resize(object sender, System.EventArgs e)
		{
			Rpt.Top = 0;
			Rpt.Left = 0;
			Rpt.Height = ScaleHeight;
			Rpt.Width = ScaleWidth;
		}

		private void Rpt_PrintButtonClicked(ref bool UseDefault)
		{
			 /*? On Error Resume Next  */

			// CommonDialog1.CancelError = true;
			// CommonDialog1.Flags = VBtoConverter.cdlPDPrintSetup+VBtoConverter.cdlPDHidePrintToFile	// - UPGRADE_WARNING: MSComDlg.CommonDialog property Flags has a new behavior.Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E";
			CommonDialog1.ShowPrinter();

			if (Information.Err().Number==MsgBoxStyle.Critical) UseDefault = false;
		}

	}
}
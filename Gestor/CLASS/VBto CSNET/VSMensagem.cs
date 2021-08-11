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
	/// Summary description for VSMensagem.
	/// </summary>
	public class VSMensagem : System.Windows.Forms.Form
	{
		private AxThreed.AxSSCommand[] cmdBotao = new AxThreed.AxSSCommand[7];
		private AxThreed.AxSSCommand cmdBotao_6;
		private System.Windows.Forms.TextBox txtEntrada;
		private System.Windows.Forms.TextBox lblMsg;
		private System.Windows.Forms.PictureBox imgInforma;
		private System.Windows.Forms.PictureBox imgEntrada;
		private System.Windows.Forms.PictureBox imgExclama;
		private System.Windows.Forms.PictureBox imgMsg;
		private System.Windows.Forms.PictureBox imgAvisa;
		private System.Windows.Forms.PictureBox imgErro;
		private System.Windows.Forms.Label lblTitulo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public VSMensagem()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			cmdBotao[6] = cmdBotao_6;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
			if (_InstancePtr == null) _InstancePtr = this;
		}

		/// <summary>
		/// Default instance for Form
		/// </summary>
		public static VSMensagem InstancePtr
		{
			get
			{
				return _InstancePtr == null ? _InstancePtr = new VSMensagem() : _InstancePtr;
			}
		}
		protected static VSMensagem _InstancePtr = null;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(VSMensagem));
			this.components = new System.ComponentModel.Container();
			this.cmdBotao_6 = new AxThreed.AxSSCommand();
			this.txtEntrada = new System.Windows.Forms.TextBox();
			this.lblMsg = new System.Windows.Forms.TextBox();
			this.imgInforma = new System.Windows.Forms.PictureBox();
			this.imgEntrada = new System.Windows.Forms.PictureBox();
			this.imgExclama = new System.Windows.Forms.PictureBox();
			this.imgMsg = new System.Windows.Forms.PictureBox();
			this.imgAvisa = new System.Windows.Forms.PictureBox();
			this.imgErro = new System.Windows.Forms.PictureBox();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.SuspendLayout();
			//
			// cmdBotao_6
			//
			this.cmdBotao_6.Name = "cmdBotao_6";
			this.cmdBotao_6.TabIndex = 2;
			this.cmdBotao_6.Location = new System.Drawing.Point(178, 123);
			this.cmdBotao_6.Size = new System.Drawing.Size(77, 29);
			this.cmdBotao_6.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)System.Drawing.FontStyle.Bold), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdBotao_6.ClickEvent += new System.EventHandler(this.cmdBotao_6_ClickEvent);
			//
			// txtEntrada
			//
			this.txtEntrada.Name = "txtEntrada";
			this.txtEntrada.Visible = false;
			this.txtEntrada.TabIndex = 0;
			this.txtEntrada.Location = new System.Drawing.Point(60, 70);
			this.txtEntrada.Size = new System.Drawing.Size(341, 22);
			this.txtEntrada.Text = "";
			this.txtEntrada.BackColor = System.Drawing.SystemColors.Window;
			this.txtEntrada.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(1)), ((System.Byte)(72)), ((System.Byte)(178)));
			this.txtEntrada.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.txtEntrada.Font = new System.Drawing.Font("Tahoma", 9.00F, ((System.Drawing.FontStyle)System.Drawing.FontStyle.Regular), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			//
			// lblMsg
			//
			this.lblMsg.Name = "lblMsg";
			this.lblMsg.TabStop = false;
			this.lblMsg.TabIndex = 3;
			this.lblMsg.Location = new System.Drawing.Point(63, 55);
			this.lblMsg.Size = new System.Drawing.Size(337, 62);
			this.lblMsg.Text = "Mensagem..................";
			this.lblMsg.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(179)), ((System.Byte)(211)), ((System.Byte)(255)));
			this.lblMsg.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(64)));
			this.lblMsg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblMsg.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.lblMsg.Multiline = true;
			this.lblMsg.ReadOnly = true;
			this.lblMsg.Font = new System.Drawing.Font("Tahoma", 9.75F, ((System.Drawing.FontStyle)System.Drawing.FontStyle.Regular), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			//
			// imgInforma
			//
			this.imgInforma.Name = "imgInforma";
			this.imgInforma.Visible = false;
			this.imgInforma.Location = new System.Drawing.Point(13, 16);
			this.imgInforma.Size = new System.Drawing.Size(32, 32);
			this.imgInforma.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.imgInforma.BackColor = System.Drawing.SystemColors.Control;
			this.imgInforma.Image = ((System.Drawing.Image)(resources.GetObject("imgInforma.Image")));
			//
			// imgEntrada
			//
			this.imgEntrada.Name = "imgEntrada";
			this.imgEntrada.Visible = false;
			this.imgEntrada.Location = new System.Drawing.Point(13, 16);
			this.imgEntrada.Size = new System.Drawing.Size(32, 32);
			this.imgEntrada.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.imgEntrada.BackColor = System.Drawing.SystemColors.Control;
			this.imgEntrada.Image = ((System.Drawing.Image)(resources.GetObject("imgEntrada.Image")));
			//
			// imgExclama
			//
			this.imgExclama.Name = "imgExclama";
			this.imgExclama.Visible = false;
			this.imgExclama.Location = new System.Drawing.Point(13, 16);
			this.imgExclama.Size = new System.Drawing.Size(32, 32);
			this.imgExclama.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.imgExclama.BackColor = System.Drawing.SystemColors.Control;
			this.imgExclama.Image = ((System.Drawing.Image)(resources.GetObject("imgExclama.Image")));
			//
			// imgMsg
			//
			this.imgMsg.Name = "imgMsg";
			this.imgMsg.Visible = false;
			this.imgMsg.Location = new System.Drawing.Point(13, 16);
			this.imgMsg.Size = new System.Drawing.Size(32, 32);
			this.imgMsg.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.imgMsg.BackColor = System.Drawing.SystemColors.Control;
			this.imgMsg.Image = ((System.Drawing.Image)(resources.GetObject("imgMsg.Image")));
			this.imgMsg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			//
			// imgAvisa
			//
			this.imgAvisa.Name = "imgAvisa";
			this.imgAvisa.Visible = false;
			this.imgAvisa.Location = new System.Drawing.Point(13, 16);
			this.imgAvisa.Size = new System.Drawing.Size(32, 32);
			this.imgAvisa.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.imgAvisa.BackColor = System.Drawing.SystemColors.Control;
			this.imgAvisa.Image = ((System.Drawing.Image)(resources.GetObject("imgAvisa.Image")));
			//
			// imgErro
			//
			this.imgErro.Name = "imgErro";
			this.imgErro.Visible = false;
			this.imgErro.Location = new System.Drawing.Point(13, 16);
			this.imgErro.Size = new System.Drawing.Size(32, 32);
			this.imgErro.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.imgErro.BackColor = System.Drawing.SystemColors.Control;
			this.imgErro.Image = ((System.Drawing.Image)(resources.GetObject("imgErro.Image")));
			//
			// lblTitulo
			//
			this.lblTitulo.Name = "lblTitulo";
			this.lblTitulo.TabIndex = 1;
			this.lblTitulo.Location = new System.Drawing.Point(62, 24);
			this.lblTitulo.Size = new System.Drawing.Size(352, 18);
			this.lblTitulo.Text = "Título...............";
			this.lblTitulo.BackColor = System.Drawing.Color.Transparent;
			this.lblTitulo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(1)), ((System.Byte)(72)), ((System.Byte)(178)));
			this.lblTitulo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.lblTitulo.Font = new System.Drawing.Font("Tahoma", 12.00F, ((System.Drawing.FontStyle)System.Drawing.FontStyle.Bold), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			//
			// VSMensagem
			//
			this.ClientSize = new System.Drawing.Size(417, 161);
			this.Controls.Add(this.cmdBotao_6);
			this.Controls.Add(this.txtEntrada);
			this.Controls.Add(this.lblMsg);
			this.Controls.Add(this.imgInforma);
			this.Controls.Add(this.imgEntrada);
			this.Controls.Add(this.imgExclama);
			this.Controls.Add(this.imgMsg);
			this.Controls.Add(this.imgAvisa);
			this.Controls.Add(this.imgErro);
			this.Controls.Add(this.lblTitulo);
			this.Name = "VSMensagem";
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(179)), ((System.Byte)(211)), ((System.Byte)(255)));
			this.ForeColor = System.Drawing.SystemColors.ControlText;
			this.ShowInTaskbar = false;
			this.MinimizeBox = false;
			this.MaximizeBox = false;
			this.ControlBox = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Activated += new System.EventHandler(this.VSMensagem_Activated);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.VSMensagem_KeyPress);
			this.Text = "Mensagem";
			this.cmdBotao_6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmdBotao_6)).EndInit();
			this.txtEntrada.ResumeLayout(false);
			this.lblMsg.ResumeLayout(false);
			this.lblTitulo.ResumeLayout(false);
			this.ResumeLayout(false);
		}
		#endregion


		//=========================================================

		private void cmdBotao_ClickEvent(short Index, object sender, System.EventArgs e)
		{
			try
			{	// On Error GoTo Trata
				CMOD.T.OpcaoBotao = Index;
				CMOD.T.Resposta = Strings.Trim(txtEntrada.Text);
				Close();
			}
			catch
			{	// Trata:
			}
		}
		private void cmdBotao_6_ClickEvent(object sender, System.EventArgs e)
		{
			cmdBotao_ClickEvent(6, sender, e);
		}

		private void VSMensagem_Activated(object sender, System.EventArgs e)
		{
			cmdBotao[6].Visible = true;
			Interaction.Beep();
		}

		private void VSMensagem_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (KeyAscii==Keys.Return) SendKeys("{TAB}");
		}

	}
}
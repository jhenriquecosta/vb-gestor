VERSION 5.00
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TNAV001 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   6059
      Caption         =   "Credenciamento de Gráficas"
      Descricao       =   "Credencia gráficas no sistema"
      corFaixa        =   32768
      Icone           =   "TNAV001.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fra 
         Height          =   1890
         Index           =   1
         Left            =   75
         TabIndex        =   1
         Top             =   735
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   3334
         Altura          =   1905
         Caption         =   " Informações Gerais"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtNomeContrib 
            Height          =   480
            Left            =   0
            TabIndex        =   2
            Top             =   360
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
      End
   End
End
Attribute VB_Name = "TNAV001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

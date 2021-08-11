VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRDC101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TRDC101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   330
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Recpção"
      Height          =   1905
      Left            =   15
      TabIndex        =   8
      Top             =   1380
      Width           =   7440
      Begin MSComctlLib.ProgressBar Progresso 
         Height          =   225
         Left            =   75
         TabIndex        =   9
         Top             =   1620
         Visible         =   0   'False
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2445
         TabIndex        =   20
         Top             =   870
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Declaracões"
         Height          =   195
         Index           =   1
         Left            =   1110
         TabIndex        =   19
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Declaracões Rejeitadas"
         Height          =   195
         Left            =   720
         TabIndex        =   17
         Top             =   1380
         Width           =   1680
      End
      Begin VB.Label lblAceitas 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2445
         TabIndex        =   16
         Top             =   1125
         Width           =   45
      End
      Begin VB.Label lblRejeitadas 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2445
         TabIndex        =   15
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Declaracões Aceitas"
         Height          =   195
         Index           =   0
         Left            =   945
         TabIndex        =   14
         Top             =   1095
         Width           =   1440
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2445
         TabIndex        =   13
         Top             =   570
         Width           =   45
      End
      Begin VB.Label Agencia 
         AutoSize        =   -1  'True
         Caption         =   "Data Arquivo"
         Height          =   195
         Left            =   1425
         TabIndex        =   12
         Top             =   570
         Width           =   945
      End
      Begin VB.Label lblMunicipio 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2445
         TabIndex        =   11
         Top             =   315
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Municipio"
         Height          =   195
         Left            =   1710
         TabIndex        =   10
         Top             =   285
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe o caminho dos arquivos de DMS.dec"
      Height          =   720
      Left            =   15
      TabIndex        =   5
      Top             =   645
      Width           =   7440
      Begin VB.TextBox txtCamminhoRemessa 
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   6690
      End
      Begin VTOcx.cmdVISUAL cmdConsultaArquivo 
         Height          =   315
         Left            =   6885
         TabIndex        =   7
         Top             =   270
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
   Begin Cabecalho.cabVISUAL cabCabecalho 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1138
      Formulario      =   "CODIGO"
      Icone           =   "TRDC101.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   6525
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   926
      CorFundo        =   -2147483632
      CorFrente       =   -2147483633
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   5640
         TabIndex        =   4
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   6690
         TabIndex        =   1
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdReceber 
         Height          =   405
         Left            =   4515
         TabIndex        =   0
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   714
         Caption         =   "&Receber"
         Acao            =   3
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
   Begin VTOcx.grdVISUAL grdDocumentos 
      Height          =   3435
      Left            =   30
      TabIndex        =   18
      Top             =   3330
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   6059
      CorBorda        =   32768
      Caption         =   "Declaracões"
      CorTitulo       =   32768
      CorDica         =   0
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TRDC101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Arquivo   As Arquivo
Dim Path         As String
Private Sub cmdConsultaArquivo_Click()
    Path = Arquivo.ChamaArquivos(Dialogo, etaDeclaracao)
    txtCamminhoRemessa = Path
End Sub

Private Sub cmdLimpar_Click()
    lblMunicipio = "": lblAceitas = "": lblData = "": lblRejeitadas = "": lblTotal = ""
    grdDocumentos.ListItems.Clear
    txtCamminhoRemessa.SetFocus
End Sub

Private Sub cmdReceber_Click()
    Dim Dec As New ArquivoDeclaracao
    cmdLimpar_Click
    If Trim(txtCamminhoRemessa) <> "" Then
        Dec.CarregaDeclaracao txtCamminhoRemessa, grdDocumentos
        cmdReceber.Enabled = False
        lblMunicipio = Dec.Municipio
        lblData = Dec.DataArquivo
        lblTotal = Dec.TotalDeclaracoes
        lblAceitas = Dec.TotalAceitas
        lblRejeitadas = Dec.TotalRejeitadas
        DoEvents
        Informa "Leitura finalizada."
        cmdReceber.Enabled = True
        txtCamminhoRemessa.SetFocus
    Else
        Informa "Informe um arquivo."
        txtCamminhoRemessa.SetFocus
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabCabecalho.Exibir Bdados, Me.Name, App.Path
    rodRodape.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Set Arquivo = New Arquivo
    prepararGrid
End Sub

Private Sub prepararGrid()
    grdDocumentos.ColumnHeaders.Clear
    grdDocumentos.ColumnHeaders.Add , , "Contribuinte", 2000
    grdDocumentos.ColumnHeaders.Add , , "Periodo", 1000
    grdDocumentos.ColumnHeaders.Add , , "Tipo", 600
    grdDocumentos.ColumnHeaders.Add , , "Data", 1500
    grdDocumentos.ColumnHeaders.Add , , "Versão", 800
    grdDocumentos.ColumnHeaders.Add , , "STATUS", 800
End Sub


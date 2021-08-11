VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form TREC101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   7680
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   34
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TREC101.frx":0000
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
      Height          =   3795
      Left            =   15
      TabIndex        =   8
      Top             =   1380
      Width           =   7440
      Begin MSComctlLib.ProgressBar Progresso 
         Height          =   225
         Left            =   75
         TabIndex        =   9
         Top             =   3480
         Visible         =   0   'False
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblTotalAberto 
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
         Left            =   2460
         TabIndex        =   33
         Top             =   2280
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total Aberto"
         Height          =   195
         Index           =   2
         Left            =   1020
         TabIndex        =   32
         Top             =   2250
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total Baixado"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   31
         Top             =   1950
         Width           =   1380
      End
      Begin VB.Label lblTotalBaixado 
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
         Left            =   2460
         TabIndex        =   30
         Top             =   1965
         Width           =   45
      End
      Begin VB.Label lblLote 
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
         Left            =   2415
         TabIndex        =   29
         Top             =   870
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   28
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total de Registros Rejeitados"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   3075
         Width           =   2115
      End
      Begin VB.Label lblRejeitado 
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
         TabIndex        =   25
         Top             =   3090
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total de Registros Aceitos"
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   2805
         Width           =   1875
      End
      Begin VB.Label lblAceitos 
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
         TabIndex        =   23
         Top             =   2790
         Width           =   45
      End
      Begin VB.Label LblTotalRegistroRemessa 
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
         Left            =   2430
         TabIndex        =   22
         Top             =   2550
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total de Registros na Remessa"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   2535
         Width           =   2220
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data de geração do arquivo"
         Height          =   195
         Left            =   330
         TabIndex        =   20
         Top             =   1380
         Width           =   2010
      End
      Begin VB.Label LblDataPagamentoDOC 
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
         Left            =   2430
         TabIndex        =   19
         Top             =   1125
         Width           =   45
      End
      Begin VB.Label LblTotalArrecardado 
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
         Left            =   2430
         TabIndex        =   18
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total Arrecadado"
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   17
         Top             =   1665
         Width           =   1650
      End
      Begin VB.Label LblDtGeracaoArquivo 
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
         Left            =   2415
         TabIndex        =   16
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data de Pagamento"
         Height          =   195
         Index           =   0
         Left            =   915
         TabIndex        =   15
         Top             =   1095
         Width           =   1425
      End
      Begin VB.Label LblAgencia 
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
         Left            =   2415
         TabIndex        =   14
         Top             =   570
         Width           =   45
      End
      Begin VB.Label Agencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         Height          =   195
         Left            =   1755
         TabIndex        =   13
         Top             =   570
         Width           =   570
      End
      Begin VB.Label LblBanco 
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
         Left            =   2400
         TabIndex        =   12
         Top             =   195
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   1860
         TabIndex        =   11
         Top             =   195
         Width           =   435
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3465
         TabIndex        =   10
         Top             =   2445
         Width           =   45
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe o caminho dos arquivos de remessa.ret"
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
      Icone           =   "TREC101.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   7155
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
      Height          =   2055
      Left            =   30
      TabIndex        =   27
      Top             =   5250
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   3625
      CorBorda        =   32768
      Caption         =   "Valores"
      CorTitulo       =   32768
      CorDica         =   0
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TREC101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Arquivo   As Arquivo
Public Cabecalho As Header
Dim Path         As String
Private Sub cmdConsultaArquivo_Click()
    Path = Arquivo.ChamaArquivos(Dialogo, etaArrecadacao)
    txtCamminhoRemessa = Path
    If Path <> "" Then
        Arquivo.CarregaArquivos IIf(Trim(Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO")) = "CBR643", etcCBR643, etcRCB), Path, LblBanco, LblAgencia, LblDtGeracaoArquivo, LblTotalArrecardado, LblDataPagamentoDOC, LblTotalRegistroRemessa, Progresso, LblPercento, lblAceitos, lblRejeitado, lblLote, lblTotalBaixado, lblTotalAberto, grdDocumentos, True
        txtCamminhoRemessa.SetFocus
        Avisa "Leitura finalizada."
    End If
End Sub

Private Sub cmdLimpar_Click()
    LblBanco = "": LblAgencia = "": LblDtGeracaoArquivo = "": LblTotalArrecardado = ""
    LblDataPagamentoDOC = "": LblTotalRegistroRemessa = "": LblPercento = ""
    lblAceitos = "": lblRejeitado = "": lblLote = "": lblTotalBaixado = "": lblTotalAberto = ""
    txtCamminhoRemessa = ""
    grdDocumentos.ListItems.Clear
    txtCamminhoRemessa.SetFocus
End Sub

Private Sub cmdReceber_Click()
    If Trim(txtCamminhoRemessa) <> "" Then
        cmdReceber.Enabled = False
        Arquivo.CarregaArquivos IIf(Trim(Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO")) = "CBR643", etcCBR643, etcRCB), Path, LblBanco, LblAgencia, LblDtGeracaoArquivo, LblTotalArrecardado, LblDataPagamentoDOC, LblTotalRegistroRemessa, Progresso, LblPercento, lblAceitos, lblRejeitado, lblLote, lblTotalBaixado, lblTotalAberto, grdDocumentos
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
    grdDocumentos.ColumnHeaders.Add , , "Documento", 2500
    grdDocumentos.ColumnHeaders.Add , , "Inscricão", 2000
    grdDocumentos.ColumnHeaders.Add , , "Imposto", 1000
    grdDocumentos.ColumnHeaders.Add , , "Periodo", 900
    grdDocumentos.ColumnHeaders.Add , , "Valor", 1500
    grdDocumentos.ColumnHeaders.Add , , "Situacão", 2500
    If Trim(Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO")) = "CBR643" Then
'        grdDocumentos.ColumnHeaders(1).Width = 0
        grdDocumentos.ColumnHeaders(2).Width = 0
        grdDocumentos.ColumnHeaders(3).Width = 0
        grdDocumentos.ColumnHeaders(4).Width = 0
        DoEvents
    End If

End Sub


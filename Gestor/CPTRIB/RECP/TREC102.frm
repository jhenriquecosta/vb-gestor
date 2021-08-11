VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.3#0"; "VTControles.ocx"
Begin VB.Form TREC102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   1935
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
   ScaleHeight     =   1935
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   330
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Icone           =   "TREC102.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   1410
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   926
      CorFundo        =   -2147483632
      CorFrente       =   -2147483633
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   5490
         TabIndex        =   4
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   6600
         TabIndex        =   1
         Top             =   90
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdReceber 
         Height          =   405
         Left            =   4245
         TabIndex        =   0
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   714
         Caption         =   "&Corrigir"
         Acao            =   3
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
End
Attribute VB_Name = "TREC102"
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
        Arquivo.CorrigeArquivo txtCamminhoRemessa
        Informa "Correcão finalizada."
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
End Sub


VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR406 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TOBR405"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   975
      Left            =   30
      TabIndex        =   8
      Top             =   1620
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1720
      Altura          =   1905
      Caption         =   " Registro"
      CorTexto        =   0
      CorFaixa        =   12632256
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtData 
         Height          =   510
         Left            =   5700
         TabIndex        =   12
         Tag             =   "Periodo Final"
         Top             =   300
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   900
         Caption         =   "Data"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtFolha 
         Height          =   510
         Left            =   3750
         TabIndex        =   11
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         Caption         =   "Folha"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtLivro 
         Height          =   510
         Left            =   1950
         TabIndex        =   10
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         Caption         =   "Livro"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtReg 
         Height          =   510
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         Caption         =   "Número"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   915
      Left            =   30
      TabIndex        =   6
      Top             =   690
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1614
      Altura          =   1905
      Caption         =   " Termo de Lancamento"
      CorTexto        =   0
      CorFaixa        =   12632256
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtNumeroProcesso 
         Height          =   510
         Left            =   60
         TabIndex        =   7
         Tag             =   "Periodo Final"
         Top             =   300
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   900
         Caption         =   "Termo"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorFundo        =   14737632
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   3045
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   2055
         TabIndex        =   3
         Top             =   105
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         Caption         =   "&Imprimir Documento"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   4305
         TabIndex        =   2
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   1138
      Icone           =   "TOBR406.frx":0000
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1620
      Left            =   60
      TabIndex        =   4
      Top             =   6345
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   2858
      Caption         =   "Taxas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
   Begin VTOcx.cboVISUAL cboTipo 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Tag             =   "Documento"
      Top             =   2670
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   556
      Caption         =   "Documento"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Requerido       =   0   'False
   End
End
Attribute VB_Name = "TOBR406"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim String_Taxas As String
Dim Total_Taxas  As String
Dim NovaData As String
Private Sub cmd_Click(Index As Integer)
    Dim Cobranca As New VSCobranca
 
        Select Case Index
            Case 1
                Unload Me
            Case 2
                Avisa "Indisponível no momento."
        End Select
End Sub

Private Sub Form_Activate()
    Dim Sql As String
    Dim rs As VSRecordset
    cboTipo.Preencher Bdados, "SELECT TGE_CODIGO,TGE_NOME FROM VIS_DOCUMENTOS_DAT ORDER BY TGE_CODIGO", 1
    Sql = "SELECT TDA_DATA_INSCRICAO,TDA_NUM_PROCESSO,TDA_REGISTRO,TDA_LIVRO,TDA_FOLHA FROM TAB_DIVIDA_ATIVA " & _
        " WHERE TDA_TOC_COD_OBRIGACAO =" & Me.Tag
    If Bdados.AbreTabela(Sql, rs) Then
        txtData = "" & rs!TDA_DATA_INSCRICAO
        txtNumeroProcesso = "" & rs!TDA_NUM_PROCESSO
        txtReg = "" & rs!TDA_REGISTRO
        txtLivro = "" & rs!TDA_LIVRO
        txtFolha = "" & rs!TDA_FOLHA
    End If
    
End Sub


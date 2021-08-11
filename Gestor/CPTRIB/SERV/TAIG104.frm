VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAIG104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAIG104"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TAIG104.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAIG104.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   27
      Top             =   5730
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5565
         TabIndex        =   5
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   4395
         TabIndex        =   4
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6735
         TabIndex        =   6
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
   Begin ActiveTabs.SSActiveTabs TabGrafica 
      Height          =   4890
      Left            =   75
      TabIndex        =   8
      Top             =   720
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   8625
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TAIG104.frx":2465
      Images          =   "TAIG104.frx":24F3
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4470
         Left            =   -99969
         TabIndex        =   11
         Top             =   30
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   7885
         _Version        =   131082
         TabGuid         =   "TAIG104.frx":2BFB
         Begin VTOcx.fraVISUAL fra 
            Height          =   4275
            Index           =   0
            Left            =   1125
            TabIndex        =   12
            Top             =   90
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   7541
            Altura          =   1905
            Caption         =   " Estabelecimento Gráfico"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL TxtDescreden 
               Height          =   480
               Left            =   75
               TabIndex        =   26
               Top             =   3660
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   847
               Caption         =   "Descredenciamento"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtValidade 
               Height          =   480
               Left            =   3795
               TabIndex        =   25
               Top             =   3180
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   847
               Caption         =   "Validade"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCredenciamento 
               Height          =   480
               Left            =   2242
               TabIndex        =   24
               Top             =   3180
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   847
               Caption         =   "Credenciamento"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtSituacao 
               Height          =   480
               Left            =   75
               TabIndex        =   23
               Top             =   3180
               Width           =   2115
               _ExtentX        =   3731
               _ExtentY        =   847
               Caption         =   "Situação"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtIMDadoGrafica 
               Height          =   480
               Left            =   75
               TabIndex        =   22
               Tag             =   "Insc. Municipal"
               Top             =   300
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   847
               Caption         =   "Insc. Municipal"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtBairroGrafica 
               Height          =   480
               Left            =   75
               TabIndex        =   21
               Top             =   2220
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   847
               Caption         =   "Bairro"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtUFGrafica 
               Height          =   480
               Left            =   4680
               TabIndex        =   20
               Top             =   2700
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   847
               Caption         =   "UF"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtLogrGrafica 
               Height          =   480
               Left            =   1425
               TabIndex        =   19
               Top             =   1260
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   847
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNumeroGrafica 
               Height          =   480
               Left            =   4680
               TabIndex        =   18
               Top             =   1260
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   847
               Caption         =   "Nº"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNomeGrafica 
               Height          =   480
               Left            =   75
               TabIndex        =   17
               Top             =   780
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtTipoLogrGrafica 
               Height          =   480
               Left            =   75
               TabIndex        =   16
               Top             =   1260
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   847
               Caption         =   "Logradouro"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCnpjDadoGraf 
               Height          =   480
               Left            =   1890
               TabIndex        =   15
               Top             =   300
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   847
               Caption         =   "CNPJ"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   2
               AlinhamentoRotulo=   1
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCompGrafica 
               Height          =   480
               Left            =   75
               TabIndex        =   14
               Top             =   1740
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   847
               Caption         =   "Complemento"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCidadeGrafica 
               Height          =   480
               Left            =   75
               TabIndex        =   13
               Top             =   2700
               Width           =   4590
               _ExtentX        =   8096
               _ExtentY        =   847
               Caption         =   "Cidade"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4470
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   7885
         _Version        =   131082
         TabGuid         =   "TAIG104.frx":2C23
         Begin VTOcx.grdVISUAL grdPesquisa 
            Height          =   3000
            Left            =   90
            TabIndex        =   3
            Top             =   1485
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   5292
            CorBorda        =   32768
            Caption         =   "Resultado Pesquisa"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1365
            Left            =   105
            TabIndex        =   10
            Top             =   90
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   2408
            Altura          =   1905
            Caption         =   " Opções de Busca"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtImGrafica 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1980
               TabIndex        =   30
               Top             =   540
               Width           =   1755
            End
            Begin VTOcx.txtVISUAL txtNumCredenciamento 
               Height          =   480
               Left            =   255
               TabIndex        =   0
               Top             =   330
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   847
               Caption         =   "Nº Credenciamento"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtCgcGrafica 
               Height          =   480
               Left            =   3765
               TabIndex        =   1
               Top             =   360
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   847
               Caption         =   "CNPJ/CPF"
               Text            =   ""
               Formato         =   2
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNome 
               Height          =   480
               Left            =   240
               TabIndex        =   2
               Top             =   825
               Width           =   5940
               _ExtentX        =   10478
               _ExtentY        =   847
               Caption         =   "Nome"
               Text            =   ""
               AlinhamentoRotulo=   1
            End
            Begin VB.Label Label1 
               Caption         =   "Insc.Municipal"
               Height          =   225
               Left            =   2010
               TabIndex        =   31
               Top             =   330
               Width           =   1275
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1138
      Icone           =   "TAIG104.frx":2C4B
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   28
      Top             =   2895
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "TAIG104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Grafica As cGraficaAidf
Dim Contribuinte As cContribuinte
Public PessoaFisica As Boolean

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdPesquisa.ListItems.Clear
    TabGrafica.Tabs(1).Selected = True
    txtImGrafica.SetFocus
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    cabVISUAL.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Set Grafica = New cGraficaAidf
    Set Contribuinte = New cContribuinte
    Screen.MousePointer = 0
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
       txtIMDadoGrafica.Formato = formNenhum
    End If
    txtCgcGrafica.Formato = formDocumento
End Sub

Private Sub cmdBuscar_Click()
Dim rs As VSRecordset
Dim Sql As String
    If Me.Tag = "TAID201" Then
        Contribuinte.PreencherGrid grdPesquisa, txtNome
    Else
        Grafica.PreencherGrid grdPesquisa, txtNumCredenciamento, txtNome, txtImGrafica, txtCgcGrafica
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Grafica = Nothing
    Set Contribuinte = Nothing
End Sub

Private Sub grdPesquisa_DblClick()
    If grdPesquisa.SelectedItem Is Nothing Then Exit Sub
    If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIMDadoGrafica.Formato = formNenhum
    End If
    txtIMDadoGrafica = grdPesquisa.SelectedItem.SubItems(1)
    txtIMDadoGrafica_lostfocus
End Sub

Private Sub txtIMDadoGrafica_lostfocus()
    Dim NomeGraf As String, TipoLogrGraf As String, LogrGraf As String, NumeroGraf As String, CompGraf As String, _
          BairroGraf As String, CepGraf As String, MunicGraf As String, UFGraf As String, DocumentoGraf As String
    If Trim(txtIMDadoGrafica) = "" Then Exit Sub
        LimpaCamposGrafica
        If Contribuinte.BuscarContribuinte(txtIMDadoGrafica, NomeGraf, TipoLogrGraf, LogrGraf, NumeroGraf, CompGraf, _
                BairroGraf, CepGraf, MunicGraf, UFGraf, DocumentoGraf) Then
            TabGrafica.Tabs(2).Selected = True
            txtNomeGrafica = NomeGraf
            txtCnpjDadoGraf = DocumentoGraf
            txtTipoLogrGrafica = TipoLogrGraf
            txtLogrGrafica = LogrGraf
            txtNumeroGrafica = NumeroGraf
            txtCompGrafica = CompGraf
            txtBairroGrafica = BairroGraf
            txtCidadeGrafica = MunicGraf
            txtUFGrafica = UFGraf
            If Grafica.Buscar(grdPesquisa.SelectedItem) Then
                txtValidade = Grafica.Validade
                txtCredenciamento = Grafica.DataInicio
                TxtDescreden = Grafica.DataDescredenciamento
                txtSituacao = IIf(Grafica.Situacao = 0, "CREDENCIADA", "DESCREDENCIADA")
            End If
        End If
End Sub

Private Sub LimpaCamposGrafica()
    txtNomeGrafica = ""
    txtTipoLogrGrafica = ""
    txtLogrGrafica = ""
    txtNumeroGrafica = ""
    txtCompGrafica = ""
    txtBairroGrafica = ""
    txtCidadeGrafica = ""
    txtUFGrafica = ""
    txtCnpjDadoGraf = ""
End Sub

Private Sub txtImGrafica_LostFocus()
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtImGrafica = Imposto.FormataInscricao(txtImGrafica, InscContrib)
    End If
End Sub


VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TAIG401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAIG401"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   49
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAIG401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   46
      Top             =   6300
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   7095
         TabIndex        =   10
         Top             =   120
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         Caption         =   "&Imprimir AIDF"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8760
         TabIndex        =   11
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9885
         TabIndex        =   12
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs TabConsulta 
      Height          =   5475
      Left            =   60
      TabIndex        =   9
      Top             =   735
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   9657
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
      Tabs            =   "TAIG401.frx":2123
      Images          =   "TAIG401.frx":21AE
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5055
         Left            =   -99969
         TabIndex        =   16
         Top             =   30
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   8916
         _Version        =   131082
         TabGuid         =   "TAIG401.frx":28B6
         Begin VTOcx.fraFUTURO fraFUTURO2 
            Height          =   5040
            Left            =   15
            TabIndex        =   17
            Top             =   15
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   8890
            Caption         =   "AIDF:"
            Descricao       =   "Informações gerais relacionado ao AIDF emitido"
            corFaixa        =   32768
            Icone           =   "TAIG401.frx":28DE
            Ocultavel       =   0   'False
            Altura          =   1905
            Begin VTOcx.fraVISUAL fraVISUAL1 
               Height          =   960
               Left            =   90
               TabIndex        =   39
               Top             =   4005
               Width           =   10755
               _ExtentX        =   18971
               _ExtentY        =   1693
               Altura          =   1905
               Caption         =   " Sequencia de Notas"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Enabled         =   0   'False
               Begin VTOcx.cboVISUAL cboEspecie 
                  Height          =   510
                  Left            =   105
                  TabIndex        =   45
                  Top             =   330
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   900
                  Caption         =   "Espécie"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
               Begin VTOcx.cboVISUAL cboSerie 
                  Height          =   510
                  Left            =   2847
                  TabIndex        =   44
                  Top             =   330
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   900
                  Caption         =   "Série/Sub-Série"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
               Begin VTOcx.txtVISUAL txtBlocos 
                  Height          =   480
                  Left            =   5916
                  TabIndex        =   43
                  Top             =   360
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   847
                  Caption         =   "Total de Blocos"
                  Text            =   ""
                  Enabled         =   0   'False
                  Restricao       =   2
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtInicio 
                  Height          =   480
                  Left            =   7563
                  TabIndex        =   42
                  Top             =   360
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   847
                  Caption         =   "Nota Inicial"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtFim 
                  Height          =   480
                  Left            =   9210
                  TabIndex        =   41
                  Top             =   360
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   847
                  Caption         =   "Nota Final"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtNotaBloco 
                  Height          =   480
                  Left            =   4314
                  TabIndex        =   40
                  Top             =   360
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   847
                  Caption         =   "Notas p/ Bloco"
                  Text            =   ""
                  Enabled         =   0   'False
                  Restricao       =   2
                  AlinhamentoRotulo=   1
               End
            End
            Begin VTOcx.fraVISUAL fra 
               Height          =   3255
               Index           =   0
               Left            =   5490
               TabIndex        =   28
               Top             =   720
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   5741
               Altura          =   1905
               Caption         =   " Estabelecimento Gráfico"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtCidadeGrafica 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   38
                  Top             =   2700
                  Width           =   4590
                  _ExtentX        =   8096
                  _ExtentY        =   847
                  Caption         =   "Cidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtCompGrafica 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   37
                  Top             =   1740
                  Width           =   5250
                  _ExtentX        =   9260
                  _ExtentY        =   847
                  Caption         =   "Complemento"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtCgcGrafica 
                  Height          =   480
                  Left            =   1890
                  TabIndex        =   36
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
               Begin VTOcx.txtVISUAL txtTipoLogrGrafica 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   35
                  Top             =   1260
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   847
                  Caption         =   "Logradouro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtNomeGrafica 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   34
                  Top             =   780
                  Width           =   5250
                  _ExtentX        =   9260
                  _ExtentY        =   847
                  Caption         =   "Nome"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtNumeroGrafica 
                  Height          =   480
                  Left            =   4680
                  TabIndex        =   33
                  Top             =   1260
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   847
                  Caption         =   "Nº"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtLogrGrafica 
                  Height          =   480
                  Left            =   1425
                  TabIndex        =   32
                  Top             =   1260
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   847
                  Caption         =   ""
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtUFGrafica 
                  Height          =   480
                  Left            =   4680
                  TabIndex        =   31
                  Top             =   2700
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   847
                  Caption         =   "UF"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtBairroGrafica 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   30
                  Top             =   2220
                  Width           =   5250
                  _ExtentX        =   9260
                  _ExtentY        =   847
                  Caption         =   "Bairro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtImGrafica 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   29
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
            End
            Begin VTOcx.fraVISUAL fra 
               Height          =   3255
               Index           =   1
               Left            =   90
               TabIndex        =   18
               Top             =   720
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   5741
               Altura          =   1905
               Caption         =   " Contribuinte"
               CorTexto        =   16777215
               CorFaixa        =   32768
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtCgc 
                  Height          =   480
                  Left            =   1890
                  TabIndex        =   47
                  Top             =   300
                  Width           =   3435
                  _ExtentX        =   6059
                  _ExtentY        =   847
                  Caption         =   "CNPJ"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtLogr 
                  Height          =   480
                  Left            =   1425
                  TabIndex        =   27
                  Top             =   1260
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   847
                  Caption         =   ""
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtNumero 
                  Height          =   480
                  Left            =   4680
                  TabIndex        =   26
                  Top             =   1260
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   847
                  Caption         =   "Nº"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtNomeContrib 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   25
                  Top             =   780
                  Width           =   5250
                  _ExtentX        =   9260
                  _ExtentY        =   847
                  Caption         =   "Nome"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtTipoLogr 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   24
                  Top             =   1260
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   847
                  Caption         =   "Logradouro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtUF 
                  Height          =   480
                  Left            =   4680
                  TabIndex        =   23
                  Top             =   2700
                  Width           =   645
                  _ExtentX        =   1138
                  _ExtentY        =   847
                  Caption         =   "UF"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtComplemento 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   22
                  Top             =   1740
                  Width           =   5250
                  _ExtentX        =   9260
                  _ExtentY        =   847
                  Caption         =   "Complemento"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtBairro 
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
               Begin VTOcx.txtVISUAL txtCidade 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   20
                  Top             =   2700
                  Width           =   4590
                  _ExtentX        =   8096
                  _ExtentY        =   847
                  Caption         =   "Cidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtIm 
                  Height          =   480
                  Left            =   75
                  TabIndex        =   19
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
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5055
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   8916
         _Version        =   131082
         TabGuid         =   "TAIG401.frx":31B8
         Begin VTOcx.fraFUTURO fraFUTURO1 
            Height          =   5040
            Left            =   15
            TabIndex        =   15
            Top             =   15
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   8890
            Caption         =   "Consulta de Emissão de AIDG"
            Descricao       =   "Informações opcionais para busca de AIDG"
            corFaixa        =   32768
            Icone           =   "TAIG401.frx":31E0
            Ocultavel       =   0   'False
            Altura          =   1905
            Begin VTOcx.cboVISUAL cboDoc 
               Height          =   510
               Left            =   2910
               TabIndex        =   51
               Top             =   675
               Width           =   6525
               _ExtentX        =   11509
               _ExtentY        =   900
               Caption         =   "Documento"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cmdVISUAL cmdBuscar 
               Height          =   375
               Left            =   9510
               TabIndex        =   50
               Top             =   1785
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               Caption         =   "&Buscar"
               Acao            =   5
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   315
               Left            =   2355
               TabIndex        =   2
               TabStop         =   0   'False
               Top             =   1350
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtNumNota 
               Height          =   480
               Left            =   7800
               TabIndex        =   7
               Top             =   1680
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   847
               Caption         =   "Nº Nota"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL TxtPeriodo1 
               Height          =   480
               Left            =   2940
               TabIndex        =   5
               Top             =   1680
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   847
               Caption         =   "Periodo Inicial"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtNumAidf 
               Height          =   480
               Left            =   705
               TabIndex        =   0
               Top             =   690
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   847
               Caption         =   "N° AIDF"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtContribuinte 
               Height          =   480
               Left            =   2925
               TabIndex        =   3
               Top             =   1200
               Width           =   6510
               _ExtentX        =   11483
               _ExtentY        =   847
               Caption         =   "Contribuinte"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtDtAutorizacao 
               Height          =   480
               Left            =   690
               TabIndex        =   4
               Top             =   1680
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   847
               Caption         =   "Dt. Autorização"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtIMBusca 
               Height          =   480
               Left            =   705
               TabIndex        =   1
               Top             =   1170
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   847
               Caption         =   "IM"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AgruparValores  =   0   'False
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL TxtPeriodo2 
               Height          =   480
               Left            =   5550
               TabIndex        =   6
               Top             =   1680
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   847
               Caption         =   "Periodo Final"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.grdVISUAL grdNotas 
               Height          =   2610
               Left            =   135
               TabIndex        =   8
               Top             =   2355
               Width           =   10650
               _ExtentX        =   18785
               _ExtentY        =   4604
               CorBorda        =   32768
               Caption         =   "Notas"
               CorTitulo       =   32768
               CorCaption      =   16777215
               CorDica         =   32768
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   1138
      Icone           =   "TAIG401.frx":34FA
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Height          =   195
      Left            =   0
      TabIndex        =   48
      Top             =   -15
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Menu mnuAidf 
      Caption         =   "AIDF"
      Visible         =   0   'False
      Begin VB.Menu mnuEmitir 
         Caption         =   "Emitir"
      End
   End
End
Attribute VB_Name = "TAIG401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Contribuinte As cContribuinte
Dim Grafica As cGraficaAidf
Dim NotaAidf As cNotaAidf

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdBuscar_Click()
    With NotaAidf
        If .PreencherGrid(grdNotas, txtNumAidf, txtDtAutorizacao, txtIMBusca, "", TxtPeriodo1, TxtPeriodo2, txtNumNota, CStr(cboDoc.Coluna(1).Valor)) = False Then
            Avisa "Nenhuma nota encontrada."
            txtContribuinte.SetFocus
        End If
    End With
End Sub

Private Sub cmdImprimir_Click()
    If grdNotas.SelectedItem Is Nothing Then Exit Sub
    NotaAidf.Imprimir grdNotas.SelectedItem
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIMBusca
End Sub

Private Sub Form_Activate()
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then txtIMBusca.Formato = formNenhum
End Sub

Private Sub mnuEmitir_Click()
    NotaAidf.Imprimir Util.ParseString(mnuEmitir.Caption, "AIDF", 2)
End Sub

Private Sub txtIm_LostFocus()
    Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
          BairroContr As String, CepContr As String, MunicContr As String, UFContr As String, DocumentoContr As String
    
    If Trim(txtIm) = "" Then Exit Sub
    With Contribuinte
        If .BuscarContribuinte(txtIm, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, _
            BairroContr, CepContr, MunicContr, UFContr, DocumentoContr) Then
                txtNomeContrib = NomeContrib
                txtCgc = DocumentoContr
                txtTipoLogr = TipoLogrContr
                txtLogr = LogrContr
                txtNumero = NumeroContr
                txtComplemento = CompContri
                txtBairro = BairroContr
                txtCidade = MunicContr
                txtUF = UFContr
        End If
    End With
End Sub

Private Sub txtIMBusca_LostFocus()
 Dim NomeContriBusca As String
    If txtIMBusca <> "" Then
        Contribuinte.BuscarContribuinte txtIMBusca, NomeContriBusca
        txtContribuinte = NomeContriBusca
    Else
        txtContribuinte = ""
    End If
End Sub

Private Sub txtImGrafica_LostFocus()
    Dim NomeGraf As String, TipoLogrGraf As String, LogrGraf As String, NumeroGraf As String, CompGraf As String, _
          BairroGraf As String, CepGraf As String, MunicGraf As String, UFGraf As String, DocumentoGraf As String
    If Trim(txtImGrafica) = "" Then Exit Sub
    With Grafica
        LimpaCamposGrafica
        If Contribuinte.BuscarContribuinte(txtImGrafica, NomeGraf, TipoLogrGraf, LogrGraf, NumeroGraf, CompGraf, _
                BairroGraf, CepGraf, MunicGraf, UFGraf, DocumentoGraf) Then
            txtNomeGrafica = NomeGraf
            txtCgcGrafica = DocumentoGraf
            txtTipoLogrGrafica = TipoLogrGraf
            txtLogrGrafica = LogrGraf
            txtNumeroGrafica = NumeroGraf
            txtCompGrafica = CompGraf
            txtBairroGrafica = BairroGraf
            txtCidadeGrafica = MunicGraf
            txtUFGrafica = UFGraf
        End If
    End With
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdNotas.ListItems.Clear
    TabConsulta.Tabs(1).Selected = True
    txtNumAidf.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Grafica = New cGraficaAidf
    Set Contribuinte = New cContribuinte
    Set NotaAidf = New cNotaAidf
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    NotaAidf.PreencherCboEspecie cboEspecie
    NotaAidf.PreencherCboSerie cboSerie
    If AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        txtIm.Formato = formNenhum
        txtImGrafica.Formato = formNenhum
    End If
    cboDoc.Preencher Bdados, "select * from vis_tipo_impressão_doc"
    cboDoc.SetarLinha 2, 1
    cboDoc.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Grafica = Nothing
    Set Contribuinte = Nothing
    Set NotaAidf = Nothing
End Sub

Private Sub grdNotas_dblClick()
    If grdNotas.SelectedItem Is Nothing Then Exit Sub
    With NotaAidf
    LimpaTodosCampos
        If .Buscar(grdNotas.SelectedItem) Then
            TabConsulta.Tabs(2).Selected = True
            fraFUTURO2.Caption = "Documento Nº" & " " & grdNotas.SelectedItem
            txtIm = .ImContribuinte
            txtIm_LostFocus
            txtImGrafica = .ImGrafica
            txtImGrafica_LostFocus
            txtInicio = .NotaInicial
            txtFim = .NotaFinal
            txtBlocos = .TotalBlocos
            cboEspecie.SetarLinha .TipoAidf, 1
            cboSerie.SetarLinha .Serie, 1
            txtNotaBloco = ((txtFim + 1) - txtInicio) / txtBlocos
        End If
    End With
End Sub

Private Sub grdNotas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And grdNotas.ListItems.Count > 0 Then
        mnuEmitir.Caption = "Emitir AIDF " & grdNotas.SelectedItem
        Me.PopupMenu mnuAidf
    End If
End Sub


Private Sub txtFim_LostFocus()
    If Trim(txtInicio) <> "" And Trim(txtFim) <> "" Then
        If ((CDbl(txtFim) - CDbl(txtInicio)) / 50) < 1 Then
            Util.Avisa "Quantidade insuficiente de notas."
            txtFim = ""
            Exit Sub
        End If
        txtBlocos = Fix((CDbl(txtFim) - CDbl(txtInicio) + 1) / 50)
    End If
End Sub

Private Sub txtInicio_LostFocus()
    If Trim(txtInicio) <> "" And Trim(txtFim) <> "" Then
        If ((CDbl(txtFim) - CDbl(txtInicio)) / 50) < 1 Then
            Util.Avisa "Quantidade insuficiente de notas."
            txtInicio = ""
            Exit Sub
        End If
        txtBlocos = Fix((CDbl(txtFim) - CDbl(txtInicio) + 1) / 50)
    End If
End Sub

Private Sub LimpaTodosCampos()
        txtIm = ""
        txtImGrafica = ""
        txtInicio = ""
        txtFim = ""
        txtBlocos = ""
        cboEspecie.ListIndex = -1
        cboSerie.ListIndex = -1
        txtNotaBloco = ""
        LimpaCamposContribuinte
        LimpaCamposGrafica
End Sub

Private Sub LimpaCamposContribuinte()
    txtNomeContrib = ""
    txtCgc = ""
    txtTipoLogr = ""
    txtLogr = ""
    txtNumero = ""
    txtComplemento = ""
    txtBairro = ""
    txtCidade = ""
    txtUF = ""
End Sub
Private Sub LimpaCamposGrafica()
    txtNomeGrafica = ""
    txtCgcGrafica = ""
    txtTipoLogrGrafica = ""
    txtLogrGrafica = ""
    txtNumeroGrafica = ""
    txtCompGrafica = ""
    txtBairroGrafica = ""
    txtCidadeGrafica = ""
    txtUFGrafica = ""
End Sub

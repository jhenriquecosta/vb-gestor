VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form THOM101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuinte"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabNotificacao 
      Height          =   4005
      Left            =   105
      TabIndex        =   14
      Top             =   2415
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   7064
      _Version        =   131082
      TabCount        =   5
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
      Tabs            =   "THOM101.frx":0000
      Images          =   "THOM101.frx":0127
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   3585
         Left            =   -99969
         TabIndex        =   41
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6324
         _Version        =   131082
         TabGuid         =   "THOM101.frx":0DC8
         Begin VB.TextBox txtTexo2 
            Appearance      =   0  'Flat
            Height          =   1515
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   1980
            Visible         =   0   'False
            Width           =   9060
         End
         Begin VB.TextBox txtTermo 
            Appearance      =   0  'Flat
            Height          =   1500
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   255
            Width           =   9060
         End
         Begin VB.Label Label2 
            Caption         =   "Observação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   45
            Top             =   1770
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Fundamentação Legal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   45
            Width           =   2325
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   3585
         Left            =   -99969
         TabIndex        =   31
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6324
         _Version        =   131082
         TabGuid         =   "THOM101.frx":0DF0
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            Height          =   3405
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   90
            Width           =   9060
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   3585
         Left            =   -99969
         TabIndex        =   18
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6324
         _Version        =   131082
         TabGuid         =   "THOM101.frx":0E18
         Begin VTOcx.txtVISUAL txtObservacao 
            Height          =   285
            Left            =   210
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   2940
            Visible         =   0   'False
            Width           =   8190
            _ExtentX        =   14446
            _ExtentY        =   503
            Caption         =   "Observação"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtFatorGerador 
            Height          =   285
            Left            =   105
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2430
            Visible         =   0   'False
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   503
            Caption         =   "Fato Gerador"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtDataVencimento 
            Height          =   285
            Left            =   4305
            TabIndex        =   34
            Top             =   540
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   503
            Caption         =   "Vencimento do Auto de Infração"
            Text            =   ""
            Formato         =   0
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtData 
            Height          =   285
            Left            =   4905
            TabIndex        =   33
            Top             =   210
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   503
            Caption         =   "Data do Auto de Infração"
            Text            =   ""
            Formato         =   0
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtAgravante 
            Height          =   300
            Left            =   5865
            TabIndex        =   37
            Top             =   1545
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   529
            Caption         =   "Agravante(%)"
            Text            =   ""
            Enabled         =   0   'False
            Formato         =   5
            Restricao       =   3
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtValorPago 
            Height          =   300
            Left            =   4380
            TabIndex        =   38
            Top             =   1920
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   529
            Caption         =   "Valor Total do Auto de Infração"
            Text            =   ""
            Enabled         =   0   'False
            TipoLetras      =   0
            Formato         =   5
            Restricao       =   3
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   285
            Left            =   5010
            TabIndex        =   36
            Top             =   1200
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   503
            Caption         =   "Valor Total de Infrações"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtValorTotalDebito 
            Height          =   285
            Left            =   5385
            TabIndex        =   35
            Top             =   870
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
            Caption         =   "Valor Total Lançado"
            Text            =   ""
            Enabled         =   0   'False
            Formato         =   5
            Restricao       =   3
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3585
         Index           =   1
         Left            =   -99969
         TabIndex        =   15
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6324
         _Version        =   131082
         TabGuid         =   "THOM101.frx":0E40
         Begin VTOcx.cmdVISUAL cmdConfirmar 
            Height          =   330
            Left            =   7860
            TabIndex        =   19
            Top             =   3165
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            Caption         =   "Confirmar"
            Acao            =   1
         End
         Begin VTOcx.txtVISUAL txtValorAgravante 
            Height          =   300
            Left            =   4035
            TabIndex        =   20
            Top             =   3180
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   529
            Caption         =   "Agravante da Infração (%)"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.grdVISUAL grdInfra 
            Height          =   3315
            Left            =   60
            TabIndex        =   21
            Top             =   75
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   5847
            CorFundo        =   -2147483638
            Caption         =   "Infrações"
            CorTitulo       =   32768
            CorCaption      =   -2147483629
            CorDica         =   -2147483627
            OcultarRodape   =   -1  'True
            CheckBox        =   -1  'True
            Ordenavel       =   0   'False
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3585
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6324
         _Version        =   131082
         TabGuid         =   "THOM101.frx":0E68
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Selecionar Todos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   17
            Top             =   4680
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   765
            Left            =   45
            TabIndex        =   22
            Top             =   2760
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   1349
            Altura          =   1905
            Caption         =   " Valores"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtOriginal 
               Height          =   300
               Left            =   6825
               TabIndex        =   26
               Top             =   375
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   529
               Caption         =   "Valor Total"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtJuros 
               Height          =   300
               Left            =   4725
               TabIndex        =   25
               Top             =   360
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   529
               Caption         =   "Juros"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtAtualizacao 
               Height          =   300
               Left            =   210
               TabIndex        =   23
               Top             =   360
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               Caption         =   "Atualização"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtMulta 
               Height          =   300
               Left            =   2670
               TabIndex        =   24
               Top             =   360
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               Caption         =   "Multa"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   3
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.grdVISUAL lstParcelas 
            Height          =   2970
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   5239
            CorFundo        =   -2147483638
            CorTitulo       =   32768
            CorCaption      =   -2147483629
            CorDica         =   -2147483627
            OcultarRodape   =   -1  'True
            CheckBox        =   -1  'True
            Ordenavel       =   0   'False
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1138
      Icone           =   "THOM101.frx":0E90
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   6450
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1005
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   6420
         TabIndex        =   7
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   7410
         TabIndex        =   8
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8430
         TabIndex        =   9
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   11
      Top             =   -525
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   10
      Top             =   -360
      Width           =   375
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1710
      Left            =   105
      TabIndex        =   27
      Top             =   675
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   3016
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   1345
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   0
         CorFrente       =   0
         CorFundo        =   16777152
      End
      Begin VTOcx.cboVISUAL CboTipoAuto 
         Height          =   315
         Left            =   435
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   -420
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         Caption         =   "Tipo de Auto"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorFundo        =   16777215
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   360
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   1005
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1035
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   503
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   285
         Left            =   735
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1350
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   503
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.txtVISUAL txtInscricao 
         Height          =   285
         Left            =   750
         TabIndex        =   1
         Tag             =   "Inscrição Cadastral"
         Top             =   720
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   503
         Caption         =   "Inscricao"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   285
         Left            =   3510
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   720
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   285
         Left            =   5250
         TabIndex        =   2
         Top             =   735
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         CorFundo        =   -2147483644
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   285
         Left            =   8685
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   735
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
      End
   End
End
Attribute VB_Name = "THOM101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim MaxCotas As Byte
Dim CodImp As String
Dim Cgc As String
Dim EnderecoContrib As String
Dim CodPagamento As String
Dim ListaDocs As String
Dim Marcou As Boolean

Dim Valor As Double
Dim Multa As Double
Dim Juros As Double
Dim Saldo As Double
Dim Atualizacao As Double
Dim Desconto As Double

Dim Periodo As String
Dim Tributo As String
Dim DataVenc As String
Dim CodAuto As String

Private Type Infracao
    CodigoInfracao  As String
    Percentual      As String
End Type
Private Valores() As Infracao

Private Sub CboTipoAuto_Click()
    If CboTipoAuto.Coluna(1).Valor = 2 Then
        cmdBuscar.Visible = False
        tabNotificacao.Tabs(2).Selected = True
        tabNotificacao.Tabs(1).Enabled = False
        txtValorTotalDebito = "0,00"
        cboImposto.Visible = False
    Else
        cboImposto.Visible = True
        tabNotificacao.Tabs(1).Enabled = True
        tabNotificacao.Tabs(1).Selected = True
        cmdBuscar.Visible = True
    End If
End Sub

Private Sub Check2_Click()
    Dim Item As Integer
    grdInfra.MarcarTodos Check2
    Pega_Valores
    
End Sub

Private Sub cmdBuscar_Click()
    Dim Conta As New ContaCorrente
    Dim Obrig As New Obrigacao
    
    If cboImposto.ListIndex = -1 Then
        Avisa "Selecione Tributo."
        cboImposto.SetFocus
        Exit Sub
    End If
    If txtImovel = "" And txtInscricao = "" Or txtInscricao <> "" And txtImovel <> "" Then
        Avisa "Informe " & txtImovel.Caption & " ou " & txtInscricao.Caption
        txtInscricao.SetFocus
        Exit Sub
    End If

    If txtInscricao <> "" Then
        Conta.ExecutaAtualizacao txtInscricao, etiContribuinte, False, , , Date, , , , , CStr(cboImposto.Coluna(0).Valor)
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, txtInscricao, CStr(cboImposto.Coluna(0).Valor), , etlNaoPagos, , etiContribuinte, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Util.Avisa "Consulta sem resultados."
        End If
    Else
        If Aplicacoes.municipio = "PETROLINA" Then
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(50), "TOC_DESCONTO", _
                "TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
        End If
        Conta.ExecutaAtualizacao txtImovel, etiImovel, , , , Date, , , , , CStr(cboImposto.Coluna(0).Valor)
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, txtImovel, CStr(cboImposto.Coluna(0).Valor), , etlNaoPagos, , etiImovel, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Util.Avisa "Consulta sem resultados."
        End If
    End If
End Sub

Private Sub cmdCancela_Click()
    Dim i As Integer
    Edita.LimpaCampos Me
    Screen.MousePointer = 0
    cmdParcela.Enabled = True
    Configura_Valores
    
    For i = 1 To grdInfra.ListItems.Count
        grdInfra.ListItems(i).Checked = False
    Next

    
    
End Sub
Private Sub Imprimir()
    Dim Rpt As New VSRelatorio
    Dim CondRelatorio As String
    
    With Rpt
        
        If Not .DefinirArquivo(Bdados, App.Path & "\TAutoInfracao.rpt") Then Exit Sub
        .Selecao = "{TAB_AUTO_INFRACAO.TAI_COD_AUTO} = '" & CodAuto & "'"
'        .Formulas "VT_ESTADO", Temp.PegaParametro(Bdados, "ESTADO")
'        .Formulas "VT_PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
'        .Formulas "VT_SECRETARIA", Temp.PegaParametro(Bdados, "SEMFAZ")
'        .Formulas "VT_SETOR", Temp.PegaParametro(Bdados, "GAF")
        .Visualizar
    End With
End Sub

Private Sub cmdConfirmar_Click()
    If grdInfra.ListItems.Count = 0 Then Exit Sub
    ReDim Preserve Valores(1 To grdInfra.ListItems.Count) As Infracao
    Valores(grdInfra.SelectedItem.Index).CodigoInfracao = grdInfra.SelectedItem
    Valores(grdInfra.SelectedItem.Index).Percentual = txtValorAgravante
    Pega_Agravante
    If txtValorAgravante <> "" And txtValorAgravante <> "0,00" Then
        Avisa "Operação concluída com sucesso."
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub
Public Function BuscaTexto(NomeTexto As String) As String
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = '" & NomeTexto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        BuscaTexto = "" & rs!TPT_TEXTO
    End If
End Function
Private Sub Pega_Valores()
    Dim i        As Integer
    Dim Infracao As Double
    
    Infracao = 0
    For i = 1 To grdInfra.ListItems.Count
        If grdInfra.ListItems(i).Checked Then
            Infracao = Infracao + CDbl(grdInfra.ListItems(i).SubItems(2))
        End If
    Next
    txtValor = Infracao * CDbl(TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ","))
    
    Pega_Agravante
End Sub

Private Sub cmdParcela_Click()
    Dim Campos      As String
    Dim Valore     As String
    Dim Condicao    As String
    Dim Obrig       As New Obrigacao
    Dim i           As Integer
    Dim Base        As Integer
    Dim Gravou      As Boolean
    Dim Marcou      As Boolean
    Dim Base2       As Integer
    Dim Conta       As New ContaCorrente
    Const StatusAtivo = 1
    For i = 1 To grdInfra.ListItems.Count
        If grdInfra.ListItems(i).Checked Then
            Marcou = True
            Exit For
        End If
    Next
    If Not Marcou And CboTipoAuto.Coluna(1).Valor = 2 Then Avisa "Selecione Infração.": Exit Sub
    If CboTipoAuto.Coluna(1).Valor = 1 Then
        cboImposto.Tag = "Tributo"
    Else
        cboImposto.Tag = ""
    End If
    If Not CriticaCampos(Me) Then Exit Sub
      
    If CDbl(Nvl(Trim(txtOriginal), 0)) = 0 Then
        CodPagamento = Obrig.CriaObrigacao(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_AUTO_INFRACAO)), Format(Month(txtData), "00") & Year(txtData), Format(Month(txtData), "00") & Year(txtData), txtInscricao, txtValorPago, etsCreditoOriginalAberto, etsCriaNova, txtDataVencimento, , , , , , , , , , etiContribuinte)
    Else
        Dim ValoresExt As String
        CodPagamento = Conta.GeraCodPagamento(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_AUTO_INFRACAO)))
        Conta.GeraPagamento txtInscricao, txtImovel, _
                Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_AUTO_INFRACAO)), _
                  Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), _
                  txtDataVencimento, CDbl(txtValorPago), 0, 0, CodPagamento, 0, 0, 0, _
                  , EtcAutoInfracao
        Campos = "TPE_INSCRICAO, TPE_COD_PAGAMENTO_EXTRATO, TPE_TGT_COD_PAGAMENTO, " & _
                "TPE_TIP_COD_IMPOSTO,TPE_SUB_VALOR,TPE_TIPO_DOCUMENTO,TPE_SUB_PERIODO"
        For i = 1 To grdInfra.ListItems.Count
            If grdInfra.ListItems(i).Checked Then
                With grdInfra.ListItems(i)
                    ValoresExt = Bdados.PreparaValor(Trim(.SubItems(1)), CodPagamento, .Text, .SubItems(11), Bdados.Converte(Trim(.SubItems(10)), TCMonetario), 4, .SubItems(3))
                    Bdados.GravaDados "TAB_PAGAMENTO_EXTRATO", ValoresExt, Campos, "TPE_TGT_COD_PAGAMENTO=" & .Text & " and TPE_COD_PAGAMENTO_EXTRATO=" & CodPagamento
                End With
            End If
        Next
    End If
    If CodPagamento <> "" And CodPagamento <> "0" Then
'        Bdados.AbreTrans
        CodAuto = CodPagamento
        Campos = "TAI_COD_AUTO,"
        Campos = Campos & "TAI_INSCRICAO,"
        Campos = Campos & "TAI_DATA_AUTO,"
        Campos = Campos & "TAI_DATA_VENCIMENTO,"
        Campos = Campos & "TAI_VALOR_AUTO,"
        Campos = Campos & "TAI_VALOR_AGRAVANTE,"
        Campos = Campos & "TAI_VALOR_TOTAL,"
        Campos = Campos & "TAI_OBS,"
        Campos = Campos & "TAI_STATUS,"
        Campos = Campos & "TAI_TIP_COD_IMPOSTO,"
        Campos = Campos & "TAI_FATO_GERADOR,"
        Campos = Campos & "TAI_RELATO_FISCAL,"
        Campos = Campos & "TAI_TIPO_AUTO,"
        Campos = Campos & "TAI_VALOR_OBRIGACAO,"
        Campos = Campos & "TAI_JUROS,"
        Campos = Campos & "TAI_MULTA,"
        Campos = Campos & "TAI_ATUALIZACAO,TAI_FUNDAMENTACAO"
        
'        Call GravarTexto("AUTO INFRACAO", txtTermo)
'        Call GravarTexto("AUTO INFRACAO2", txtTexo2)
        Valore = Bdados.PreparaValor(Bdados.Converte(CodPagamento, tctexto), _
                Bdados.Converte(txtInscricao, tctexto), Bdados.Converte(txtData, TCDataHora), _
                Bdados.Converte(txtDataVencimento, TCDataHora), _
                Bdados.Converte(txtValor, TCMonetario), Bdados.Converte(txtAgravante, TCMonetario), _
                Bdados.Converte(txtValorPago, TCMonetario), Bdados.Converte(txtObservacao, tctexto), _
                StatusAtivo, Bdados.Converte(cboImposto.Coluna(0).Valor, tctexto), _
                Bdados.Converte(txtFatorGerador, tctexto), Bdados.Converte(txtTexto, tctexto), _
                CboTipoAuto.Coluna(1).Valor, _
                Bdados.Converte(CDbl(Nvl(Trim(txtOriginal), 0)) - (CDbl(Nvl(Trim(txtMulta), 0)) + CDbl(Nvl(Trim(txtJuros), 0)) + CDbl(Nvl(Trim(txtAtualizacao), 0)) + CDbl(Nvl(Trim(txtValor), 0))), TCMonetario), _
                Bdados.Converte(txtJuros, TCMonetario), _
                Bdados.Converte(txtMulta, TCMonetario), Bdados.Converte(txtAtualizacao, TCMonetario), _
                Bdados.Converte(txtTermo, tctexto))
        Condicao = "TAI_COD_AUTO = " & Bdados.Converte(CodPagamento, tctexto)
        
        If Bdados.GravaDados("TAB_AUTO_INFRACAO", Valore, Campos, Condicao) Then
            Gravou = True
            'GRAVO OS DETALHES...
            
            If Bdados.DeletaDados("TAB_INFRACAO_AUTO", "TIA_COD_AUTO = " & CodPagamento) Then
                Campos = "TIA_COD_AUTO,TIA_INFRACAO,TAI_AGRAVANTE"
                For i = 1 To grdInfra.ListItems.Count
                    If grdInfra.ListItems(i).Checked Then
                           Valore = Bdados.PreparaValor(CodPagamento, grdInfra.ListItems(i), Valores(i).Percentual)
                           Bdados.GravaDados "TAB_INFRACAO_AUTO", Valore, Campos, "TIA_COD_AUTO = " & CodPagamento & " AND TIA_INFRACAO = " & grdInfra.ListItems(i)
                           Base = Base + 1
                    End If
                Next
                'SE CHEGOU ATÉ AQUI ENTÃO TUDO CORREU CERTO...
                For i = 1 To grdInfra.ListItems.Count
                    If grdInfra.ListItems(i).Checked Then
                        Base2 = Base2 + 1
                    End If
                Next
                If Base = Base2 Then
'                    Bdados.GravaTrans
                    Avisa "Operação Concluída com Sucesso."
                    If Confirma("Deseja imprimir Auto de Infração?", "CIAP") Then
                        Imprimir
                    End If
                    cmdCancela_Click
                Else
                    Avisa "Erro ao gravar infrações."
'                    Bdados.CancelaTrans
                End If
            Else
                Avisa "Erro ao excluir infrações."
'                Bdados.CancelaTrans
            End If
            
        Else
            Avisa "Erro ao gravar Auto de Infração."
'            Bdados.CancelaTrans
        End If
    Else
        Avisa "Erro ao gerar obrigação."
        'Bdados.CancelaTrans
    End If
    
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricao
End Sub

Private Sub cmdSair_Click()
    
    Unload Me
End Sub


Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    PreencherGrid
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    CodImp = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ATRANS))
    DataVenc = Imposto.BuscaDataVencimento(CodImp, Year(Date))
    txtValor.Enabled = False
    Obrig.PreencheComboTributo cboImposto, True
    Configura_Valores
    CboTipoAuto.PreencherGeral Bdados, "TIPO DE AUTO"
    txtTermo = BuscaTexto("AUTO INFRACAO")
    txtTexo2 = BuscaTexto("AUTO INFRACAO2")
    CboTipoAuto.SetarLinha 1, 1
    CboTipoAuto_Click
End Sub

Private Sub txtCotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub Configura_Valores()
    txtData = Date
    txtDataVencimento = DateAdd("d", Temp.PegaParametro(Bdados, "VENCIMENTO_AUTO"), txtData)
End Sub

Private Sub PreencherGrid()
      Dim Sql As String
    
    Sql = " SELECT TIN_COD_INFRACAO AS Infração,"
    Sql = Sql & " tin_descricao_infracao as Descrição,"
    Sql = Sql & " tin_valor_ufm As VALOR, tin_artigo As Artigo,"
    Sql = Sql & " tin_Agravante_UFM as Agravante"
    Sql = Sql & " From TAB_INFRACAO"
    
    grdInfra.Preencher Bdados, Sql
End Sub


Private Sub Text1_Change()

End Sub

Private Sub TiraLixoAgravante()
    Dim Contador As Integer
    
    For Contador = 1 To grdInfra.ListItems.Count
        If Not grdInfra.ListItems(Contador).Checked Then
            Valores(grdInfra.ListItems(Contador).Index).CodigoInfracao = ""
            Valores(grdInfra.ListItems(Contador).Index).Percentual = ""
        End If
    Next
End Sub
Private Sub AtualizaAgravante()
    On Error Resume Next
    Dim Contador As Integer
     
    'TIRO O LIXO
    TiraLixoAgravante
    For Contador = 1 To grdInfra.ListItems.Count
        If grdInfra.ListItems(Contador).Selected Then
            If Valores(grdInfra.ListItems(Contador).Index).CodigoInfracao = grdInfra.ListItems(Contador) Then
                txtValorAgravante = Valores(grdInfra.ListItems(Contador).Index).Percentual
            End If
        End If
    Next
End Sub





Private Sub grdInfra_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    txtValorAgravante = "0,00"
    If Item.Checked = False Then
        txtValorAgravante = "0,00"
    End If
    If txtValorAgravante = "0,00" Then
        txtValorAgravante = Format(Nvl(Item.SubItems(4), 0), Const_Monetario)
        cmdConfirmar_Click
    End If
    Item.Selected = True
    Pega_Valores
    AtualizaAgravante

End Sub

Private Sub grdInfra_ItemClick(ByVal Item As MSComctlLib.IListItem)
    txtValorAgravante = "0,00"
    AtualizaAgravante
    Pega_Agravante
End Sub

Private Function Pega_Agravante() As Double
    On Error Resume Next
    Dim Contador As Integer
    Dim Agravante As Double
    
    
    For Contador = 1 To grdInfra.ListItems.Count
        If grdInfra.ListItems(Contador) = Valores(Contador).CodigoInfracao Then
            Agravante = Agravante + Valores(Contador).Percentual
        End If
    Next
    txtValorPago = Nvl(txtValor, 0) + (Agravante * Nvl(txtValor, 0) / 100)
    txtValorPago = Format(CCur(Nvl(txtValorPago, 0)) + CCur(txtValorTotalDebito), Const_Monetario)
    txtAgravante = Agravante
    
End Function

Private Sub lstParcelas_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    Item.Selected = True
    AtualizaValores
    Pega_Valores
    Pega_Agravante
    
End Sub

Private Sub txtInscricao_LostFocus()
    If txtInscricao = "" Then Exit Sub
    txtInscricao = BuscaContribuinte(txtInscricao, txtRazao, txtEndereco)
End Sub
Public Function GravarTexto(NomeTexto As String, Texto As String) As Boolean
    Dim Valores As String
    Dim Campos As String
    Valores = Bdados.PreparaValor(NomeTexto, Texto)
    Campos = "tpt_parametro,TPT_TEXTO"
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Valores, Campos, "TPT_PARAMETRO = '" & NomeTexto & "'"
End Function

Private Sub txtTotalParc_Change()
'    txtValorPago = CDbl(Nvl(Trim(txtTotalParc), 0)) + CDbl(Nvl(Trim(txtdebitoRestante), 0))
End Sub
Private Sub AtualizaValores()
    Dim i As Integer
    Valor = 0
    Juros = 0
    Multa = 0
    Saldo = 0
    Atualizacao = 0
    Desconto = 0
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked = True Then
            Valor = Valor + lstParcelas.ListItems(i).SubItems(5)
            Atualizacao = Atualizacao + lstParcelas.ListItems(i).SubItems(6)
            Juros = Juros + lstParcelas.ListItems(i).SubItems(7)
            Multa = Multa + lstParcelas.ListItems(i).SubItems(8)
            Desconto = Desconto + lstParcelas.ListItems(i).SubItems(9)
            Saldo = Saldo + lstParcelas.ListItems(i).SubItems(10)
        End If
    Next
    txtAtualizacao = Format(Atualizacao, Const_Monetario)
    txtJuros = Format(Juros, Const_Monetario)
    txtMulta = Format(Multa, Const_Monetario)
    txtOriginal = Format(Saldo, Const_Monetario)
    txtValorTotalDebito = txtOriginal
End Sub

Private Sub txtValorAgravante_LostFocus()
    If txtValorAgravante <> "" Then
        cmdConfirmar_Click
    End If
End Sub

VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREL401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Dinâmico"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleMode       =   0  'User
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fraDados 
      Height          =   870
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Top             =   4800
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1535
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      Begin Threed.SSCommand cmdAtualizarInformacoes 
         Height          =   255
         Left            =   3795
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Atualizar lista de informações"
         Top             =   3270
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         _Version        =   196610
         Font3D          =   3
         MousePointer    =   16
         ForeColor       =   128
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "!"
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   3570
      Index           =   1
      Left            =   4320
      TabIndex        =   20
      Top             =   660
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   6297
      _Version        =   196610
      ForeColor       =   128
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " Ordem"
      Alignment       =   2
      Begin VB.PictureBox grdOrdem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   120
         ScaleHeight     =   2970
         ScaleWidth      =   3195
         TabIndex        =   1
         Top             =   240
         Width           =   3225
      End
      Begin VTOcx.cmdVISUAL cmdSobeOrdem 
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   3270
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         Caption         =   "+"
         CorBorda        =   8421504
         CorFrente       =   128
      End
      Begin VTOcx.cmdVISUAL cmdRetirarOrdem 
         Height          =   255
         Left            =   3090
         TabIndex        =   41
         Top             =   3270
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         Caption         =   "X"
         CorBorda        =   8421504
         CorFrente       =   128
      End
      Begin VTOcx.cmdVISUAL cmdDesceOrdem 
         Height          =   255
         Left            =   2430
         TabIndex        =   42
         Top             =   3270
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         Caption         =   "-"
         CorBorda        =   8421504
         CorFrente       =   128
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   3165
      Index           =   2
      Left            =   7830
      TabIndex        =   21
      Top             =   660
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   5583
      _Version        =   196610
      ForeColor       =   128
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " Filtro"
      Alignment       =   2
      Begin VB.ComboBox cboOperador 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   495
         Width           =   1005
      End
      Begin VB.TextBox txtFiltro 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1170
         TabIndex        =   3
         Top             =   495
         Width           =   2130
      End
      Begin VB.PictureBox grdFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   105
         ScaleHeight     =   1920
         ScaleWidth      =   3165
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   855
         Width           =   3195
      End
      Begin VTOcx.cmdVISUAL cmdIncluirFiltro 
         Height          =   255
         Left            =   2670
         TabIndex        =   4
         Top             =   2850
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         Caption         =   "!"
         CorBorda        =   8421504
         CorFrente       =   128
      End
      Begin VTOcx.cmdVISUAL cmdRetirarFiltro 
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   2850
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         Caption         =   "X"
         CorBorda        =   8421504
         CorFrente       =   128
      End
      Begin VB.Label lblKeyCampo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   38
         Top             =   2850
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblTipoCampo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   2850
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblCampo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(campo)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   210
         Width           =   720
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   930
      Index           =   5
      Left            =   8370
      TabIndex        =   25
      Top             =   8235
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1640
      _Version        =   196610
      ForeColor       =   128
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Impressão"
      Begin VB.TextBox txtPaginaInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   495
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "1"
         Top             =   195
         Width           =   450
      End
      Begin VB.TextBox txtPaginaFinal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1275
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "1"
         Top             =   195
         Width           =   435
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   510
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Início"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   27
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fim"
         Height          =   195
         Index           =   3
         Left            =   1005
         TabIndex        =   26
         Top             =   225
         Width           =   240
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   570
      Index           =   6
      Left            =   45
      TabIndex        =   28
      Top             =   8595
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   1005
      _Version        =   196610
      ForeColor       =   128
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Configuração da Página"
      Begin VB.TextBox txtTotalPaginas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6855
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   210
         Width           =   420
      End
      Begin VB.ComboBox cboTipoPapel 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Tipos de Papel"
         Top             =   210
         Width           =   2925
      End
      Begin VB.TextBox txtLarguraMaxima 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3825
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "0"
         Top             =   210
         Width           =   420
      End
      Begin VB.TextBox txtLarguraAtual 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4845
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   210
         Width           =   420
      End
      Begin VB.TextBox txtTamanhoFonte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7770
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "0"
         Top             =   210
         Width           =   420
      End
      Begin VB.TextBox txtLinhasPorPagina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5790
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "0"
         Top             =   210
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Páginas"
         Height          =   195
         Index           =   1
         Left            =   6270
         TabIndex        =   36
         Top             =   255
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usadas"
         Height          =   195
         Index           =   0
         Left            =   4305
         TabIndex        =   34
         Top             =   255
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caracteres"
         Height          =   195
         Index           =   0
         Left            =   3015
         TabIndex        =   33
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fonte"
         Height          =   195
         Index           =   4
         Left            =   7335
         TabIndex        =   31
         Top             =   255
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Linhas"
         Height          =   195
         Index           =   1
         Left            =   5325
         TabIndex        =   29
         Top             =   255
         Width           =   450
      End
   End
   Begin Threed.SSFrame fraResultado 
      Height          =   2520
      Left            =   45
      TabIndex        =   17
      Top             =   4245
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4445
      _Version        =   196610
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Resultado"
      Begin VB.PictureBox grdResultado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2205
         Left            =   90
         ScaleHeight     =   2175
         ScaleWidth      =   10965
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Width           =   10995
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TREL401.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   10140
      TabIndex        =   6
      Top             =   3870
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      Caption         =   "Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdLimpar 
      Height          =   375
      Left            =   10170
      TabIndex        =   8
      Top             =   8340
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   10170
      TabIndex        =   0
      Top             =   8790
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1380
      Left            =   60
      TabIndex        =   10
      Top             =   6810
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2434
      _Version        =   196610
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Resultado"
      Begin VB.TextBox txtSelect 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   1035
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   11010
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   75
      TabIndex        =   30
      Top             =   8310
      Width           =   390
   End
   Begin VB.Menu mnuGrid 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRetirar 
         Caption         =   "&Retirar campo da lista"
      End
      Begin VB.Menu mnulinha 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTamanho 
         Caption         =   "&Mudar tamanho do campo"
      End
   End
End
Attribute VB_Name = "TREL401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private strTabelasJoin As String
Private strFiltrosSemAlias As String
Private Const cteEspacamentoColunas As Integer = 2
Private Const cteLinhasCabecalho As Integer = 9
Private Const cteLinhasRodape As Integer = 4
Private Const MaxTamCont As Integer = 3
Private intTamanhoNumeroSequencial As Integer

Private Sub cboOperador_Click()
    txtFiltro.SetFocus
End Sub

Private Sub cboTipoPapel_Click()
    Dim x As Integer
    If cboTipoPapel <> "" Then
        'EXEMPLO: cboTipoPapel.AddItem "A4-Vertical (40 Colunas, 42 Linhas)", 0
        x = PosPic(cboTipoPapel, "(") + 1
        txtLarguraMaxima = Mid(cboTipoPapel, x, InStr(x, cboTipoPapel, " ") - x)
        txtTituloRelatorio.MaxLength = txtLarguraMaxima
        x = PosPic(cboTipoPapel, ",") + 2
        txtLinhasPorPagina = Mid(cboTipoPapel, x, InStr(x, cboTipoPapel, " ") - x)
    End If
End Sub

Private Sub cmdAtualizarInformacoes_Click()
    MontarArvore
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo trata
'********************************
'           FUNCAO
'      cmdBuscar_Click
'********************************
'         PROPOSITO
'   Ler as opcoes definidas pelo
'usuario e gerar a sql correspondente
'********************************

    Dim strCampos As String, strCamposSemAlias As String
    Dim strAliasUtilizados As String
    Dim strTabelas As String, strTabelasSelecionadas() As String
    Dim strJoin As String, strTabelasAdicionais As String
    Dim strCondicao As String, strCamposFiltro As String
    Dim strSql As String
    
    fraResultado.Caption = "Resultado :  Processando..."
    Screen.MousePointer = vbHourglass
    '1 - Campos selecionados para o resultado
    strCampos = BuscaCampos(grdOrdem, strCamposSemAlias)
    If strCampos = "" Then
        '2 - Limpa resultado se nao achar campos para exibir
        strSql = ""
    Else
        '3 - Criterio de filtro definido pelo usuario
        strCondicao = BuscaCondicao(grdFiltro, strCamposFiltro)
        '4 - Tabelas que participarao da query
        strAliasUtilizados = BuscaAliasUtilizados(grdOrdem, grdFiltro)
        strTabelas = BuscaTabelas(strCamposSemAlias, strCamposFiltro, strAliasUtilizados, strTabelasSelecionadas)
        '5 - Relacionamento entre as tabelas participantes
        strJoin = BuscaJoins(strTabelasSelecionadas, strTabelasAdicionais)
        '6 - Monta query
        'strSql = "SELECT " & strCampos & " FROM " & strTabelas & IIf(strJoin = "", "", " WHERE ") & strJoin & IIf(strCondicao = "", "", IIf(strJoin = "", " WHERE ", " AND ") & strCondicao)
        strSql = "SELECT " & strCampos & " FROM " & strTabelas & IIf(Trim(strJoin) = "", "", " WHERE " & strJoin) & IIf(Trim(strCondicao) = "", "", IIf(Trim(strJoin) = "", " WHERE ", " AND ") & strCondicao)
    End If
    '7 - Exibe o resultado
    txtSelect = strSql
    If Not MontaGrid(Bdados, grdResultado, strSql) Then
        Avisa "Não foi possível interpretar as condições informados."
    End If
    fraResultado.Caption = "Resultado : " & grdResultado.ListItems.Count
    
    '8 - Total de paginas que o relatorio vai usar
    DoEvents
    Call txtLinhasPorPagina_LostFocus
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
trata:
    Util.Erro Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
    Resume
End Sub

Private Sub cmdDesceOrdem_Click()
    Dim strChavePosterior As String
    Dim strCampoPosterior As String
    Dim strTamanhoPosterior As String
    Dim strTipoPosterior As String
    Dim strChaveAtual As String
    
    If Not grdOrdem.SelectedItem Is Nothing Then
        If grdOrdem.SelectedItem.Index < grdOrdem.ListItems.Count Then
            '1 - Guarda as chaves dos itens envolvidos
            strChavePosterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).Key
            strChaveAtual = grdOrdem.SelectedItem.Key
            '2 - Apaga as chaves que serao trocadas
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).Key = ""
            grdOrdem.SelectedItem.Key = ""
            '3 - Guarda o conteudo do item posterior
            strCampoPosterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).Text
            strTamanhoPosterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).SubItems(1)
            strTipoPosterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).SubItems(2)
            '4 - Troca as informacoes do item posterior
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).Key = strChaveAtual
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).Text = grdOrdem.SelectedItem.Text
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).SubItems(1) = grdOrdem.SelectedItem.SubItems(1)
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).SubItems(2) = grdOrdem.SelectedItem.SubItems(2)
            '5 - Atualiza o item selecionado
            grdOrdem.SelectedItem.Key = strChavePosterior
            grdOrdem.SelectedItem.Text = strCampoPosterior
            grdOrdem.SelectedItem.SubItems(1) = strTamanhoPosterior
            grdOrdem.SelectedItem.SubItems(2) = strTipoPosterior
            '6 - Seleciona a nova posicao do item
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index + 1).Selected = True
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo trata
    Dim intPagina As Integer
    Dim intRegistroInicial As Integer, intRegistroFinal As Integer
    Dim intRegistrosPorPagina As Integer
    Dim i As Integer
    Dim LarguraTotal As Integer
    Dim PrimeiroReg As Integer
    
    If grdResultado.ListItems.Count = 0 Then
        Informa "Não exite informação para ser impressa."
    Else
        Screen.MousePointer = 11
        intRegistroInicial = 0
        intRegistrosPorPagina = Val(txtLinhasPorPagina)
        
        LarguraTotal = CalcularLarguraNecessaria(grdResultado.ListItems.Count, grdResultado.ColumnHeaders.Count, txtLarguraMaxima)
        PrimeiroReg = 0
        
        For intPagina = txtPaginaInicial To txtPaginaFinal
            If cboTipoPapel.ListIndex = 1 Then
                Printer.Orientation = 2 'horizontal
            Else
                Printer.Orientation = 1 ' vertical
            End If
            intRegistroInicial = (intRegistrosPorPagina * (intPagina - 1)) + 1
            If PrimeiroReg = 0 Then PrimeiroReg = intRegistroInicial
            intRegistroFinal = intRegistroInicial + intRegistrosPorPagina - 1
            ImprimirCabecalho LarguraTotal
            ImprimirCorpo intRegistroInicial, intRegistroFinal, LarguraTotal
            i = 0
            For i = intRegistroFinal + 1 To intRegistrosPorPagina * intPagina
                Printer.Print ""
            Next i
            ImprimirRodape intPagina, txtTotalPaginas, LarguraTotal, IIf(intPagina = txtPaginaFinal, intRegistroFinal - (PrimeiroReg - 1), 0)
            Printer.NewPage
        Next intPagina
        Printer.EndDoc
        Screen.MousePointer = 0
        Informa "Fim de Impressão."
    End If
    Exit Sub
    
trata:
    Screen.MousePointer = 0
    Util.Erro Err.Description
    Printer.KillDoc
    Exit Sub
    Resume
End Sub

Private Sub cmdIncluirFiltro_Click()
'********************************
'           FUNCAO
'      cmdIncluirFiltro_Click
'********************************
'         PROPOSITO
'   Montar a lista de criterios de
'filtro do usuario
'********************************
'AUTOR : Sergio Queiroz
'DATA : 15.03.2002
'********************************
'         REVISOES
'AUTOR :
'DATA :
'ACAO :
'********************************
    Dim itmFiltro As ListItem
    Dim intInicio As Integer, intTamanho As Integer
    Dim strFiltro As String, strCampo As String, intTipoCampo As Integer
    
    '1 - Verifica se as informacoes foram passadas pelo usuario
    If Not grdOrdem.SelectedItem Is Nothing Then
        If Trim$(txtFiltro) <> "" Then
            '2 - Busca o nome do campo
            'intInicio = InStr(1, grdOrdem.SelectedItem.Key, ".")
            If lblTipoCampo = "" Then
                strCampo = grdOrdem.SelectedItem.Key
            Else
                strCampo = lblKeyCampo
            End If
            '3 - Busca o tipo do campo
            If lblTipoCampo = "" Then
                intTipoCampo = Nvl(grdOrdem.SelectedItem.SubItems(2), 0)
            Else
                intTipoCampo = lblTipoCampo
            End If
            '4 - Monta a expressao do filtro
            Select Case intTipoCampo
                Case tipTexto
                    strFiltro = strCampo & " " & cboOperador & " '" & txtFiltro & "'"
                Case tipData
                    strFiltro = strCampo & " " & cboOperador & " " & Bdados.Converte(txtFiltro, TCDataHora)
                Case tipInteiro, tipMoeda
                    strFiltro = strCampo & " " & cboOperador & " " & txtFiltro
                Case Else
                    strFiltro = strCampo & " " & cboOperador & " '" & txtFiltro & "'"
            End Select
            '5 - Inclui o filtro
            If lblTipoCampo = "" Then
                Set itmFiltro = grdFiltro.ListItems.Add(, strFiltro, grdOrdem.SelectedItem.Text)
            Else
                Set itmFiltro = grdFiltro.ListItems.Add(, strFiltro, lblCampo)
            End If
            itmFiltro.SubItems(1) = cboOperador
            itmFiltro.SubItems(2) = txtFiltro
            itmFiltro.SubItems(3) = intTipoCampo
            itmFiltro.SubItems(4) = strCampo
            lblTipoCampo = ""
            lblKeyCampo = ""
        End If
    End If
    '4 - Prepara a tela para novo filtro
    cboOperador.ListIndex = -1: txtFiltro = ""
    cboOperador.SetFocus
End Sub

Private Sub cmdLimpar_Click()
    Dim i As Integer, j As Integer
    
    '1 - Limpa a arvore
    j = treOpcao.NodesCollection.Count
    For i = 1 To j
        If treOpcao.Nodes(i).Image Like "*CHECK*" Then
            treOpcao.Value(i) = OptionTreeCheckNone
        End If
    Next i
    '2 - Limpa as grades
    grdFiltro.ListItems.Clear
    grdOrdem.ListItems.Clear
    Util.MontaGrid Bdados, grdResultado, ""
    txtTituloRelatorio = ""
    fraResultado.Caption = "Resultado"
    txtLarguraAtual = 0
    txtTotalPaginas = 0
    lblCampo = "(campo)"
    cboOperador.ListIndex = -1
    txtFiltro = ""
End Sub

Private Sub cmdRetirarFiltro_Click()
    If Not grdFiltro.SelectedItem Is Nothing Then
        grdFiltro.ListItems.Remove grdFiltro.SelectedItem.Index
        lblTipoCampo = ""
        lblKeyCampo = ""
    End If
End Sub

Private Sub cmdRetirarOrdem_Click()
    If Not grdOrdem.SelectedItem Is Nothing Then
        If Util.Confirma("Retirar " & grdOrdem.SelectedItem & " ?") Then
            txtLarguraAtual = CDbl(txtLarguraAtual) - grdOrdem.SelectedItem.SubItems(1) - cteEspacamentoColunas
            grdOrdem.ListItems.Remove grdOrdem.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    'If confirma("Deseja mesmo sair e cancelar o relatório atual?") Then
        Unload Me
    'End If
End Sub

Private Sub cmdSobeOrdem_Click()
'********************************
'           FUNCAO
'      cmdSobeOrdem_Click
'********************************
'         PROPOSITO
'   Diminuir o numero de ordem do
'campo selecionado na lista de
'campos escolhidos
'********************************
'AUTOR : Sergio Queiroz
'DATA : 15.03.2002
'********************************
'         REVISOES
'AUTOR :
'DATA :
'ACAO :
'********************************
    Dim strChaveAnterior As String
    Dim strCampoAnterior As String
    Dim strTamanhoAnterior  As String
    Dim strTipoAnterior  As String
    
    Dim strChaveAtual As String
    
    If Not grdOrdem.SelectedItem Is Nothing Then
        If grdOrdem.SelectedItem.Index > 1 Then
            '1 - Guarda as chaves dos itens envolvidos
            strChaveAnterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).Key
            strChaveAtual = grdOrdem.SelectedItem.Key
            '2 - Apaga as chaves que serao trocadas
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).Key = ""
            grdOrdem.SelectedItem.Key = ""
            '3 - Guarda o conteudo do item Anterior
            strCampoAnterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).Text
            strTamanhoAnterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).SubItems(1)
            strTipoAnterior = grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).SubItems(2)
            '4 - Troca as informacoes do item Anterior
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).Key = strChaveAtual
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).Text = grdOrdem.SelectedItem.Text
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).SubItems(1) = grdOrdem.SelectedItem.SubItems(1)
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).SubItems(2) = grdOrdem.SelectedItem.SubItems(2)
            '5 - Atualiza o item selecionado
            grdOrdem.SelectedItem.Key = strChaveAnterior
            grdOrdem.SelectedItem.Text = strCampoAnterior
            grdOrdem.SelectedItem.SubItems(1) = strTamanhoAnterior
            grdOrdem.SelectedItem.SubItems(2) = strTipoAnterior
            '6 - Seleciona a nova posicao do item
            grdOrdem.ListItems(grdOrdem.SelectedItem.Index - 1).Selected = True
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    MontarArvore
    txtTamanhoFonte = 10
    
    
    'Tipos pre-definidos de papel
    cboTipoPapel.Clear
    cboTipoPapel.AddItem "A4-Ver. (84 Caracteres, 60 Linhas)", 0
    cboTipoPapel.AddItem "A4-Hor. (129 Caracteres, 37 Linhas)", 1
    cboTipoPapel.AddItem "Matricial (158 Caracteres, 46 Linhas)", 2
    cboTipoPapel.ListIndex = 0
    Call cboTipoPapel_Click
    
End Sub

Private Sub MontarArvore()
'********************************
'           FUNCAO
'        MontarArvore
'********************************
'         PROPOSITO
'   Prepara as opcoes de campos
'********************************
'AUTOR : Sergio Queiroz
'DATA : 15.03.2002
'********************************
'         REVISOES
'AUTOR : Andre Fabiano
'DATA :24.04.2002
'ACAO :Criar sub pastas para diminuir tamanho da arvore de tabelas
'********************************
    'treOpcao.Clear
    
    Dim rsTabelas As VSRecordset, rsCampos As VSRecordset, RSVinculados As VSRecordset
    If Bdados.AbreTabela("SELECT * FROM TAB_TABELAS where TTA_TTA_COD_TABELA_ORIGEM = 0 ORDER BY TTA_ALIAS_USUARIO ", rsTabelas) Then
        rsTabelas.MoveFirst
        Do While Not rsTabelas.EOF
            treOpcao.AddFolder "TTA_NOME=" & rsTabelas!TTA_NOME & " as [" & rsTabelas!TTA_ALIAS_USUARIO & "]", , rsTabelas!TTA_ALIAS_USUARIO
            'ADICIONA CAMPOS
            If Bdados.AbreTabela("SELECT * FROM TAB_CAMPOS_TABELAS WHERE TCP_TTA_COD_TABELA=" & rsTabelas!TTA_COD_TABELA & " ORDER BY TCP_ALIAS_USUARIO", rsCampos) Then
                Do While Not rsCampos.EOF
                    treOpcao.AddCheck "TCP_NOME=" & rsCampos!TCP_NOME & ":TCP_TIPO=" & rsCampos!TCP_TIPO & ":TCP_TAMANHO=" & rsCampos!TCP_TAMANHO & ":TCP_COD_CAMPO=" & rsCampos!TCP_COD_CAMPO, treOpcao.NodesCollection("TTA_NOME=" & rsTabelas!TTA_NOME & " as [" & rsTabelas!TTA_ALIAS_USUARIO & "]"), rsCampos!TCP_ALIAS_USUARIO, OptionTreeCheckNone
                    rsCampos.MoveNext
                Loop
            End If
            Bdados.FechaTabela rsCampos
            'BUSCA SUB - TABELAS
            If Bdados.AbreTabela("SELECT * FROM TAB_TABELAS WHERE TTA_TTA_COD_TABELA_ORIGEM =" & rsTabelas!TTA_COD_TABELA & " ORDER BY TTA_ALIAS_USUARIO ", RSVinculados) Then
                RSVinculados.MoveFirst
                Do While Not RSVinculados.EOF
                    treOpcao.AddFolder "TTA_NOME=" & RSVinculados!TTA_NOME & " as [" & RSVinculados!TTA_ALIAS_USUARIO & "]", treOpcao.NodesCollection("TTA_NOME=" & rsTabelas!TTA_NOME & " as [" & rsTabelas!TTA_ALIAS_USUARIO & "]"), RSVinculados!TTA_ALIAS_USUARIO
                    If Bdados.AbreTabela("SELECT * FROM TAB_CAMPOS_TABELAS WHERE TCP_TTA_COD_TABELA=" & RSVinculados!TTA_COD_TABELA & " ORDER BY TCP_ALIAS_USUARIO", rsCampos) Then
                        Do While Not rsCampos.EOF
                            treOpcao.AddCheck "TCP_NOME=" & rsCampos!TCP_NOME & ":TCP_TIPO=" & rsCampos!TCP_TIPO & ":TCP_TAMANHO=" & rsCampos!TCP_TAMANHO & ":TCP_COD_CAMPO=" & rsCampos!TCP_COD_CAMPO, treOpcao.NodesCollection("TTA_NOME=" & RSVinculados!TTA_NOME & " as [" & RSVinculados!TTA_ALIAS_USUARIO & "]"), rsCampos!TCP_ALIAS_USUARIO, OptionTreeCheckNone
                            rsCampos.MoveNext
                        Loop
                    End If
                    Bdados.FechaTabela rsCampos
                    RSVinculados.MoveNext
                Loop
            End If
            rsTabelas.MoveNext
        Loop
    End If
    Bdados.FechaTabela rsTabelas
End Sub

Private Function BuscaTabelas(strCamposSemAlias As String, strCamposFiltro As String, strAliasUtilizados As String, ByRef retTabelasSelecionadas() As String) As String
    Dim strSql As String
    Dim rstTabela As VSRecordset
    Dim i As Integer, j As Integer, intQtdTabelas As Integer
    
    Erase retTabelasSelecionadas()
    BuscaTabelas = ""
    intQtdTabelas = 0
    
    'strSql = "SELECT DISTINCT TAB_TABELAS.* FROM TAB_TABELAS,TAB_CAMPOS_TABELAS " & _
                " WHERE TTA_COD_TABELA=TCP_TTA_COD_TABELA AND " & _
                        " TCP_NOME IN (" & strCamposSemAlias & IIf(Len(strCamposFiltro) > 0, ",", "") & strCamposFiltro & ") AND" & _
                        " TCP_ALIAS_USUARIO IN (" & strAliasUtilizados & ")"
                        
                        
                        'IIf(Len(strCamposFiltro) > 0, ",", "") & strCamposFiltro & ")
    strSql = "SELECT DISTINCT TAB_TABELAS.* FROM TAB_TABELAS,TAB_CAMPOS_TABELAS " & _
                " WHERE TTA_COD_TABELA=TCP_TTA_COD_TABELA AND (" & _
                        " TCP_NOME IN (" & strCamposSemAlias & IIf(Trim(strFiltrosSemAlias) <> "", ", ", "") & strFiltrosSemAlias & ") AND" & _
                        " TCP_ALIAS_USUARIO IN (" & strAliasUtilizados & "))"
    If Bdados.AbreTabela(strSql, rstTabela) Then
        rstTabela.MoveFirst
        Do While Not rstTabela.EOF
            BuscaTabelas = BuscaTabelas & IIf(Len(BuscaTabelas) > 0, ",", "")
            If ("" & rstTabela!TTA_WHERE) = "" Then
                BuscaTabelas = BuscaTabelas & rstTabela!TTA_NOME & " " '& " as [" & rstTabela!TTA_ALIAS_USUARIO & "]"
            Else
                BuscaTabelas = BuscaTabelas & "(SELECT * FROM " & rstTabela!TTA_NOME & " WHERE " & rstTabela!TTA_WHERE & ") as [" & rstTabela!TTA_ALIAS_USUARIO & "]"
            End If
            ReDim Preserve retTabelasSelecionadas(0 To 2, 0 To intQtdTabelas)
            retTabelasSelecionadas(0, intQtdTabelas) = rstTabela!TTA_COD_TABELA
            retTabelasSelecionadas(1, intQtdTabelas) = "[" & rstTabela!TTA_ALIAS_USUARIO & "]."
            retTabelasSelecionadas(2, intQtdTabelas) = rstTabela!TTA_NOME
            intQtdTabelas = intQtdTabelas + 1
            rstTabela.MoveNext
        Loop
    End If
    Bdados.FechaTabela rstTabela
End Function

Private Sub grdFiltro_DblClick()
    If Not grdFiltro.SelectedItem Is Nothing Then
        lblCampo = grdFiltro.SelectedItem.Text
        cboOperador.ListIndex = ListIndexDe(cboOperador, grdFiltro.SelectedItem.SubItems(1))
        txtFiltro = grdFiltro.SelectedItem.SubItems(2)
        lblTipoCampo = grdFiltro.SelectedItem.SubItems(3)
        lblKeyCampo = grdFiltro.SelectedItem.SubItems(4)
        grdFiltro.ListItems.Remove grdFiltro.SelectedItem.Index
    End If
End Sub

Private Sub grdOrdem_Click()
    If Not grdOrdem.SelectedItem Is Nothing Then
        lblCampo = grdOrdem.SelectedItem
    End If
End Sub

Private Sub grdOrdem_DblClick()
    On Error Resume Next
    cboOperador = "="
End Sub

Private Sub grdOrdem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        mnuTamanho.Caption = "&Mudar tamanho de " & grdOrdem.SelectedItem
        mnuRetirar.Caption = "&Retirar " & grdOrdem.SelectedItem
        Me.PopupMenu mnuGrid
    End If
End Sub


Private Sub mnuRetirar_Click()
    cmdRetirarOrdem_Click
End Sub

Private Sub mnuTamanho_Click()
    Dim x As String
    If Not grdOrdem.SelectedItem Is Nothing Then
        lblCampo = grdOrdem.SelectedItem
        x = InputBox("Digite o novo tamanho do campo '" & grdOrdem.SelectedItem.Text & "':", "Novo Tamanho", grdOrdem.SelectedItem.SubItems(1))
        If x <> "" Then
            Do Until IsNumeric(x)
                x = InputBox("Digite um novo VALOR NUMÉRICO para o tamanho do campo '" & grdOrdem.SelectedItem.Text & "':", "Novo Tamanho de " & grdOrdem.SelectedItem.Text, grdOrdem.SelectedItem.SubItems(1))
                If x = "" Then Exit Do
            Loop
        End If
        If x <> "" Then
            txtLarguraAtual = CDbl(txtLarguraAtual) - grdOrdem.SelectedItem.SubItems(1) - cteEspacamentoColunas
            grdOrdem.SelectedItem.SubItems(1) = x
            txtLarguraAtual = Val(txtLarguraAtual) + x + cteEspacamentoColunas
        End If
    End If
End Sub

Private Sub treOpcao_CheckClick(ItemNode As ComctlLib.INode, Value As cTreeOpt.OptionTreeCheckTypes)
'********************************
'           FUNCAO
'     treOpcao_CheckClick
'********************************
'         PROPOSITO
'   Preencher a lista de campos e
'tabelas que sera usada na construcao
'da sql
'********************************
    On Error Resume Next
    
    Dim i As Integer, j As Integer, intInicio As Integer, intTamanho As Integer
    Dim itmOrdem As ListItem, itmTabela As ListItem
    Dim strCampo As String, strTabela As String, strAliasTabela As String
    Dim intTipoCampo As Integer
    Dim Index As Integer
    '1 - Busca o nome do campo
    intInicio = Len("TCP_NOME=")
    intTamanho = InStr(intInicio, ItemNode.Key, ":")
    strCampo = Mid(ItemNode.Key, intInicio + 1, intTamanho - intInicio - 1)
    
    '2 - Decifra as informacoes da tabela
    If Not ItemNode.Parent Is Nothing Then
        intInicio = Len("TTA_NOME=")
        strTabela = Mid(ItemNode.Parent.Key, intInicio + 1)
        intInicio = PosPic(strTabela, " as [")
        strAliasTabela = Trim(Mid(strTabela, 1, intInicio))
    End If
    
    If Value = OptionTreeCheckNone Then
        intInicio = InStr(1, ItemNode.Key, "TCP_TAMANHO")
        intTamanho = Len("TCP_TAMANHO=")
        intTamanho = Mid(ItemNode.Key, intInicio + intTamanho, InStr(intInicio + intTamanho, ItemNode.Key, ":") - (intInicio + intTamanho))
        txtLarguraAtual = CDbl(txtLarguraAtual) - intTamanho - cteEspacamentoColunas
        '3 - Remove o campo
        grdOrdem.ListItems.Remove IIf(strAliasTabela = "", "", strAliasTabela & ".") & strCampo
    Else
        '4 - Insere o campo*
        Index = grdOrdem.ListItems.Count + 1
        grdOrdem.ListItems.Add Index, IIf(strAliasTabela = "", "", strAliasTabela & ".") & strCampo, ItemNode.Text
        
        '5 - Tamanho do campo
        intInicio = InStr(1, ItemNode.Key, "TCP_TAMANHO")
        intTamanho = Len("TCP_TAMANHO=")
        intTamanho = Mid(ItemNode.Key, intInicio + intTamanho, InStr(intInicio + intTamanho, ItemNode.Key, ":") - (intInicio + intTamanho))
        grdOrdem.ListItems.Item(Index).SubItems(1) = intTamanho
        txtLarguraAtual = Val(txtLarguraAtual) + intTamanho + cteEspacamentoColunas
        '6 - Tipo do campo
        intInicio = InStr(1, ItemNode.Key, "TCP_TIPO")
        intTamanho = Len(":TCP_TIPO=")
        intTipoCampo = Mid(ItemNode.Key, intInicio + intTamanho - 1, InStr(intInicio + intTamanho - 1, ItemNode.Key, ":") - (intInicio + intTamanho - 1))
        itmOrdem.SubItems(2) = intTipoCampo
    End If
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

Private Function BuscaCampos(grdCampos As Object, ByRef retCamposSemAlias As String) As String
'********************************
'           FUNCAO
'         BuscaCampos
'********************************
'         PROPOSITO
'   Montar a parte Select da sql
'********************************
'AUTOR : Sergio Queiroz
'DATA : 15.03.2002
'********************************
'         REVISOES
'AUTOR :
'DATA :
'ACAO :
'********************************
   Dim i As Integer, j As Integer, k As Integer
    
    BuscaCampos = ""
    retCamposSemAlias = ""
    j = grdCampos.ListItems.Count
    For i = 1 To j
        BuscaCampos = BuscaCampos & IIf(Len(BuscaCampos) > 0, ",", "") & grdCampos.ListItems(i).Key & " as [" & grdCampos.ListItems(i).Text & "]"
        k = PosPic(grdCampos.ListItems(i).Key, ".")
        retCamposSemAlias = retCamposSemAlias & IIf(Len(retCamposSemAlias) > 0, ",", "") & "'" & Mid(grdCampos.ListItems(i).Key, k + 1) & "'"
    Next
End Function

Private Function BuscaCondicao(grdFiltros As Object, ByRef retCamposFiltro As String) As String
'********************************
'           FUNCAO
'         BuscaCondicao
'********************************
'         PROPOSITO
'   Montar a parte da clausula Where
'responsavel pelo filtro da sql
'********************************
'AUTOR : Sergio Queiroz
'DATA : 15.03.2002
'********************************
'         REVISOES
'AUTOR :
'DATA :
'ACAO :
'********************************
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim m As Integer
    
    BuscaCondicao = ""
    retCamposFiltro = ""
    strFiltrosSemAlias = ""
    j = grdFiltros.ListItems.Count
    For i = 1 To j
        BuscaCondicao = BuscaCondicao & IIf(Len(BuscaCondicao) > 0, " AND ", "") & grdFiltros.ListItems(i).Key
        k = PosPic(grdFiltros.ListItems(i).Key, "=")
        
        If k > 0 Then
            l = PosPic(grdFiltros.ListItems(i).Key, "]")
            m = PosPic(grdFiltros.ListItems(i).Key, "=")
            'retCamposFiltro = retCamposFiltro & IIf(Len(retCamposFiltro) > 0, ",", "") & "'" & Mid(grdFiltros.ListItems(i).Key, l + 2, Len(grdFiltros.ListItems(i).Key) - k) & "'"
            'éderson
            retCamposFiltro = retCamposFiltro & IIf(Len(retCamposFiltro) > 0, ",", "") & Mid(grdFiltros.ListItems(i).Key, l + 2, Len(grdFiltros.ListItems(i).Key))
            strFiltrosSemAlias = strFiltrosSemAlias & IIf(Len(strFiltrosSemAlias) > 0, ",", "") & "'" & Mid(grdFiltros.ListItems(i).Key, l + 2, m - l - 3) & "'"
        End If
    Next
End Function


Private Function BuscaJoins(strTabelasSelecionadas() As String, ByRef retTabelasAdicionais As String) As String
'********************************
'           FUNCAO
'         BuscaJoins
'********************************
'         PROPOSITO
'   Montar a parte da clausula Where
'responsavel pelo joining das tabelas
'envolvidas
'********************************
'AUTOR : Sergio Queiroz
'DATA : 16.03.2002
'********************************
'         REVISOES
'AUTOR :
'DATA :
'ACAO :
'********************************
    Dim strSql As String
    Dim rstJoins As VSRecordset
    Dim intCodJoin As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim strAliasTabelaPai As String, strAliasTabelaFilho As String
    Dim strNomeTabelaPai As String, strNomeTabelaFilho As String
    Dim strAliasTabelaAuxiliar As String
    Dim TipoJoin As String
    Dim MesmaTabelaPai As Boolean
    Dim AuxJoin As String
    Dim InsereParentese As Boolean
    Dim Joins() As String
    
    
    ReDim Joins(1 To 1) As String
    Dim GrupoJoin As Integer
    
    BuscaJoins = ""
    retTabelasAdicionais = ""
    GrupoJoin = 1
    '1 - REALIZA A PERMUTACAO SEQUENCIAL DAS TABELAS ENVOLVIDAS
    'Os lacos aninhados realizam a permutacao sequencial, ou seja, digamos que a query envolvera as tabelas 1, 2, 3 e 4.
    'Os lacos, na forma como estao, permitirao a consulta dos relacionamentos da tabela 1 com as tabelas 2, 3 e 4.
    'Depois, consultara os relacionamentos da tabela 2 com as tabelas 3 e 4. E, por fim, da tabela 3 com a 4.
    'Isto garante que todas as combinacoes possiveis serao lidas uma unica vez.
    j = UBound(strTabelasSelecionadas, 2)
    For i = 0 To j
        For k = 0 To j
            If i <> k Then
                '2 - VERIFICA A NECESSIDADE DE TABELAS ADICIONAIS
                'Sempre que tivermos um relacionamento n para n, a 3ª forma normal nos manda construir uma terceira
                'tabela de tal maneira que possamos quebra-lo em dois relacionamento 1 para n
                strSql = "SELECT TJX_TJO_COD_JOIN, TJX_NOME_TABELA" & _
                            " FROM TAB_TABELAS_AUXILIARES_JOIN,TAB_JOINS_TABELAS" & _
                            " WHERE TJO_COD_JOIN=TJX_TJO_COD_JOIN AND" & _
                                " TJO_TTA_COD_TABELA_PAI=" & strTabelasSelecionadas(0, i) & " AND " & _
                                " TJO_TTA_COD_TABELA_FILHO=" & strTabelasSelecionadas(0, k)
                strAliasTabelaAuxiliar = ""
                If Bdados.AbreTabela(strSql, rstJoins) Then
                    retTabelasAdicionais = retTabelasAdicionais & IIf(Len(retTabelasAdicionais) > 0, ",", "") & rstJoins!TJX_NOME_TABELA & " as [JOIN" & rstJoins!TJX_TJO_COD_JOIN & "]"
                    strAliasTabelaAuxiliar = "[JOIN" & rstJoins!TJX_TJO_COD_JOIN & "]."
                End If
                Bdados.FechaTabela rstJoins
                '3 - CAMPOS QUE PARTICIPARAO DO JOIN
                strSql = "SELECT TCJ_TJO_COD_JOIN, TCJ_NOME_CAMPO_PAI, TCJ_NOME_CAMPO_FILHO, TCJ_SENTIDO_RELACIONAMENTO,TJO_TIPO_JOIN FROM TAB_CAMPOS_JOIN,TAB_JOINS_TABELAS" & _
                            " WHERE TJO_COD_JOIN=TCJ_TJO_COD_JOIN AND" & _
                                " TJO_TTA_COD_TABELA_PAI=" & strTabelasSelecionadas(0, i) & " AND " & _
                                " TJO_TTA_COD_TABELA_FILHO=" & strTabelasSelecionadas(0, k)
'                strSql = strSql & " OR TJO_COD_JOIN_VINCULADO = (SELECT TJO_COD_JOIN FROM TAB_JOINS_TABELAS WHERE TJO_TTA_COD_TABELA_PAI=" & strTabelasSelecionadas(0, i) & " AND TJO_TTA_COD_TABELA_FILHO=" & strTabelasSelecionadas(0, k) & ") ORDER BY TJO_COD_JOIN_VINCULADO  ASC"
                If Bdados.AbreTabela(strSql, rstJoins) Then
                    Do While Not rstJoins.EOF
'                        BuscaJoins = BuscaJoins & IIf(Len(BuscaJoins) > 0, " AND ", "")
                        MesmaTabelaPai = IIf(strAliasTabelaPai = strTabelasSelecionadas(1, i), True, False)
                        strAliasTabelaPai = strTabelasSelecionadas(1, i)
                        strAliasTabelaFilho = strTabelasSelecionadas(1, k)
                        strNomeTabelaPai = strTabelasSelecionadas(2, i)
                        strNomeTabelaFilho = strTabelasSelecionadas(2, k)
                        TipoJoin = IIf(rstJoins!TJO_TIPO_JOIN = 0, " INNER JOIN ", IIf(rstJoins!TJO_TIPO_JOIN = 1, "  LEFT JOIN ", " RIGHT JOIN "))
                        Select Case rstJoins!TCJ_SENTIDO_RELACIONAMENTO
                            Case 1 'TABAuxiliar x Pai
                                strAliasTabelaFilho = strTabelasSelecionadas(1, i)
                            Case 2 'TABAuxiliar x Filho
                                strAliasTabelaFilho = strTabelasSelecionadas(1, k)
                        End Select
                        If strAliasTabelaAuxiliar <> "" Then
                            strAliasTabelaPai = strAliasTabelaAuxiliar
                        End If
                        'BuscaJoins = BuscaJoins & strAliasTabelaPai & rstJoins!TCJ_NOME_CAMPO_PAI & " = " & strAliasTabelaFilho & rstJoins!TCJ_NOME_CAMPO_FILHO
                        If MesmaTabelaPai Then
                            InsereParentese = IIf(Trim(Right(Joins(GrupoJoin), 2)) = ")", False, True)
                            'AuxJoin = IIf(InsereParentese, ")", "") & TipoJoin & strNomeTabelaFilho & " ON " & strNomeTabelaPai & "." & rstJoins!TCJ_NOME_CAMPO_PAI & " = " & strNomeTabelaFilho & "." & rstJoins!TCJ_NOME_CAMPO_FILHO & IIf(InsereParentese = False, ") ", "")
                            AuxJoin = IIf(Len(Trim(AuxJoin)) = 0, "", " AND ") & strNomeTabelaPai & "." & rstJoins!TCJ_NOME_CAMPO_PAI & " = " & strNomeTabelaFilho & "." & rstJoins!TCJ_NOME_CAMPO_FILHO & IIf(InsereParentese = False, ") ", "")
                        Else
                            InsereParentese = IIf(Trim(Joins(GrupoJoin)) = "", False, IIf(Trim(Right(Joins(GrupoJoin), 2)) = ")", False, True))
'                            AuxJoin = IIf(InsereParentese, ") ", "") & IIf(GrupoJoin = 1, strNomeTabelaPai, "") & TipoJoin & strNomeTabelaFilho & " ON " & strNomeTabelaPai & "." & rstJoins!TCJ_NOME_CAMPO_PAI & " = " & strNomeTabelaFilho & "." & rstJoins!TCJ_NOME_CAMPO_FILHO
                            AuxJoin = IIf(Len(Trim(AuxJoin)) = 0, "", " AND ") & strNomeTabelaPai & "." & rstJoins!TCJ_NOME_CAMPO_PAI & " = " & strNomeTabelaFilho & "." & rstJoins!TCJ_NOME_CAMPO_FILHO
                        End If
                        
                        'Joins(GrupoJoin) = IIf(MesmaTabelaPai, IIf(Left(Joins(GrupoJoin), 1) = "(", "(", "(("), "") & Joins(GrupoJoin) & AuxJoin & IIf(InsereParentese, ") ", "")
                        Joins(GrupoJoin) = Joins(GrupoJoin) & AuxJoin
                        GrupoJoin = GrupoJoin + 1
                        ReDim Preserve Joins(1 To GrupoJoin) As String
                        MesmaTabelaPai = False
                        rstJoins.MoveNext
                    Loop
                End If
                Bdados.FechaTabela rstJoins
            End If
        Next k
    Next i
    For i = 1 To UBound(Joins)
        BuscaJoins = BuscaJoins & " " & Joins(i)
    Next
End Function

Public Function PreencherCampo(Campo As String, Tamanho As Integer, Tipo As enuTipoCampo) As String
'********************************
'           FUNCAO
'         PreencherCampo
'********************************
'         PROPOSITO
'Preencher o campo com espacos ou
'zeros de forma que o tamanho que
'fique homogeneo
'********************************
'AUTOR : Sergio Queiroz
'DATA : 16.03.2002
'********************************
'         REVISOES
'AUTOR : Sergio Queiroz
'DATA : 17.03.2002
'ACAO : * Caractere de preenchimento
'restringido a espaco
'       * Formatacao do campo moeda
'********************************

    'ASC(" ")=32
    Dim caractere As Long
    caractere = 32
   
   '1 - Consistencia do parametro Campo
    Dim str As String
    str = Campo
    Dim dec As String
    If Tipo = tipMoeda Then
        str = Format(Campo, "Standard")
    End If
    
    '2 - Trunca o tamanho do Campo
    If Tipo <> tipTexto Then
        str = Right(str, Tamanho)
    Else
        str = Left(str, Tamanho)
    End If
    
   '3 - Quantidade de posicoes que faltam preencher
    Dim i As Integer
    i = Tamanho - Len(str)
    If i < 0 Then i = 0
    
   '4 - Preenchimento das posicoes
    Dim preenche As String
    preenche = String(i, caractere)
    
    If Tipo = tipTexto Then
        '5 - Alinha texto à esquerda
        str = UCase(str) & preenche
    Else
        '6 - Alinha numericos à direita
        str = preenche & str
    End If
    Campo = str
    PreencherCampo = str
End Function

Public Function AlinharCampo(Campo As String, Posicao As enuAlinhamentoCampo, Tamanho As Integer) As String
'********************************
'           FUNCAO
'         AlinharCampo
'********************************
'         PROPOSITO
'Posicionar a informacao horizontalmente
'na pagina
'********************************
'AUTOR : Sergio Queiroz
'DATA : 17.03.2002
'********************************
'         REVISOES
'AUTOR :
'DATA :
'ACAO :
'********************************
    Dim intQuantidadeEspacos
    Select Case Posicao
        Case aliEsquerda
            intQuantidadeEspacos = 0
            
        Case aliCentro
            intQuantidadeEspacos = IIf(Len(Campo) > Tamanho, 0, (Tamanho - Len(Campo)) / 2)
        
        Case aliDireita
            intQuantidadeEspacos = Tamanho - Len(Campo)
    End Select
    
    AlinharCampo = String(intQuantidadeEspacos, " ") & Campo
End Function

Public Sub ImprimirCabecalho(LarguraMaxima As Integer)
    Printer.Font.Name = "Courier New"
    Printer.Font.Size = txtTamanhoFonte
    
    '1 - Estado
    Printer.Print AlinharCampo(UCase(Temp.PegaParametro(Bdados, "ESTADO")), aliCentro, LarguraMaxima)
    
    '2 - Prefeitura
    Printer.Print AlinharCampo(UCase(Temp.PegaParametro(Bdados, "CLIENTE")), aliCentro, LarguraMaxima)
    
    '3 - Secretaria
    Printer.Print AlinharCampo(UCase(Temp.PegaParametro(Bdados, "SEMFAZ")), aliCentro, LarguraMaxima)
    
    '4 - Departamento
    Printer.Print AlinharCampo(UCase(Temp.PegaParametro(Bdados, "SETOR")), aliCentro, LarguraMaxima)

    '5 - Branco
    Printer.Print
    '6 - Titulo
    Printer.Print AlinharCampo(UCase(txtTituloRelatorio), aliCentro, LarguraMaxima)
    '7 - Branco
    Printer.Print
    
    Dim intColuna As Integer, intQuantidadeColunas As Integer
    Dim strColunas As String, strLinha As String
    
    intQuantidadeColunas = grdOrdem.ListItems.Count
    For intColuna = 1 To intQuantidadeColunas
        strColunas = strColunas & IIf(Len(strColunas) = 0, Space(intTamanhoNumeroSequencial), "") & Space(cteEspacamentoColunas) & PreencherCampo(grdOrdem.ListItems(intColuna).Text, Nvl(grdOrdem.ListItems(intColuna).SubItems(1), 1), tipTexto)
        strLinha = strLinha & IIf(Len(strLinha) = 0, Space(intTamanhoNumeroSequencial), "") & Space(cteEspacamentoColunas) & String(Nvl(grdOrdem.ListItems(intColuna).SubItems(1), 1), "-")
    Next intColuna
    '8 - Nome das colunas
    Printer.Print strColunas
    
    '9 - Linhas
    Printer.Print strLinha
End Sub

Public Sub ImprimirCorpo(LinhaInicial As Integer, LinhaFinal As Integer, Largura As Integer)
    On Error GoTo trata
    Dim intLinha As Integer, intColuna As Integer, intQuantidadeColunas As Integer
    Dim strLinha As String
    Dim strValorCampo As String
    
    Printer.Font.Name = "Courier New"
    Printer.Font.Size = txtTamanhoFonte
    
    intQuantidadeColunas = grdOrdem.ListItems.Count
    If LinhaFinal > grdResultado.ListItems.Count Then LinhaFinal = grdResultado.ListItems.Count
    For intLinha = LinhaInicial To LinhaFinal
        '1 - Numero de ordem da linha
        strLinha = PreencherCampo(CStr(intLinha), intTamanhoNumeroSequencial, tipInteiro)
        '2 - Primeiro campo
        strValorCampo = IIf(Nvl(grdOrdem.ListItems(1).SubItems(2), 0) = 3, Format(grdResultado.ListItems(intLinha), Const_Monetario), grdResultado.ListItems(intLinha))
        strLinha = strLinha & Space(cteEspacamentoColunas) & PreencherCampo(strValorCampo, Nvl(grdOrdem.ListItems(1).SubItems(1), 0), Nvl(grdOrdem.ListItems(1).SubItems(2), 0))
        '3 - Demais campos
        For intColuna = 1 To intQuantidadeColunas - 1
            strValorCampo = IIf(Nvl(grdOrdem.ListItems(1).SubItems(2), 0) = 3, Format(grdResultado.ListItems(intLinha).SubItems(intColuna), Const_Monetario), grdResultado.ListItems(intLinha).SubItems(intColuna))
            strLinha = strLinha & Space(cteEspacamentoColunas) & PreencherCampo(strValorCampo, Nvl(grdOrdem.ListItems(intColuna + 1).SubItems(1), 0), Nvl(grdOrdem.ListItems(intColuna + 1).SubItems(2), 0))
        Next intColuna
        '4 - Imprime a linha
        Printer.Print strLinha
    Next intLinha
    Exit Sub
trata:
    Erro Err.Description
    Exit Sub
    Resume
End Sub

Public Sub ImprimirRodape(NumeroPagina As Integer, TotalPaginas As Integer, LarguraMaxima As Integer, Total As Integer)
    Dim strLinha As String

    Printer.Font.Name = "Courier New"
    Printer.Font.Size = txtTamanhoFonte
    '1 - Usuario / pagina
    strLinha = Aplicacoes.Usuario & Space(8) & "Página " & NumeroPagina & " de " & TotalPaginas
    Printer.Print IIf(Total > 0, AlinharCampo("* * *   FIM DE RELATÓRIO - " & Total & " REGISTRO(S)   * * *", aliCentro, LarguraMaxima), "")
    Printer.Print AlinharCampo(strLinha, aliDireita, LarguraMaxima)
    strLinha = String(Len(strLinha), "-")
    Printer.Print AlinharCampo(strLinha, aliDireita, LarguraMaxima)
    '2 - Data / hora
    strLinha = Format(Now, "dd/mm/yyyy") & " - " & Format(Now, "hh:mm:ss")
    Printer.Print AlinharCampo(strLinha, aliDireita, LarguraMaxima)
End Sub

Private Sub txtLarguraMaxima_KeyPress(KeyAscii As Integer)
' Tamanho do papel, ou seja, num. colunas
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLinhasPorPagina_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLinhasPorPagina_LostFocus()
    'Total de paginas que o relatorio vai usar
    txtTotalPaginas = CalcularTotalPaginas(grdResultado.ListItems.Count, txtLinhasPorPagina)
    txtPaginaInicial = 1
    txtPaginaFinal = txtTotalPaginas
End Sub

Private Sub txtPaginaFinal_Change()
    If Val(txtPaginaFinal) > Val(txtTotalPaginas) Then
        txtPaginaFinal = txtTotalPaginas
    End If
End Sub

Private Sub txtPaginaFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPaginaInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtTamanhoFonte_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtTamanhoFonte_LostFocus()
    If Not IsNumeric(txtTamanhoFonte) Then txtTamanhoFonte = 10
    If CInt(txtTamanhoFonte) > 20 Then
        Avisa "O tamanho máximo da fonte é 20."
        txtTamanhoFonte = 20
    End If
    If CInt(txtTamanhoFonte) < 5 Then
        Avisa "O tamanho mínimo da fonte é 5."
        txtTamanhoFonte = 5
    End If
End Sub

Private Sub txtTituloRelatorio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

Private Function CalcularTotalPaginas(QuantidadeRegistros As Variant, LinhasPorPagina As Variant) As Integer
    Dim dblQuantidadePaginas As Double
    
    dblQuantidadePaginas = QuantidadeRegistros / LinhasPorPagina
    CalcularTotalPaginas = CInt(dblQuantidadePaginas + 0.5) ' by Silmar Bosing

End Function

Private Function CalcularLarguraNecessaria(intQtdRegistros As Integer, intQtdColunas As Integer, intLarguraAtual As Integer) As Integer
    Dim intTamanhoEspacoColunas As Integer
    
    If CInt(Len(CStr(grdResultado.ListItems.Count))) <= MaxTamCont Then
        intTamanhoNumeroSequencial = Len(CStr(intQtdRegistros))
    Else
        intTamanhoNumeroSequencial = 0
    End If
    intTamanhoEspacoColunas = (intQtdColunas - 1) * cteEspacamentoColunas
    CalcularLarguraNecessaria = intLarguraAtual + intTamanhoNumeroSequencial
End Function

Private Function BuscaAliasUtilizados(grdCampos As Object, grdFiltros As Object) As String
   Dim i As Integer, j As Integer, k As Integer
    
    BuscaAliasUtilizados = ""
    j = grdCampos.ListItems.Count
    For i = 1 To j
        BuscaAliasUtilizados = BuscaAliasUtilizados & IIf(Len(BuscaAliasUtilizados) > 0, ",", "") & "'" & grdCampos.ListItems(i) & "'"
    Next
    j = grdFiltros.ListItems.Count
    For i = 1 To j
        BuscaAliasUtilizados = BuscaAliasUtilizados & IIf(Len(BuscaAliasUtilizados) > 0, ",", "") & "'" & grdFiltros.ListItems(i) & "'"
    Next
End Function


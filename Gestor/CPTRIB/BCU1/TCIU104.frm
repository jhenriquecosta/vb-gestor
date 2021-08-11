VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU104"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleMode       =   0  'User
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   14
      Top             =   3075
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5115
         TabIndex        =   25
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3435
         TabIndex        =   9
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         Caption         =   "&Gerar Valores"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6285
         TabIndex        =   10
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1695
      Left            =   75
      TabIndex        =   13
      Top             =   720
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2990
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lblCont 
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   2250
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   503
         _Version        =   196610
         ForeColor       =   8388608
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   11
         Left            =   3060
         TabIndex        =   22
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cod Logr"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   6
         Left            =   5385
         TabIndex        =   21
         Top             =   960
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tipo Imóvel"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   8
         Left            =   1740
         TabIndex        =   20
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Insc. Municipal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   49
         Left            =   105
         TabIndex        =   19
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Insc. Cadastral"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   7
         Left            =   3060
         TabIndex        =   18
         Top             =   960
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Setor"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   4020
         TabIndex        =   17
         Top             =   960
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Quadra"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   16
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Bairro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   4020
         TabIndex        =   15
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   196610
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Logradouro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtic 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   105
         TabIndex        =   0
         Tag             =   "TIM_IC"
         Top             =   600
         Width           =   1560
      End
      Begin VB.ComboBox cboBairro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TCIU104.frx":0000
         Left            =   105
         List            =   "TCIU104.frx":000D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "TBA_NOME"
         Top             =   1207
         Width           =   2910
      End
      Begin VB.ComboBox cboLogr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TCIU104.frx":002E
         Left            =   5385
         List            =   "TCIU104.frx":003B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "TLG_NOME"
         Top             =   600
         Width           =   1920
      End
      Begin VB.ComboBox cboTipoLogr 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TCIU104.frx":005C
         Left            =   4020
         List            =   "TCIU104.frx":005E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "TTL_NOME"
         Top             =   600
         Width           =   1305
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4020
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "TIM_IC"
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtSecao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "TIM_IC"
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtIM 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1725
         MaxLength       =   11
         TabIndex        =   1
         Tag             =   "TIM_TCI_IM"
         Top             =   600
         Width           =   1275
      End
      Begin VB.ComboBox cboTipo 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TCIU104.frx":0060
         Left            =   5385
         List            =   "TCIU104.frx":006A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "TTL_NOME"
         Top             =   1207
         Width           =   1920
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3060
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "TIM_IC"
         Top             =   600
         Width           =   915
      End
      Begin Threed.SSPanel lblIsentos 
         Height          =   330
         Left            =   15
         TabIndex        =   24
         Top             =   2640
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   582
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   12045006
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         X1              =   0
         X2              =   7380
         Y1              =   2625
         Y2              =   2625
      End
   End
   Begin VB.PictureBox prgIptu 
      Height          =   195
      Left            =   4545
      ScaleHeight     =   135
      ScaleWidth      =   2670
      TabIndex        =   12
      Top             =   1875
      Width           =   2730
   End
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   2550
      Top             =   1050
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   195
      Left            =   330
      TabIndex        =   26
      Top             =   2730
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Threed.SSPanel lblStatus 
      Height          =   240
      Left            =   330
      TabIndex        =   27
      Top             =   2490
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   423
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0%"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   3
      Alignment       =   6
      RoundedCorners  =   0   'False
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   180
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1138
      Icone           =   "TCIU104.frx":0084
   End
End
Attribute VB_Name = "TCIU104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim cadastro As VSImposto
Dim Cobranca As New VSCobranca
Dim CalculoIptu As New VSIptu
Dim sqlValores As String
Private Function AbreSelecao(Record As VSRecordset) As Boolean
    ' BUSCA OS IMÓVEIS PARA O QUAL SERÁ GERADO O IMPOSTO
    Dim Query As String
    Dim Tabela As String
    Dim Operador As String
    Dim Sql As String
    'PARA GERAR  O IMPOSTO
    Tabela = "SELECT * FROM  Vis_Imovel ,Tab_Contribuinte where tim_tci_im=tci_im AND TIM_SITUACAO_LOTE <> 1"
    
    sqlValores = "SELECT tgt_tim_ic as IC, tip_sigla_imposto as Imposto, tgt_valor_tributo + tgt_taxa_expediente as Valor" & _
        " FROM TAB_GERACAO_TRIBUTO, VIS_IMOVEL, TAB_CONTRIBUINTE, TAB_IMPOSTO" & _
        " WHERE tgt_tim_ic=tim_ic AND tgt_im=tci_im AND tgt_tip_cod_imposto=tip_cod_imposto AND tgt_tip_cod_imposto IN ('" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "', '" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITU)) & "') "
    If Trim(txtIC) <> "" Then
        If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
            If CInt(Right(txtIC, 3)) <> 200 Then
                Sql = Sql & " and tim_ic ='" & txtIC & "'"
            Else
                Sql = Sql & " and tim_ic > '" & txtIC & "' and  tim_ic  <= '" & Left(txtIC, 12) & "300'"
            End If
        Else
            Sql = Sql & " and tim_ic ='" & txtIC & "'"
        End If
    End If
    If Trim(txtIM) <> "" Then
        Sql = Sql & " and tim_tci_im = '" & txtIM & "'"
    End If
    If Trim(cboTipoLogr) <> "" Then
        Sql = Sql & " and TTL_NOME = '" & cboTipoLogr & "'"
    End If
    If Trim(cboLogr) <> "" Then
        Sql = Sql & " and tlg_nome = '" & cboLogr & "'"
    End If
    If Trim(cboBairro) <> "" Then
        Sql = Sql & " and TBA_NOME = '" & cboBairro & "'"
    End If
    
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
        If Trim(txtSecao) <> "" Then
            Sql = Sql & " and tim_ic like '__" & txtSecao & "%'"
        End If
        If Trim(txtQuadra) <> "" Then
            Sql = Sql & " and tim_ic like '____0" & txtQuadra & "%'"
        End If
    End If
    'GAMBIARRA SILMAR
'    Sql = Sql & " AND TIM_IC NOT IN (SELECT DISTINCT tgt_tim_ic FROM Tab_Geracao_Tributo WHERE tgt_data_vencimento = CONVERT(DATETIME, '30/12/2002', 103) AND tgt_tip_cod_imposto = '11120200')"
    Tabela = Tabela & Sql & _
    "  order by tim_ic ASC ,tim_unidade ASC"
    sqlValores = sqlValores & Sql
    Screen.MousePointer = 11
    Tabela = Replace(Tabela, "TTL_NOME", "TIPOLOGRADOURO")
     Tabela = Replace(Tabela, "tlg_nome", "LOGRADOURO")
     
    AbreSelecao = Bdados.AbreTabela(Tabela, Record, Estatico, SomenteLeitura) '(tabela)
    
End Function

Private Sub ImprimeBoleto()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsComp As VSRecordset
    Dim Operador As String
    Dim aux As Byte
    Dim AreaTotal As String
    Dim AreaConstruida As String
    Dim Desconto As Double
    Dim RsDesconto As VSRecordset
    Dim RsAux As VSRecordset
    Dim ValorMetro As Double
    Dim NomeLogr As String
    Dim Logr As String
    Dim CodigoLogr As String
    Dim Bairro As String
    Dim Cobranca As New VSCobranca
    Dim Conta As New ContaCorrente
    Dim CodImposto As String
    Dim Controle As Control
    Dim NomeImposto As String
    Dim Sigla As String
    Dim EnderecoCont As String
    
    Screen.MousePointer = 11
    aux = 0
    CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    Sql = "select tip_sigla_imposto, tip_nome_imposto from tab_imposto where tip_cod_imposto ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        Sigla = rs!TIP_sigla_IMPOSTO
        NomeImposto = rs!tip_nome_imposto
    End If
    Bdados.FechaTabela rs
    Sql = "SELECT * FROM VIS_IMOVEL,tab_geracao_tributo where tgt_tim_ic = tim_ic and tgt_tip_cod_imposto ='" & CodImposto & "' and tgt_im = tim_tci_im "
    For Each Controle In Controls
        If Controle.Tag <> "" Then
            If Trim(Controle.Text) <> "" Then
                Sql = Sql & " and " & Controle.Tag & " = '" & Trim(Controle.Text) & "'"
            End If
        End If
    Next
    Sql = Sql & " ORDER BY tim_IC ASC"
    If Bdados.AbreTabela(Sql, rs) Then
'        MontaGrid Bdados,lstIptu, Sql, 1400
        
        DoEvents
        rs.MoveFirst
        Sql = "Select TGE_NOME from tab_geral where TGE_TIPO = 755 and TGE_CODIGO > 0"
        If Bdados.AbreTabela(Sql, RsDesconto) Then
            Desconto = RsDesconto(0)
        End If
        Do While Not rs.EOF
            Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & rs!TIM_IC & _
                "' and (tdi_tco_cod_componente=110 or tdi_tco_cod_componente=108)"
            If Bdados.AbreTabela(Sql, RsComp) Then
                RsComp.MoveFirst
                Do While Not RsComp.EOF
                    If RsComp(0) = 110 Then
                        AreaTotal = RsComp(1)
                    ElseIf RsComp(0) = 108 Then
                        AreaConstruida = RsComp(1)
                    End If
                    RsComp.MoveNext
                Loop
            End If
            Bdados.FechaTabela RsComp
            
            Sql = "select tvl_valor  as ValorMetro from TAB_VALOR_TERRENO where tvl_tlg_cod_logradouro='" & rs!tim_tlg_cod_logradouro & "'"
            If Bdados.AbreTabela(Sql, RsAux) Then
                ValorMetro = RsAux!ValorMetro
            End If
            EnderecoCont = rs!tci_logradouro & " " & rs!tci_nome_logradouro & " " & rs!tci_NUMERO & " " & rs!tci_BAIRRO
            Bdados.FechaTabela RsAux
            Sql = "select ttl_nome as Logr,TLG_NOME AS Nome  from tab_logradouro,tab_tipo_logr where tlg_cod_logradouro='" & rs!tim_tlg_cod_logradouro & "' and tlg_ttl_cod_tip_logr = ttl_cod_tip_logr "
            If Bdados.AbreTabela(Sql, RsAux) Then
                Logr = RsAux!Logr
                NomeLogr = RsAux!Nome
            End If
            Bdados.FechaTabela RsAux
            Cobranca.ImprimeDam Rpt, rs!tgt_cod_pagamento, rs!TGT_im, rs!tci_nome, "", EnderecoCont, _
             rs!TIM_IC, rs!TTL_NOME & " " & " " & rs!tlg_nome & " " & rs!tim_numero & " " & rs!TBA_NOME, CodImposto, _
            Sigla, NomeImposto, rs!TGT_PERIODO, IIf(rs!TGT_PARCELA = 0, "UNICA", rs!TGT_PARCELA), 1, rs!TGT_DATA_VENCIMENTO, rs!tim_valor, rs!TGT_VALOR_TRIBUTO, rs!TGT_VALOR_MULTA, rs!tgt_valor_juros, rs!tgt_taxa_expediente, 0, "", "", , , , , ValorMetro, , CDbl(AreaTotal), CDbl(AreaConstruida), , , , tdiImpressora
'            Util.Pausa 1000
            AreaTotal = 0
            AreaConstruida = 0
            rs.MoveNext
        Loop
    Else
        Avisa "Nenhum Registro encontrado."
    End If
    Screen.MousePointer = 0
    Bdados.FechaTabela rs
    Bdados.FechaTabela RsAux
    Bdados.FechaTabela RsComp
End Sub


Private Sub cmd_Click(Index As Integer)
    
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
End Sub



Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Calculo As New CalculoObrigacao
    Dim Sql As String
    Dim rs As VSRecordset
    
    Screen.MousePointer = 11
    If Confirma("Confirma o cálculo do valor venal?") Then
        If AbreSelecao(rs) Then
            BarraProgresso.Visible = True
            rs.MoveFirst
            If rs.RecordCount > 0 Then BarraProgresso.Max = rs.RecordCount
            lblStatus = "Executando  0%"
            Do
                If AplicacoesVTFuncoes.municipio = "PETROLINA" Then
                    Call Calculo.IptuSinfaz(Trim(rs!TIM_IC), Year(Date), Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), "", True)
                Else
                    Call Calculo.Iptu(Trim(rs!TIM_IC), Year(Date), Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), "", True)
                End If
                rs.MoveNext
                If rs.EOF = False Then
                    BarraProgresso.Value = rs.AbsolutePosition
                    lblStatus = "Executando  " & CInt((BarraProgresso.Value * 100) / BarraProgresso.Max) & "%"
                End If
                DoEvents
            Loop While Not rs.EOF
        End If
    End If
    Screen.MousePointer = 0
    Avisa "Cálculo(s) finalizado(s)."
End Sub
Private Sub Form_Load()
            
    Dim Controle As Control
    Dim i As Byte
    Set cadastro = New VSImposto
    Call Edita.AtualizaCombo(Bdados, cboLogr, "Select DISTINCT(tlg_nome) From Tab_Logradouro ")
    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr")
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select DISTINCT(tba_nome) From Tab_Bairro ")
    'txtMens = Temp.PegaParametro(Bdados, "MENSAGEM IPTU")
    cboLogr.AddItem ""
    cboTipoLogr.AddItem ""
    cboBairro.AddItem ""
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub txtAnoFi_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtAnoIni_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus()
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        txtIC = cadastro.FormataInscricao(txtIC, InscImovel)
    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    txtIM = cadastro.FormataInscricao(txtIM, InscContrib)
End Sub


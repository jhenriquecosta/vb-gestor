VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TOBR408 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   27
      Top             =   7080
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8715
         TabIndex        =   14
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   9870
         TabIndex        =   15
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   1
         Left            =   2595
         TabIndex        =   28
         Top             =   180
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   318
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
         Caption         =   "Relatório"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.ComboBox cboRelatorio 
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
         ItemData        =   "TOBR408.frx":0000
         Left            =   3390
         List            =   "TOBR408.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   120
         Width           =   3585
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   6990
         TabIndex        =   13
         Top             =   90
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL lstImp 
      Height          =   1830
      Left            =   30
      TabIndex        =   11
      Top             =   5250
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   3228
      Caption         =   "Totais"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Threed.SSFrame fra 
      Height          =   1740
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   660
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   3069
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Consulta"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.ComboBox cboCotas 
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
         ItemData        =   "TOBR408.frx":003F
         Left            =   7260
         List            =   "TOBR408.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1290
         Width           =   1560
      End
      Begin VB.ComboBox cborestricao 
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
         ItemData        =   "TOBR408.frx":0068
         Left            =   1875
         List            =   "TOBR408.frx":0075
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1290
         Width           =   3060
      End
      Begin VB.TextBox txtExercicio1 
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
         Left            =   7260
         MaxLength       =   8
         TabIndex        =   5
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox txtExercicio2 
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
         Left            =   8790
         MaxLength       =   8
         TabIndex        =   6
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox txtIc 
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
         Left            =   1890
         MaxLength       =   15
         TabIndex        =   2
         Top             =   930
         Width           =   2055
      End
      Begin VB.ComboBox cboImposto 
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
         Left            =   1890
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   562
         Width           =   1515
      End
      Begin VB.TextBox txtPeriodo2 
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
         Left            =   8790
         MaxLength       =   12
         TabIndex        =   4
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox txtPeriodo1 
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
         Left            =   7260
         MaxLength       =   12
         TabIndex        =   3
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox txtRazao 
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
         Left            =   4230
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   210
         Width           =   6030
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
         Height          =   315
         Left            =   1890
         MaxLength       =   13
         TabIndex        =   0
         Top             =   210
         Width           =   1485
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   11
         Left            =   315
         TabIndex        =   18
         Top             =   255
         Width           =   1530
         _ExtentX        =   2699
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
         Caption         =   "Inscrição Muncipal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   5415
         TabIndex        =   19
         Top             =   615
         Width           =   1755
         _ExtentX        =   3096
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
         Caption         =   "Período(dd/mm/aaaa)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   1155
         TabIndex        =   20
         Top             =   615
         Width           =   690
         _ExtentX        =   1217
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
         Caption         =   "Imposto"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   22
         Top             =   975
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Inscrição Cadastral"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   6435
         TabIndex        =   23
         Top             =   975
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "Exercício"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   1035
         TabIndex        =   24
         Top             =   1350
         Width           =   780
         _ExtentX        =   1376
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
         Caption         =   "Restrição"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   8970
         TabIndex        =   9
         Top             =   1260
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   6
         Left            =   6630
         TabIndex        =   26
         Top             =   1350
         Width           =   480
         _ExtentX        =   847
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
         Caption         =   "Cotas"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3390
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   210
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaIC 
         Height          =   315
         Left            =   3960
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   930
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   900
      TabIndex        =   21
      Top             =   1500
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   1138
      Icone           =   "TOBR408.frx":00A7
   End
   Begin VTOcx.grdVISUAL lstIptu 
      Height          =   2820
      Left            =   15
      TabIndex        =   10
      Top             =   2415
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4974
      Caption         =   "Valores Lançados"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.Menu OpcoesDam 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "Reimprimir DAM"
      End
      Begin VB.Menu mnuEstorno 
         Caption         =   "Estornar DAM"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "TOBR408"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As New VSImposto
Dim String_Taxas As String
Dim Total_Taxas As Double
Private Sub cmd_Click(Index As Integer)
    On Error Resume Next
    Dim Sql                 As String
    Dim rs                  As VSRecordset
    Dim Condicao            As String
    Dim CondicaoRelat       As String
    Static Imposto          As String
    Dim Tabelas             As String
    Dim CalusuraWhere       As String
    Dim i                   As Integer
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Screen.MousePointer = 11
            Condicao = ""
            Tabelas = " Tab_Geracao_Tributo INNER JOIN Tab_Imposto ON Tab_Geracao_Tributo.tgt_tip_cod_imposto = Tab_Imposto.tip_cod_imposto LEFT OUTER JOIN Tab_Darm_Recebido ON Tab_Geracao_Tributo.tgt_cod_pagamento = Tab_Darm_Recebido.tdr_tgt_cod_pagamento "
            Sql = "Select tgt_im as [IM], tgt_tim_ic as [IC], "
            Sql = Sql & " tgt_periodo as Periodo,tgt_data_vencimento as Vencimento,tip_cod_imposto as [Cod Tributo], "
            Sql = Sql & " tip_sigla_imposto as Descricao," & FuncaoReal("tgt_valor_tributo") & " as [Vl Tributo], "
            Sql = Sql & FuncaoReal("tgt_taxa_expediente") & " as [Taxas],"
            Sql = Sql & "  tgt_cod_pagamento as [Num DOC], tgt_cod_pagamento_vinculado as [Doc Vinculo], tgt_cod_pagamento_original as [Doc Origem],tgt_parcela as [Cota],TGT_TIPO AS Título "
            Sql = Sql & " from "
            
            Condicao = " where  tgt_tip_cod_imposto=tip_cod_imposto and (tgt_ativo =0 or tgt_ativo is null)"
            If Trim(txtIm) <> "" Then
                Condicao = Condicao & " and tgt_im = '" & txtIm & "'"
            End If
             If cboImposto.ListIndex > 0 Then
                Imposto = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto='" & cboImposto & "'")
                Condicao = Condicao & " and tgt_tip_cod_imposto  = '" & Imposto & "'"
            End If
            If Trim(txtPeriodo1) <> "" And Trim(txtPeriodo2) <> "" Then
                Condicao = Condicao & " and tgt_data_geracao >= " & Bdados.Converte(txtPeriodo1, TCDataHora) & " and tgt_data_geracao <= " & Bdados.Converte(txtPeriodo2, TCDataHora)
            End If
            If Trim(txtExercicio1) <> "" And Trim(txtExercicio2) <> "" Then
                Condicao = Condicao & " and tgt_periodo >= " & IIf(Len(txtExercicio1) = 4, txtExercicio1, Right(txtExercicio1, 4) & Left(txtExercicio1, 2)) & " and tgt_periodo <= " & IIf(Len(txtExercicio2) = 4, txtExercicio2, Right(txtExercicio2, 4) & Left(txtExercicio2, 2))
            End If
            
            If Trim(txtic) <> "" Then
                Condicao = Condicao & " and tgt_tim_ic LIKE '" & Trim(txtic) & "%'"
            End If
            
            Condicao = Condicao & " and tgt_tip_cod_imposto not in ('" & Const_Notificacao & "','" & Const_Extrato & "')"
            If cboRestricao.ListIndex = 0 Then
                Condicao = Condicao & " and Tab_Darm_Recebido.tdr_tgt_cod_pagamento is null"
            ElseIf cboRestricao.ListIndex = 1 Then
                Condicao = Condicao & " and Tab_Darm_Recebido.tdr_tgt_cod_pagamento is not null "
            End If
            
            If cboCotas.ListIndex = 0 Then
                Condicao = Condicao & " and Tab_Geracao_Tributo.tgt_parcela = 0"
            ElseIf cboCotas.ListIndex = 1 Then
                Condicao = Condicao & " and Tab_Geracao_Tributo.tgt_parcela <> 0"
            End If
            Sql = Sql & Tabelas & Condicao
            Sql = Sql & " AND TIP_COD_IMPOSTO not in ('" & Const_Notificacao & "','" & Const_Extrato & "')"
            Sql = Sql & " ORDER BY tgt_periodo, tip_sigla_imposto"
            If Bdados.AbreTabela(Sql, rs) Then
                lstIptu.Preencher Bdados, Sql, 1200, 1500, 800
                For i = 1 To lstIptu.ListItems.Count
                    lstIptu.ListItems(i).SubItems(12) = Pega_Tipo(lstIptu.ListItems(i).SubItems(12))
                Next
                lstIptu.Mensagem = "Total: R$ " & Format(lstIptu.Colunas(7).Soma, Const_Monetario)
                Sql = "Select tgt_tip_cod_imposto as [Cod Tributo],tip_sigla_imposto as Descricao,"
                Sql = Sql & Bdados.Converte("Sum(tgt_valor_tributo)", TCDuplo) & " as [Vl Tributo], Count(*) as Documentos "
                Sql = Sql & " from " & Tabelas
                
                Sql = Sql & Condicao
                Sql = Sql & " AND TIP_COD_IMPOSTO not in ('" & Const_Notificacao & "','" & Const_Extrato & "')"
                Sql = Sql & " group by tgt_tip_cod_imposto,tip_sigla_imposto"
                lstImp.Preencher Bdados, Sql, 1400
                lstImp.Mensagem = "Total de Documentos: " & lstImp.Colunas(4).Soma
                DoEvents
                Sql = "SELECT SUM(tgt_valor_tributo),count(*) FROM " & Tabelas & " where TGT_PARCELA = 0"
                Sql = Sql & IIf(Trim(Condicao) <> "", " and ", " ") & Mid(Condicao, 7)

            Else
                lstImp.Mensagem = ""
                lstIptu.Mensagem = ""
                Util.Mensagem "Nenhum registro encontrado."
                Bdados.FechaTabela rs
                lstImp.ListItems.Clear
                lstIptu.ListItems.Clear
                Screen.MousePointer = 0
                Exit Sub
            End If
            Bdados.FechaTabela rs
            Screen.MousePointer = 0
            Exit Sub
        Case 2      ' IMPRESSÃO
            Screen.MousePointer = 11
            CondicaoRelat = ""
            If Trim(txtIm) <> "" Then
                CondicaoRelat = " and {Tab_Geracao_Tributo.tgt_im} = '" & txtIm & "'"
            End If
            If cboImposto.ListIndex > 0 Then
                Imposto = BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto='" & cboImposto & "'")
                CondicaoRelat = CondicaoRelat & " and {Tab_Geracao_Tributo.tgt_tip_cod_imposto}  = '" & Imposto & "'"
            End If
            If Trim(txtPeriodo1) <> "" And Trim(txtPeriodo2) <> "" Then
                CondicaoRelat = CondicaoRelat & " and {Tab_Geracao_Tributo.tgt_data_geracao} in  Date (" & Year(txtPeriodo1) & "," & Month(txtPeriodo1) & "," & Day(txtPeriodo1) & ") to Date (" & Year(txtPeriodo2) & "," & Month(txtPeriodo2) & "," & Day(txtPeriodo2) & ")"
            End If
            If Trim(txtic) <> "" Then
                CondicaoRelat = CondicaoRelat & " and {Tab_Geracao_Tributo.tgt_tim_ic} like '" & txtic & "*'"
            End If
            If Trim(txtExercicio1) <> "" And Trim(txtExercicio2) <> "" Then
                CondicaoRelat = CondicaoRelat & " and {Tab_Geracao_Tributo.tgt_periodo} >= " & IIf(Len(txtExercicio1) = 4, txtExercicio1, Right(txtExercicio1, 4) & Left(txtExercicio1, 2)) & " and {Tab_Geracao_Tributo.tgt_periodo} <= " & IIf(Len(txtExercicio2) = 4, txtExercicio2, Right(txtExercicio2, 4) & Left(txtExercicio2, 2))
            End If
            With Rpt
                If cboRelatorio.ListIndex = -1 Then
                    If Not .DefinirArquivo(Bdados, App.Path & "\TResumoLancamentoGeral.rpt") Then Exit Sub
                    If Trim(txtPeriodo1) <> "" And Trim(txtPeriodo2) <> "" Then
                        .Selecao = "( {Tab_Geracao_Tributo.tgt_data_geracao} in  Date (" & Year(txtPeriodo1) & "," & Month(txtPeriodo1) & "," & Day(txtPeriodo1) & ") to Date " & _
                                    "(" & Year(txtPeriodo2) & "," & Month(txtPeriodo2) & "," & Day(txtPeriodo2) & ") AND {TAB_GERACAO_TRIBUTO.TGT_PARCELA} = 0) "
                        .Formulas "FILTRO ", "PERÍODO : " & txtPeriodo1 & " a " & txtPeriodo2
                    Else
                        .Selecao = "{TAB_GERACAO_TRIBUTO.TGT_PARCELA} = 0 " & CondicaoRelat
                        
                        .Formulas "FILTRO", "TODOS OS LANÇAMENTOS ATÉ " & Date
                    End If
                    
                Else
                    If Aplicacoes.Codigo_Municipio = 1 And cboImposto = "IPTU" Then
                        If Not .DefinirArquivo(Bdados, App.Path & "\TXavier.rpt") Then Exit Sub
                    Else
                        If Not .DefinirArquivo(Bdados, App.Path & "\TDamLancado.rpt") Then Exit Sub
                    End If
                    
                    CondicaoRelat = CondicaoRelat & " and not ({Tab_Geracao_Tributo.tgt_tip_cod_imposto} like ['" & Const_Extrato & "', '" & Const_Notificacao & "'])" & _
                                                    " and ({Tab_Geracao_Tributo.tgt_ativo} =0 or isnull({Tab_Geracao_Tributo.tgt_ativo}))"
                    If cboRelatorio.ListIndex = 0 Then
                        .Selecao = " isnull({Tab_Darm_Recebido.tdr_tgt_cod_pagamento}) " & CondicaoRelat
                    ElseIf cboRelatorio.ListIndex = 1 Then
                        .Selecao = " not isnull({Tab_Darm_Recebido.tdr_tgt_cod_pagamento}) " & CondicaoRelat
                    Else
                        .Selecao = Mid(CondicaoRelat, 5)
                    End If
                End If
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                .Arvore = False
                .Visualizar
            End With
            Screen.MousePointer = 0
            Set Rpt = Nothing
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    lstImp.Preencher Bdados, ""
    lstIptu.Preencher Bdados, ""
    txtIm.SetFocus
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    AplicacoesVTFuncoes.BuscaNoEconomico TcoJuridica, txtIm
End Sub

Private Sub cmdPesquisaIC_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtic
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
     Call Edita.AtualizaCombo(Bdados, cboImposto, "select distinct(TIP_sigla_IMPOSTO) from tab_imposto")
    cboImposto.AddItem " "
    'Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    AtualizaCabecalho lstImp
    AtualizaCabecalho lstIptu
    txtRazao.Locked = True
End Sub

Private Sub lstIptu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 And lstIptu.ListItems.Count > 0 Then
        mnuEstorno.Caption = "Estornar DAM " & lstIptu.SelectedItem.SubItems(8)
        mnuReimprime.Caption = "Imprimir  DAM " & lstIptu.SelectedItem.SubItems(8)
        'GAMBIARRA
        mnuEstorno.Enabled = podeEstornar()
        'GAMBIARRA
        Me.PopupMenu OpcoesDam
    End If
End Sub

Private Function podeEstornar() As Boolean
    podeEstornar = Bdados.AbreTabela("SELECT * FROM TAB_ACESSO_USUARIO WHERE TAU_TMO_COD_MODULO ='TCOB' and TAU_TFO_COD_FORMULARIO =301 AND TAU_TUS_COD_USUARIO='" & Aplicacoes.Usuario & "'")
    Bdados.FechaTabela
End Function

Private Sub mnuReimprime_Click()
    'MUDAR AQUI
'    TCOB204.Show
'    TCOB204.txtDAM = lstIptu.SelectedItem.SubItems(8)
'    TCOB204.txtDAM.SetFocus
    SendKeys "{tab}"
    'Aplicacoes.Abre_Aplicacao "TCOB204", 0, Cod_sis, Sistema, Desc_Form
    'TCOB204.txtObservacao.SetFocus
End Sub

Private Sub Timer1_Timer()
    FocalizaCaixa Me
End Sub

Private Sub txtExercicio1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio1_LostFocus()
    If Len(txtExercicio1) = 6 Then
        txtExercicio1 = Left(txtExercicio1, 2) & "/" & Right(txtExercicio1, 4)
    End If
End Sub

Private Sub txtExercicio2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio2_LostFocus()
    If Len(txtExercicio2) = 6 Then
        txtExercicio2 = Left(txtExercicio2, 2) & "/" & Right(txtExercicio2, 4)
    End If
End Sub

Private Sub txtic_LostFocus()
    If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
        txtic = Imposto.FormataInscricao(txtic, InscImovel)
    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtIm) <> "" Then
        If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
            txtIm = cadastro.FormataInscricao(txtIm, InscContrib)
        End If
        Sql = "SELECT tci_nome FROM Tab_Contribuinte where tci_im='" & txtIm & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtRazao = rs(0)
        Else
            Call Avisa("Contribuinte não encontrado.")
            txtIm = ""
            txtIm.SetFocus
        End If
    Else
        txtRazao = ""
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtPeriodo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo1_LostFocus()
    If IsNumeric(txtPeriodo1) Then
        txtPeriodo1 = Edita.FormataTexto(txtPeriodo1, Data)
    End If
End Sub

Private Sub txtPeriodo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo2_LostFocus()
    If IsNumeric(txtPeriodo2) Then
        txtPeriodo2 = Edita.FormataTexto(txtPeriodo2, Data)
    End If
End Sub

Private Function Pega_Tipo(Cod As String) As String
    Dim Sql As String
    Sql = "Select tge_nome from vis_tipo_dam where tge_codigo = " & Bdados.Converte(Cod, tctexto)
    If Bdados.AbreTabela(Sql) Then
        Pega_Tipo = Bdados.Tabela(0)
    End If
End Function




VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Begin VB.Form TCTA101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraMsg 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   4155
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   3570
      Begin VB.Label LblMsg 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Valores Lançados..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   885
         TabIndex        =   27
         Top             =   435
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Consultando"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   1065
         TabIndex        =   26
         Top             =   150
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         Height          =   900
         Left            =   30
         Top             =   30
         Width           =   3510
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   7545
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1005
      CorFundo        =   12632256
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   9240
         TabIndex        =   8
         Top             =   105
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   10440
         TabIndex        =   9
         Top             =   105
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCTA101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1138
      Icone           =   "TCTA101.frx":2123
   End
   Begin VTOcx.grdVISUAL lstTotalPago 
      Height          =   1875
      Left            =   75
      TabIndex        =   16
      Top             =   5670
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4339
      Caption         =   "Valores Resgatados por Tributo"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.grdVISUAL lstPago 
      Height          =   1935
      Left            =   75
      TabIndex        =   15
      Top             =   3660
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   3413
      Caption         =   "Valores Resgatados"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.grdVISUAL lstTotalAberto 
      Height          =   1875
      Left            =   4935
      TabIndex        =   17
      Top             =   5670
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   4339
      Caption         =   "Valores em Aberto por Tributo"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin Threed.SSFrame fra 
      Height          =   1335
      Index           =   2
      Left            =   75
      TabIndex        =   10
      Top             =   630
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   2355
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
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   585
         Width           =   9165
      End
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   315
         Index           =   1
         Left            =   3750
         TabIndex        =   20
         Top             =   210
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   7590
         TabIndex        =   6
         Top             =   960
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
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
         Height          =   315
         Left            =   4455
         MaxLength       =   16
         TabIndex        =   1
         Top             =   225
         Width           =   1545
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
         Left            =   3405
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Width           =   1215
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
         Left            =   2145
         MaxLength       =   10
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtRazao 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   6435
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   4890
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
         Left            =   2160
         MaxLength       =   13
         ScrollBars      =   3  'Both
         TabIndex        =   0
         Top             =   210
         Width           =   1545
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   11
         Left            =   570
         TabIndex        =   11
         Top             =   255
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
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
         Caption         =   "Inscrição Muncipal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   1020
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   397
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
         Caption         =   "Período"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   315
         Index           =   1
         Left            =   9900
         TabIndex        =   7
         Top             =   960
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "&Consultar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   4230
         TabIndex        =   19
         Top             =   270
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   397
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
         Caption         =   "IC"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesqIc 
         Height          =   315
         Left            =   6045
         TabIndex        =   23
         Top             =   225
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   1350
         TabIndex        =   24
         Top             =   615
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   397
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
         Caption         =   "Endereço"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   900
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3015
      Top             =   4620
   End
   Begin VTOcx.grdVISUAL lstIptu 
      Height          =   1665
      Left            =   75
      TabIndex        =   14
      Top             =   1980
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   2937
      Caption         =   "Valores Lançados"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VB.Menu OpcoesDAM 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEstorno 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   ""
      End
   End
   Begin VB.Menu OpcoesPago 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEstornoPago 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReimprimePago 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "TCTA101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As New VSImposto
Dim sql As String

Dim VitIM(5) As String
Dim VitIc(5) As String
Dim VitPeriodo(5) As String
Dim VitImposto(5) As String

Private Sub cboImposto_Click()
    Dim cLSImposto As New VSImposto
    If CboImposto.Text = cLSImposto.NomeTributo(ttr_IPTU) Or CboImposto.Text = cLSImposto.NomeTributo(ttr_ALVARA) Then
        txtPeriodo1.MaxLength = 4
        txtPeriodo2.MaxLength = 4
    Else
        txtPeriodo1.MaxLength = 6
        txtPeriodo2.MaxLength = 6
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Obrig As New obrigacao
    Dim rs As VSRecordset
    Dim Condicao As String
    Static Imposto As String
    Dim Aux As Byte
    Dim SqlPago As String
    Dim SqlResto As String
    Dim CondicaoPago As String
    Condicao = ""
    CondicaoPago = ""
    
    lstTotalAberto.Mensagem = ""
    lstIptu.Mensagem = ""
    lstPago.Mensagem = ""
    lstTotalPago.Mensagem = ""
    FraMsg.Visible = True
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Screen.MousePointer = 11
            Dim i As Byte
            For i = 0 To 4
                VitIM(i) = ""
                VitPeriodo(i) = ""
                VitImposto(i) = ""
                VitIc(i) = ""
            Next
            
            'BUSCANDO VALORES LANCADOS - 1ª grade
            Condicao = ""
            sql = "SELECT TCC_INSCRICAO AS INSCRICAO, TCC_STATUS_CONTA AS STATUS,TIP_SIGLA_IMPOSTO AS TRIBUTO, " & _
                    "TCC_PERIODO AS PERIODO, TCC_IMPOSTO_ORIGINAL AS VL_ORIG," & _
                    "TCC_DATA_VENCIMENTO AS VENCIMENTO," & _
                    "TCC_IMPOSTO_ATUAL + TCC_JUROS_ATUAL + TCC_MULTA_ATUAL + TCC_CORRECAO_MONETARIA AS VL_CORR," & _
                    "TCC_TDR_VALOR_REAL_PAGO AS TOT_PAGO," & _
                    "TCC_DATA_MOVIMENTO AS ULT_TRANS," & _
                    "TCC_SALDO_ATUAL AS SALDO_DEV,TCC_CREDITO_ATUAL AS SALDO_CRED," & _
                    "TCC_DESCONTO_CONCEDIDO AS DESC,TCC_NAO_TRIBUTADA AS REST,TCC_STATUS_CONTA AS SIT,TCC_PARCELA AS PARCELA " & _
                    "FROM TAB_CONTA_CONTRIBUINTE INNER JOIN TAB_IMPOSTO ON " & _
                    "TCC_TIP_COD_IMPOSTO=TIP_COD_IMPOSTO "
            If Trim(txtPeriodo1) <> "" And Trim(txtPeriodo2) <> "" Then
                Condicao = " and TCC_PERIODO >= " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " and TCC_PERIODO <= " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
                CondicaoPago = " and tdr_periodo >= " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " and tDR_periodo <= " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
                VitPeriodo(0) = "{Vis_Valores_Lancados.Periodo} in " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " to " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
                VitPeriodo(1) = "{Vis_Valores_Resgatados.Periodo} in " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " to " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
                VitPeriodo(2) = "{Vis_Saldo_Conta_Tributaria.Periodo} in " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " to " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
                VitPeriodo(3) = "{Vis_Valores_Abertos.Periodo} in " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " to " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
                VitPeriodo(4) = "{Vis_Total_Vit.Periodo} in " & IIf(Len(txtPeriodo1) = 4, txtPeriodo1, Right(txtPeriodo1, 4) & Left(txtPeriodo1, 2)) & " to " & IIf(Len(txtPeriodo2) = 4, txtPeriodo2, Right(txtPeriodo2, 4) & Left(txtPeriodo2, 2))
            End If
            If Trim(CboImposto) <> "" Then
                Condicao = Condicao & " and TCC_TIP_COD_IMPOSTO = '" & CboImposto.Coluna(0).Valor & "'"
                CondicaoPago = CondicaoPago & " and tdr_tip_cod_imposto = '" & BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & CboImposto & "'") & "'"
                VitImposto(0) = "{Vis_Valores_Lancados.Tributo} = '" & CboImposto & "'"
                VitImposto(1) = "{Vis_Valores_Resgatados.Tributo} = '" & BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & CboImposto & "'") & "'"
                VitImposto(2) = "{Vis_Saldo_Conta_Tributaria.CodReceita} = '" & BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & CboImposto & "'") & "'"
                VitImposto(3) = "{Vis_Valores_Abertos.CodReceita} = '" & BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & CboImposto & "'") & "'"
                VitImposto(4) = "{Vis_Total_Vit.CodReceita} = '" & BuscaCodigo("Select tip_cod_imposto from tab_imposto where tip_sigla_imposto ='" & CboImposto & "'") & "'"
            End If
            If Trim(txtIM) <> "" Then
                Dim SqlImovel As String
                Condicao = Condicao & " and  ( TCC_INSCRICAO = '" & txtIM & "'"
                CondicaoPago = CondicaoPago & " and  ( tdr_INSCRICAO = '" & txtIM & "'"
                If Trim(txtIC) = "" Then
                    SqlImovel = "(select tim_ic from  TAB_IMOVEL where tim_tci_im = '" & txtIM & "')"
                    Condicao = Condicao & " or  TCC_INSCRICAO in " & SqlImovel
                    CondicaoPago = CondicaoPago & " or  tdr_INSCRICAO in " & SqlImovel
                End If
                Condicao = Condicao & ")"
                CondicaoPago = CondicaoPago & ")"
                VitIM(0) = "{Vis_Valores_Lancados.IM} = '" & txtIM & "'"
                VitIM(1) = "{Vis_Valores_Resgatados.IM} = '" & txtIM & "'"
                VitIM(2) = "{Vis_Saldo_Conta_Tributaria.IM} = '" & txtIM & "'"
                VitIM(3) = "{Vis_Valores_Abertos.IM} = '" & txtIM & "'"
                VitIM(4) = "{Vis_Total_Vit.IM} = '" & txtIM & "'"
            End If
            
            If Trim(txtIC) <> "" Then
                Condicao = Condicao & " and  (TCC_INSCRICAO = '" & txtIC & "' or TCC_TIM_IC = '" & txtIC & "')"
                CondicaoPago = CondicaoPago & " and  (tdr_INSCRICAO = '" & txtIC & "' OR TDR_TIM_IC ='" & txtIC & "')"
                
                VitIc(0) = "{Vis_Valores_Lancados.Ic} = '" & txtIC & "'"
                VitIc(1) = "{Vis_Valores_Resgatados.Ic} = '" & txtIC & "'"
                VitIc(2) = "{Vis_Saldo_Conta_Tributaria.Ic} = '" & txtIC & "'"
                VitIc(3) = "{Vis_Valores_Abertos.Ic} = '" & txtIC & "'"
                VitIc(4) = "{Vis_Total_Vit.Ic} = '" & txtIC & "'"
            End If

            Dim Cont As Integer
            Dim Total As Double
'            If Bdados.AbreTabela(sql & Condicao, Rs) Then
        FraMsg.Visible = True
            lstIptu.Mensagem = ""
            sql = ""
 '           Set Obrig = New OBRIGACAO
  '          If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), txtIM, , , txtPeriodoInicial, _
   '                 txtPeriodoFinal, , , , txtImovel, , , txtDAM) Then
    '            Avisa "Nenhum registro encontrado."
     '           cboImposto.SetFocus
      '      End If
            'If lstIptu.Preencher(Bdados, Sql & Condicao) Then
            If Obrig.MostraObrigacaoGerada(lstIptu, CStr(CboImposto.Coluna(0).Valor), txtIM, , , , _
                   , txtPeriodo1, txtPeriodo2, , txtIC) Then
                If lstIptu.ListItems.Count > 0 Then lstIptu.Mensagem = "Total Lançado: R$" & Format(lstIptu.Colunas(5).Soma, Const_Monetario) & "  x  Total Corrigido: R$" & Format(lstIptu.Colunas(7).Soma, Const_Monetario)
                LblMsg.Caption = "Valores Resgatados..."
                
                'BUSCANDO VALORES RESGATADO(PAGOS) - 2ª grade
                sql = "SELECT TDR_INSCRICAO AS CONTRIBUINTE,TIP_SIGLA_IMPOSTO AS TRIBUTO," & _
                 "TDR_PERIODO AS PERIODO,tdr_valor_real_pago as [Vl Pago], " & _
                " tdr_data_pagamento as Pagamento,TLP_TAR_COD_AGENTE AS BANCO,TLP_NUM_SUCURSAL AS AGENCIA, TDR_TLP_COD_LOTE AS Lote ," & _
                "TDR_SEQUENCIA_DAM_LOTE Seq,TDR_PARCELA AS Parcela from TAB_LOTE_PAGAMENTO,tab_darm_recebido,TAB_IMPOSTO " & " where   tdr_sit_pago <> 2 and TDR_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO AND TDR_TLP_COD_LOTE =TLP_COD_LOTE "
                
                lstPago.Preencher Bdados, sql & CondicaoPago
                If lstPago.ListItems.Count > 0 Then lstPago.Mensagem = "Total Resgatado: R$" & Format(lstPago.Colunas(4).Soma, Const_Monetario)
                LblMsg.Caption = "Valores Pago/Aberto"
                'BUSCA TOTALIZADORES EM ABERTO E PAGO
                '3ª grade
                SqlPago = "Select tip_sigla_imposto as Tributo,"
                SqlPago = SqlPago & Bdados.Converte("Sum(tdr_valor_real_pago - tdr_valor_real_juros - tdr_valor_real_multa)", TCDuplo) & " as Valor, "
                SqlPago = SqlPago & Bdados.Converte("Sum(tdr_valor_real_juros)", TCDuplo) & " as Juros ,"
                SqlPago = SqlPago & Bdados.Converte("Sum(tdr_valor_real_multa)", TCDuplo) & " as Multa, "
                SqlPago = SqlPago & Bdados.Converte("Sum(tdr_valor_real_pago)", TCDuplo) & " as Vl_Total "
                SqlPago = SqlPago & " from tab_imposto,tab_darm_recebido "
                SqlPago = SqlPago & " where tdr_tip_cod_imposto = tip_cod_imposto "
                SqlPago = SqlPago & CondicaoPago
                
                lstTotalPago.Preencher Bdados, SqlPago & " group by tip_sigla_imposto"
                
                lstTotalPago.Mensagem = ""
                
                If lstTotalPago.ListItems.Count > 0 Then
                    Total = 0
                    For Cont = 1 To lstTotalPago.ListItems.Count
                        Total = Format(lstTotalPago.ListItems(Cont).SubItems(4) + Total, Const_Monetario)
                    Next
                    lstTotalPago.Mensagem = "Total Resgatado: R$ " & Format(Total, Const_Monetario)
                End If
                '4ª grade
                SqlResto = "Select tip_sigla_imposto as Tributo,"
                SqlResto = SqlResto & Bdados.Converte("Sum(tcc_imposto_original)", TCDuplo) & " as Valor, " & Bdados.Converte("Sum(tcc_juros_atual)", TCDuplo) & " as Juros,"
                SqlResto = SqlResto & Bdados.Converte("Sum(tcc_multa_atual)", TCDuplo) & " as Multa ,"
                SqlResto = SqlResto & Bdados.Converte("Sum(TCC_CORRECAO_MONETARIA)", TCDuplo) & " as Correcao ," & Bdados.Converte("Sum(tcc_imposto_original + tcc_juros_atual + tcc_multa_atual + TCC_CORRECAO_MONETARIA)", TCDuplo) & " as Valor_Total "
                SqlResto = SqlResto & " from tab_imposto,tab_conta_contribuinte "
                SqlResto = SqlResto & " where tcc_tip_cod_imposto = tip_cod_imposto and tcc_saldo_atual <> 0  and tcc_status_conta <> '3'"
                SqlResto = SqlResto & Condicao
                
                lstTotalAberto.Preencher Bdados, SqlResto & " group by tip_sigla_imposto", 700, 700, 700, 700, 1000, 1000
                If lstTotalAberto.ListItems.Count > 0 Then lstTotalAberto.Mensagem = "Imp: R$ " & Format(lstTotalAberto.Colunas(2).Soma, Const_Monetario) & " - Jr: R$ " & Format(lstTotalAberto.Colunas(3).Soma, Const_Monetario) & " - Mt: R$ " & Format(lstTotalAberto.Colunas(4).Soma, Const_Monetario) & " - Cr R$ " & Format(lstTotalAberto.Colunas(5).Soma, Const_Monetario) & " - Tot: R$ " & Format(lstTotalAberto.Colunas(6).Soma, Const_Monetario)
                DoEvents
                FraMsg.Visible = False
            Else
                FraMsg.Visible = False
                Avisa "Nenhum registro encontrado."
                lstIptu.ListItems.Clear
                lstPago.ListItems.Clear
                lstTotalAberto.ListItems.Clear
                lstTotalPago.ListItems.Clear
                Screen.MousePointer = 0
                DoEvents
                Exit Sub
            End If
            Screen.MousePointer = 0
            Bdados.FechaTabela rs
    End Select
    cmdImprimir.Enabled = IIf(lstIptu.ListItems.Count > 0, True, False)
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub CmdImprimir_Click()
Dim Rpt      As VSRelatorio
Dim Condicao As String
Screen.MousePointer = 11
Set Rpt = New VSRelatorio
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TVisaoLancamento.rpt") Then Exit Sub
        .Formulas "Relatorio", "VISÃO INTEGRAL TRIBUTÁRIA"
        .Formulas "VTContribuinte", "'" & txtRazao & "'"
        If txtIM <> "" Then
            .Formulas "VTIm", "'" & txtIM & "'"
        Else
            .Formulas "VTIm", "'" & txtIC & "'"
        End If
        Condicao = " 1 = 1"
        'im - ic - Periodo - tributo
        If txtIM <> "" Then
           Condicao = Condicao & " and  {VIS_INSCRICAO.VIN_INSCRICAO} = '" & txtIM & "'"
        ElseIf txtIC <> "" Then
           Condicao = Condicao & " and  {VIS_INSCRICAO.VIN_INSCRICAO} = '" & txtIC & "'"
        End If
        If txtPeriodo1 <> "" And txtPeriodo2 <> "" Then
            Condicao = Condicao & "  and {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtPeriodo1 & " and {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <=  " & txtPeriodo2
        ElseIf txtPeriodo1 <> "" And txtPeriodo2 = "" Then
            Condicao = Condicao & "  and {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtPeriodo1 & " and {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <=  " & txtPeriodo1
        End If
        If CboImposto.ListIndex >= 0 Then
            Condicao = Condicao & " and {TAB_IMPOSTO.tip_sigla_imposto} = '" & CboImposto.Text & "'"
        End If
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Me.Name, Aplicacoes.Usuario, Vertical
        .SELECAO = Condicao
        .Arvore = False
        .Visualizar
        
    End With
    Screen.MousePointer = 0
    Set Rpt = Nothing
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIM, txtRazao
End Sub

Private Sub cmdPesqIc_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtIC
End Sub

Private Sub Form_Activate()
    txtIM.SetFocus
End Sub

Private Sub Form_Load()
    CboImposto.Preencher Bdados, "SELECT TIP_COD_IMPOSTO,TIP_sigla_IMPOSTO " & _
                " FROM TAB_IMPOSTO ORDER BY TIP_sigla_IMPOSTO  ASC", 1
    cabVisual.Exibir Bdados, Me.Name, App.Path
    CboImposto.AddItem " "
    txtRazao.Locked = True
    AtualizaCabecalho lstIptu
    AtualizaCabecalho lstPago
    AtualizaCabecalho lstTotalAberto
    AtualizaCabecalho lstTotalPago
End Sub

Private Sub lstIptu_DblClick()
'    If lstIptu.SelectedItem Is Nothing Then Exit Sub
'    EncontraItem lstPago, lstIptu.SelectedItem
'    EncontraItem lstTotalPago, lstIptu.SelectedItem.SubItems(2)
'    EncontraItem lstTotalAberto, lstIptu.SelectedItem.SubItems(2)
End Sub

Private Sub EncontraItem(Grid As Object, Codigo As String)
    'Dim Item As ListItem
    Dim Item As Variant
        Set Item = Grid.FindItem(Codigo)
        If Not Item Is Nothing Then
            Item.Selected = True
            Item.EnsureVisible
        End If
        Set Item = Nothing
End Sub

Private Sub lstIptu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 2 Then
'        mnuEstorno.Caption = "Estornar DAM " & lstIptu.SelectedItem
'        mnuReimprime.Caption = "Imprimir  DAM " & lstIptu.SelectedItem
'        Me.PopupMenu OpcoesDAM
'    End If
End Sub

Private Sub lstPago_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If lstPago.ListItems.Count > 0 Then
'        If Button = 2 Then
'            mnuEstornoPago.Caption = "Estornar DAM " & lstPago.SelectedItem
'            mnuReimprimePago.Caption = "Imprimir  DAM " & lstPago.SelectedItem
'            Me.PopupMenu OpcoesPago
'        End If
'    End If
End Sub

Private Sub mnuEstornoPago_Click()
    'TARR301.txtDAM = lstPago.SelectedItem
    'SendKeys "{tab}"
    'TARR301.Show
    'TARR301.txtMotivo.SetFocus
End Sub

Private Sub mnuReimprimePago_Click()
'    TCOB204.txtDAM = lstPago.SelectedItem
'    SendKeys "{tab}"
'    TCOB204.Show
'    TCOB204.txtObservacao.SetFocus
End Sub

Private Sub Timer1_Timer()
    FocalizaCaixa Me
    Timer1.Enabled = False
End Sub

Private Sub txtIC_LostFocus()
    Dim Ic As String
  
    If Trim(txtIC) <> "" Then
        txtIC = BuscaContribuinte(txtIC, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtIC) = "" Then
            Avisa "Inscricão não encontrada"
            txtIM.SetFocus
        End If
    End If
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtIm_LostFocus()
    If Trim(txtIM) <> "" Then
        txtIM = BuscaContribuinte(txtIM, txtRazao, txtEndereco, , etiContribuinte)
        If Trim(txtIM) = "" Then
            Avisa "Inscricão não encontrada"
            txtIM.SetFocus
        End If
    End If
End Sub

Private Sub txtPeriodo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtPeriodo1_LostFocus()
    If IsNumeric(txtPeriodo1) Then
        If Len(txtPeriodo1) = 6 Then
            txtPeriodo1.MaxLength = 7
            txtPeriodo1 = Left(txtPeriodo1, 2) & "/" & Right(txtPeriodo1, 4)
        End If
    End If
End Sub

Private Sub txtPeriodo2_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, numero)
End Sub

Private Sub txtPeriodo2_LostFocus()
    If IsNumeric(txtPeriodo2) Then
        If Len(txtPeriodo2) = 6 Then
            txtPeriodo2.MaxLength = 7
            txtPeriodo2 = Left(txtPeriodo2, 2) & "/" & Right(txtPeriodo2, 4)
        End If
    End If
End Sub

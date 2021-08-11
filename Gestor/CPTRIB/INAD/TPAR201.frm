VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TPAR201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR201.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.TextBox txtMotivo 
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
      Height          =   435
      Left            =   1980
      TabIndex        =   3
      Tag             =   "Motivo"
      Top             =   4560
      Width           =   7740
   End
   Begin Threed.SSFrame fra 
      Height          =   675
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1191
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
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtNumParc 
         Alignment       =   1  'Right Justify
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
         Left            =   1830
         TabIndex        =   0
         Tag             =   "Exercicio"
         Top             =   240
         Width           =   2025
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   13
         Left            =   150
         TabIndex        =   15
         Top             =   262
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   476
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
         Caption         =   "Nº do Parcelamento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   690
      Index           =   1
      Left            =   45
      TabIndex        =   8
      Top             =   2385
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   1217
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
      Caption         =   "Detalhes do Parcelamento"
      ShadowStyle     =   1
      Begin VB.TextBox txtValorParcelamento 
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
         Left            =   5220
         TabIndex        =   6
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox txtCotas 
         Alignment       =   1  'Right Justify
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
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   5
         Top             =   270
         Width           =   855
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   1275
         TabIndex        =   9
         Top             =   315
         Width           =   495
         _ExtentX        =   873
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
         Caption         =   "Cotas"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   9
         Left            =   3210
         TabIndex        =   10
         Top             =   315
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Valor Parcelamento(R$)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   11
      Top             =   -270
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   975
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   1380
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1720
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
      Caption         =   "Contribuinte"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtDescTrib 
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
         Left            =   3285
         TabIndex        =   21
         Top             =   555
         Width           =   6330
      End
      Begin VB.TextBox txtImposto 
         Alignment       =   1  'Right Justify
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
         Left            =   1830
         TabIndex        =   17
         Top             =   555
         Width           =   1440
      End
      Begin VB.TextBox txtIm 
         Alignment       =   1  'Right Justify
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
         Left            =   1830
         TabIndex        =   4
         Top             =   210
         Width           =   1440
      End
      Begin VB.TextBox txtContrib 
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
         Left            =   3285
         TabIndex        =   13
         Top             =   210
         Width           =   6330
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   15
         Left            =   990
         TabIndex        =   14
         Top             =   255
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "Inscrição"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   270
         TabIndex        =   18
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
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
         Caption         =   "Código de Receita"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin MSComctlLib.ListView lstParc 
      Height          =   1410
      Left            =   30
      TabIndex        =   16
      Top             =   3120
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   2487
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2540
      EndProperty
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   2
      Left            =   135
      TabIndex        =   19
      Top             =   4560
      Width           =   1800
      _ExtentX        =   3175
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
      Caption         =   "Motivo Cancelamento"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   4
      RoundedCorners  =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TPAR201.frx":2123
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8595
      TabIndex        =   2
      Top             =   5055
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdParcela 
      Height          =   375
      Left            =   6255
      TabIndex        =   1
      Top             =   5055
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "&Excluir Parcelamento"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TPAR201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim MaxCotas As Byte
Dim CodImp As String
Dim ContaParcelada As Double
Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstParc.ListItems.Clear
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdParcela_Click()
    Dim Sql As String
    Dim CCorrente As New ContaCorrente
    Dim rs As VSRecordset
    Dim campos As String
    Dim Valores As String
    Dim Conta As New ContaCorrente
    Dim Obrig As New ContaCorrente
    Dim Obriga As New Obrigacao
    Dim i As Integer
    On Error Resume Next
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    If Confirma("Confirma a exclusão do Parcelamento nº " & txtNumParc & ",  no valor total de R$ " & txtValorParcelamento & "?") Then
        Screen.MousePointer = 11
        campos = "TPA_DATA_CANCELAMENTO,TPA_MOTIVO_CANCELAMENTO,TPA_TUS_COD_USUARIO_CANCEL, TPA_STATUS_PARCELAMENTO"
        Valores = Bdados.PreparaValor(Bdados.Converte(Date, TCDataHora), txtMotivo, Aplicacoes.Usuario, stsParcelamentoCancelado)
        Bdados.AtualizaDados "Tab_Parcelamento", Valores, campos, "TPA_NUM_PARCELAMENTO=" & txtNumParc
        '*****
        'troca status das contas parceladas
        Sql = "Select TOC_COD_OBRIGACAO,TOC_STATUS_OBRIGACAO,TOC_STATUS_ANTERIOR_OBRIGACAO,TOC_VALOR_OBRIGACAO_ORIGINAL from tab_obrigacao_contribuinte where " & _
            "TOC_TPA_COD_PARCELAMENTO = " & txtNumParc
        If Bdados.AbreTabela(Sql, rs) Then
            rs.MoveFirst
            Do
                Obriga.TrocaSitObrigacao rs!TOC_COD_OBRIGACAO, Nvl("" & rs!TOC_STATUS_ANTERIOR_OBRIGACAO, -1)
                Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(0, rs.Fields("TOC_VALOR_OBRIGACAO_ORIGINAL")), "TOC_TPA_COD_PARCELAMENTO,TOC_VALOR_OBrIGACAO", "TOC_COD_OBRIGACAO=" & rs!TOC_COD_OBRIGACAO
                rs.MoveNext
            Loop While Not rs.EOF
        End If
        
        Bdados.AtualizaDados "TAB_CONTA_CONTRIBUINTE", Bdados.PreparaValor(0), "TCC_TPA_COD_PARCELAMENTO", "TCC_TPA_COD_PARCELAMENTO = " & txtNumParc
        Bdados.DeletaDados "TAB_COTAS_PARCELAMENTO", "TCP_TPA_COD_PARCELAMENTO = " & txtNumParc
        'APAGO OS PARCELAMENTOS DE ACORDO COM O NOVO PROCESSO DO PARCELAMENTO
        For i = 1 To lstParc.ListItems.Count
            Bdados.DeletaDados "TAB_OBRIGACAO_CONTRIBUINTE", "TOC_COD_OBRIGACAO = " & lstParc.ListItems(i)
        Next
        Avisa "Parcelamento Cancelado."
        LimpaCampos Me
        lstParc.ListItems.Clear
        txtNumParc.SetFocus
        Bdados.FechaTabela rs
    End If
    Bdados.FechaTabela rs
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Trim(Me.Tag) <> "" Then
        txtNumParc = Me.Tag
        txtNumParc_LostFocus
        cmdParcela.SetFocus
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    
End Sub


Private Sub lstParc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstParc, ColumnHeader
End Sub

Private Sub txtCotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCotas_LostFocus()
    If Trim(txtCotas) = "" Then Exit Sub
    If CDbl(Trim(txtCotas)) > MaxCotas Then
        Avisa "Limite máximo de cotas é igual a " & MaxCotas & " ."
        txtCotas.SetFocus
    End If
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtNumParc_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtNumParc_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtNumParc) = "" Then Exit Sub
    Sql = "Select TPA_INSCRICAO,TPA_TIP_COD_IMPOSTO,TPA_NUM_COTAS,TPA_VALOR_PARCELADO,TPA_PERIODO," & _
        "TIP_NOME_IMPOSTO FROM TAB_PARCELAMENTO,TAB_IMPOSTO where TPA_TIP_COD_IMPOSTO=TIP_COD_IMPOSTO " & _
        "AND TPA_NUM_PARCELAMENTO=" & txtNumParc & " AND TPA_STATUS_PARCELAMENTO <>  8 "
    If Bdados.AbreTabela(Sql, rs) Then
        Screen.MousePointer = 11
        txtIm = "" & rs!tpa_inscricao
        txtImposto = "" & rs!tpa_tip_cod_imposto
        txtDescTrib = "" & rs!TIP_NOME_IMPOSTO
        txtCotas = "" & rs!TPA_NUM_COTAS
        txtValorParcelamento = "" & Format(rs!TPA_VALOR_PARCELADO, Const_Monetario)
        
            Sql = " Select TCp_NUM_COTA AS Documento,tcp_inscricao as Inscrição,"
            Sql = Sql & " TIP_SIGLA_IMPOSTO AS Tributo,"
            Sql = Sql & " TPA_PERIODO AS Periodo,"
            Sql = Sql & " TCp_DATA_VENCIMENTO AS Vencimento,TCp_NUM_PARCELA as Cota,"
            Sql = Sql & " TCp_VALOR_PARCELA As Valor, TCp_VALOR_JUROS As Juros, TCp_VALOR_PARCELA"
            Sql = Sql & " + TCp_VALOR_JUROS as Total,tip_cod_imposto as Imposto,"
            Sql = Sql & " tip_nome_imposto as Descrição"
            Sql = Sql & " From tab_parcelamento, tab_cotas_parcelamento, tab_imposto"
            Sql = Sql & " where  tpa_num_parcelamento = TCp_TPA_COD_PARCELAMENTO and"
            Sql = Sql & " TPA_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
            Sql = Sql & " AND TCP_TPA_COD_PARCELAMENTO =   '" & Me.Tag & "'"

        MontaGrid Bdados, lstParc, Sql, 1200, 700, 1200, 1200, 1000, 700
        
        Screen.MousePointer = 0
    Else
        Avisa "Parcelamento inexistente."
        txtNumParc = ""
        lstParc.ListItems.Clear
        txtNumParc.SetFocus
    End If
    Bdados.FechaTabela rs
End Sub

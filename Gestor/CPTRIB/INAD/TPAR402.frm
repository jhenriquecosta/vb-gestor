VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   23
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR402.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   675
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   10905
      _ExtentX        =   19235
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
         TabIndex        =   13
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
      Left            =   30
      TabIndex        =   7
      Top             =   2415
      Width           =   10935
      _ExtentX        =   19288
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
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtPeriodoFinal 
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
         Left            =   9780
         MaxLength       =   14
         TabIndex        =   20
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtPeriodoIncial 
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
         Left            =   7560
         MaxLength       =   14
         TabIndex        =   5
         Top             =   270
         Width           =   975
      End
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
         Left            =   4950
         TabIndex        =   4
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
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   3
         Top             =   270
         Width           =   855
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   5
         Left            =   1425
         TabIndex        =   8
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
         Index           =   6
         Left            =   6630
         TabIndex        =   9
         Top             =   315
         Width           =   855
         _ExtentX        =   1508
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
         Caption         =   "Per. Inicial"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   9
         Left            =   2940
         TabIndex        =   10
         Top             =   315
         Width           =   1950
         _ExtentX        =   3440
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
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   8850
         TabIndex        =   21
         Top             =   300
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
         Caption         =   "Per. Final"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
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
      Height          =   1005
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   1380
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1773
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
      Begin VB.TextBox txtNomeImposto 
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
         Left            =   3060
         TabIndex        =   19
         Top             =   600
         Width           =   6405
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
         TabIndex        =   16
         Top             =   600
         Width           =   1185
      End
      Begin VB.TextBox txtIc 
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
         TabIndex        =   14
         Top             =   240
         Width           =   2025
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   990
         TabIndex        =   15
         Top             =   285
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
         TabIndex        =   17
         Top             =   645
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TPAR402.frx":2123
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   5475
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Sair"
      Acao            =   7
   End
   Begin VTOcx.cmdVISUAL cmdParcela 
      Height          =   375
      Left            =   8310
      TabIndex        =   1
      Top             =   5475
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
   End
   Begin VTOcx.grdVISUAL grdCotas 
      Height          =   2295
      Left            =   30
      TabIndex        =   22
      Top             =   3150
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   4048
      Caption         =   "Parcelamentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
End
Attribute VB_Name = "TPAR402"
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
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdParcela_Click()
    Dim Sql As String
    Dim CCorrente As New ContaCorrente
    Dim rs As VSRecordset
    Dim Campos As String
    Dim Valores As String
    
    On Error Resume Next
    If Not Edita.CriticaCampos(Me) Then Exit Sub

     With Rpt
            If txtImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
                If Not .DefinirArquivo(Bdados, App.Path + "\TermoObrigacao.rpt") Then Exit Sub
            Else
                If Not .DefinirArquivo(Bdados, App.Path + "\TermoParcela.rpt") Then Exit Sub
            End If
            'If Not .DefinirArquivo(Bdados, App.Path + "\TermoParcela.rpt") Then Exit Sub
            .Formulas "NumParcelamento ", CStr(txtNumParc)
            .Formulas "Municipio ", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            .Formulas "Imposto ", CStr(txtNomeImposto)
            .Formulas "Inscricao", Trim(txtIc)
            If IsNumeric(txtValorParcelamento) Then
                .Formulas "ValorExtenso", "(" + VBA.UCase(Extenso(CDbl(txtValorParcelamento), "Reais", "Real")) + ")"
            End If
            .Formulas "VT_Periodo ", IIf(Len(txtPeriodoIncial) = 4, txtPeriodoIncial, Right(txtPeriodoIncial, 2) & "/" & Left(txtPeriodoIncial, 4)) & " a " & IIf(Len(txtPeriodoFinal) = 4, txtPeriodoFinal, Right(txtPeriodoFinal, 2) & "/" & Left(txtPeriodoFinal, 4))
            .Selecao = "{Tab_Parcelamento.TPA_NUM_PARCELAMENTO} = " & txtNumParc
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Titulo = "Termo de Parcelamento"
            .Arvore = False
            .Visualizar
    End With
    Set Rpt = Nothing
End Sub

Private Sub cmdsair_Click()
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
    On Error GoTo TrataErro
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtNumParc) = "" Then Exit Sub
    If IsNumeric(txtNumParc) = False Then Exit Sub
    Sql = "Select TPA_INSCRICAO,TPA_TCI_IM,TPA_TIM_IC,TPA_TIP_COD_IMPOSTO,TPA_NUM_COTAS,TPA_VALOR_PARCELADO,TPA_PERIODO_INICIAL,TPA_PERIODO_FINAL FROM TAB_PARCELAMENTO where TPA_NUM_PARCELAMENTO=" & txtNumParc
    If Bdados.AbreTabela(Sql, rs) Then
        Screen.MousePointer = 11
        txtIc = "" & rs!TPA_INSCRICAO
        txtImposto = "" & rs!TPA_TIP_COD_IMPOSTO
        txtCotas = "" & rs!TPA_NUM_COTAS
        txtValorParcelamento = "" & Format(rs!TPA_VALOR_PARCELADO, Const_Monetario)
        txtPeriodoIncial = "" & rs!TPA_PERIODO_INICIAL
        txtPeriodoFinal = "" & rs!TPA_PERIODO_FINAL
        txtNomeImposto = Bdados.BuscaCodigo("Select tip_nome_imposto from tab_imposto where tip_cod_imposto = '" & txtImposto & "'")
        If txtImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) Then
            Sql = " Select TCO_COD_OBRIGACAo_PARCELA AS Documento,tco_inscricao as Inscrição,"
            Sql = Sql & " TIP_SIGLA_IMPOSTO AS Tributo,tco_periodo as Ano,"
            Sql = Sql & " TCO_DATA_VENCIMENTO AS Vencimento,TCO_NUM_PARCELA as Cota,"
            Sql = Sql & " TCO_VALOR_PARCELA as Valor,TCO_VALOR_JUROS as Juros ,"
            Sql = Sql & " TCO_VALOR_PARCELA + TCO_VALOR_JUROS as Total,"
            Sql = Sql & " tip_cod_imposto as Imposto,tip_nome_imposto as Descrição,Tge_Nome as Situação"
            Sql = Sql & " From  TAB_COTAS_OBRIGACAO,"
            Sql = Sql & " tab_imposto , vis_status_obrigacao"
            Sql = Sql & " Where"
            Sql = Sql & " TCO_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
            Sql = Sql & " AND TCO_TPA_COD_PARCELAMENTO =  '" & Me.Tag & "'"
            Sql = Sql & " and tge_codigo = TCO_STATUS_OBRIGACAO_PARCELA"
        Else
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
        End If
        grdCotas.Preencher Bdados, Sql, 0, 1200, 1200, 800, 1200, 600, 900, 800, 900, 0
        If grdCotas.ListItems.Count > 0 Then grdCotas.Mensagem = "Total Parcelamento: R$" & Format(grdCotas.Colunas(7).Soma, Const_Monetario) & " x Acréscimo na dívida original: R$" & Format(grdCotas.Colunas(9).Soma - grdCotas.Colunas(7).Soma, Const_Monetario)
        Screen.MousePointer = 0
    Else
        Avisa "Parcelamento inexistente."
        txtNumParc.SetFocus
    End If
    Bdados.FechaTabela rs
    
    Exit Sub
TrataErro:
    If Err.Number = -2147217900 Then 'Incorrect syntax near '4'.
        txtNumParc.Text = ""
        Screen.MousePointer = 0
        Exit Sub
    Else
        Util.Erro Err.Description
        Screen.MousePointer = 0
    End If
End Sub

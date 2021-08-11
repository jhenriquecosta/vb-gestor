VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   1138
      Icone           =   "TMPU101.frx":0000
   End
   Begin MSComctlLib.ListView lstBvt 
      Height          =   3780
      Left            =   60
      TabIndex        =   9
      Top             =   1620
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   6668
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Bairro"
         Text            =   "Bairro"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Logr"
         Text            =   "Logr"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Nome Logr"
         Text            =   "Nome Logr"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Lado"
         Text            =   "Lado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Valor"
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin Threed.SSFrame fra 
      Height          =   885
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   1561
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
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   315
         Left            =   3105
         TabIndex        =   13
         Top             =   375
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboBairro 
         Height          =   315
         Left            =   1110
         TabIndex        =   12
         Top             =   375
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   12582912
         CorTexto        =   12582912
      End
      Begin VB.TextBox txtCod 
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
         Left            =   135
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Valor"
         Top             =   375
         Width           =   945
      End
      Begin VB.TextBox txtValor 
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
         Left            =   9285
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Valor"
         Top             =   330
         Width           =   915
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   1125
         TabIndex        =   5
         Top             =   180
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   476
         _Version        =   196610
         CaptionStyle    =   1
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
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   9345
         TabIndex        =   6
         Top             =   75
         Width           =   750
         _ExtentX        =   1323
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
         Caption         =   "Valor(R$)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   7485
         TabIndex        =   7
         Top             =   135
         Width           =   405
         _ExtentX        =   714
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
         Caption         =   "Lado"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   196610
         CaptionStyle    =   1
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
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   10
         Top             =   165
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
         Caption         =   "Codigo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   315
         Left            =   4920
         TabIndex        =   14
         Top             =   360
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboLado 
         Height          =   315
         Left            =   7455
         TabIndex        =   15
         Top             =   345
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   9135
      TabIndex        =   3
      Top             =   5475
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   7935
      TabIndex        =   2
      Top             =   5475
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   5955
      TabIndex        =   11
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "TMPU101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto

Private Sub AtualizaLista()
    Dim Sql As String
    Dim rs As VSRecordset
    If Bdados.Conexao.FormatoBanco = SQLServer Then
    Sql = "SELECT TAB_LOGRADOURO.TLG_COD_LOGRADOURO as COD, TAB_BAIRRO.TBA_NOME as Bairro, TAB_TIPO_LOGR.TTL_NOME as Logr, " & _
        " TAB_LOGRADOURO.tlg_nome as Nome, Tab_Geral.TGE_NOME as Lado, " & _
        " Tab_Valor_Terreno.tvl_valor as [Valor(R$)] FROM TAB_BAIRRO,TAB_TIPO_LOGR," & _
        " TAB_LOGRADOURO,Tab_Geral,Tab_Valor_Terreno WHERE " & _
        " Tab_Valor_Terreno.tvl_tlg_cod_logradouro = TAB_LOGRADOURO.tlg_cod_logradouro AND " & _
        " TAB_TIPO_LOGR.TTL_COD_TIP_LOGR = TAB_LOGRADOURO.tlg_ttl_cod_tip_logr AND " & _
        " TAB_BAIRRO.TBA_COD_BAIRRO = TAB_LOGRADOURO.tlg_tba_cod_bairro " & _
        " AND Tab_Valor_Terreno.tvl_lado =Tab_Geral.TGE_CODIGO " & _
        " AND Tab_Geral.TGE_CODIGO >0 and Tab_Geral.TGE_TIPO =8 and  tlg_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO") & " and tba_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO") & _
        " ORDER BY TAB_BAIRRO.TBA_NOME, TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "SELECT TAB_LOGRADOURO.TLG_COD_LOGRADOURO as COD, TAB_BAIRRO.TBA_NOME as Bairro, TAB_TIPO_LOGR.TTL_NOME as Logr, " & _
        " TAB_LOGRADOURO.tlg_nome as Nome, Tab_Geral.TGE_NOME as Lado, " & _
        " Tab_Valor_Terreno.tvl_valor as Valor FROM TAB_BAIRRO,TAB_TIPO_LOGR," & _
        " TAB_LOGRADOURO,Tab_Geral,Tab_Valor_Terreno WHERE " & _
        " Tab_Valor_Terreno.tvl_tlg_cod_logradouro = TAB_LOGRADOURO.tlg_cod_logradouro AND " & _
        " TAB_TIPO_LOGR.TTL_COD_TIP_LOGR = TAB_LOGRADOURO.tlg_ttl_cod_tip_logr AND " & _
        " TAB_BAIRRO.TBA_COD_BAIRRO = TAB_LOGRADOURO.tlg_tba_cod_bairro " & _
        " AND Tab_Valor_Terreno.tvl_lado =Tab_Geral.TGE_CODIGO " & _
        " AND Tab_Geral.TGE_CODIGO >0 and Tab_Geral.TGE_TIPO =8 and  tlg_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO") & " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & _
        " ORDER BY TAB_BAIRRO.TBA_NOME, TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome "
    End If
    Call MontaGrid(Bdados, lstBvt, Sql, 1400)
End Sub
Private Sub cboBairroA_Click()
    Call cadastro.BuscaLogradouro(TipoLogr, cboBairro, cboTipoLogr)
End Sub

Private Sub cboTipoLogra_Click()
    Call cadastro.BuscaLogradouro(Rua, cboTipoLogr, cboLogr)
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim CodLogradouro As Long
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    CodLogradouro = cadastro.PegaCodLogr(cboBairro.Text, cboTipoLogr.Text, cboLogr.Text)
    Campos = "tvl_tlg_cod_logradouro,tvl_lado,tvl_valor,tvl_num_trecho,tvl_tba_cod_bairro"
    Valores = Bdados.PreparaValor(CodLogradouro, cboLado.ListIndex + 1, Bdados.Converte(txtValor, TCDuplo), 1, cboBairro.ListIndex + 1)
    If cboLado.ListIndex + 1 = 3 Then
        Bdados.DeletaDados "Tab_Valor_Terreno", "tvl_tlg_cod_logradouro='" & CodLogradouro & "'"
    Else
        Bdados.DeletaDados "Tab_Valor_Terreno", "tvl_tlg_cod_logradouro='" & CodLogradouro & "' and tvl_lado=3"
    End If
    Call Bdados.GravaDados("Tab_Valor_Terreno", Valores, Campos, "tvl_tlg_cod_logradouro='" & CodLogradouro & "' and tvl_lado=" & cboLado.ListIndex + 1)
    AtualizaLista
    'Informa "Transação completada."
    Edita.LimpaCampos Me
    txtCod.SetFocus
End Sub

Private Sub Form_Activate()
    cboLogr.Preencher Bdados, "Select tlg_cod_logradouro, tlg_nome From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio, 1
    cboTipoLogr.Preencher Bdados, "Select TTL_COD_TIP_LOGR, ttl_nome From Tab_Tipo_Logr", 1
    cboBairro.Preencher Bdados, "Select TBA_COD_BAIRRO, tba_nome From Tab_Bairro where tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio, 1
    cboLado.Preencher Bdados, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =8 ORDER BY TGE_CODIGO ASC"
    Set cadastro = New VSImposto
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    AtualizaLista
End Sub

Private Sub lstBvt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstBvt, ColumnHeader
End Sub

Private Sub txtCod_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCod) = "" Then Exit Sub
    Sql = "Select * from vis_bvt where tlg_cod_logradouro ='" & txtCod & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        cboBairro.SetarLinha rs!TLG_TBA_COD_BAIRRO, 0
        cboTipoLogr.SetarLinha rs!ttl_nome, 1
        cboLogr.SetarLinha rs!tlg_nome, 1
        cboLado.ListIndex = 2
        txtValor.SetFocus
    Else
        Avisa "Logradouro inexistente."
        txtCod.SetFocus
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValor_LostFocus()
    txtValor = Edita.FormataTexto(txtValor, Monetario, True)
End Sub


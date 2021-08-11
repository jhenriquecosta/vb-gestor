VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU601 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstBvt 
      Height          =   2055
      Left            =   45
      TabIndex        =   12
      Top             =   1620
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   3625
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
      TabIndex        =   7
      Top             =   690
      Width           =   9225
      _ExtentX        =   16272
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
      Begin VB.ComboBox cboBairro 
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
         ItemData        =   "TMPU601.frx":0000
         Left            =   90
         List            =   "TMPU601.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Bairro"
         Top             =   390
         Width           =   1485
      End
      Begin VB.ComboBox cboLado 
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
         ItemData        =   "TMPU601.frx":0004
         Left            =   5910
         List            =   "TMPU601.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Lado"
         Top             =   390
         Width           =   1635
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
         Left            =   7620
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Valor"
         Top             =   390
         Width           =   1515
      End
      Begin VB.ComboBox cboTipoLogr 
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
         ItemData        =   "TMPU601.frx":0032
         Left            =   1740
         List            =   "TMPU601.frx":003F
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Tipo Logradouro"
         Top             =   390
         Width           =   1665
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
         ItemData        =   "TMPU601.frx":0060
         Left            =   3420
         List            =   "TMPU601.frx":006D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Nome Logradouro"
         Top             =   390
         Width           =   2415
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   150
         Width           =   1080
         _ExtentX        =   1905
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
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   1
         Left            =   7620
         TabIndex        =   9
         Top             =   180
         Width           =   900
         _ExtentX        =   1588
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
         Caption         =   "Valor(R$)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   5940
         TabIndex        =   10
         Top             =   150
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   476
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
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   1770
         TabIndex        =   11
         Top             =   150
         Width           =   1080
         _ExtentX        =   1905
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
         Caption         =   "Logradouro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8100
      TabIndex        =   6
      Top             =   3735
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
      Left            =   6915
      TabIndex        =   5
      Top             =   3735
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1138
      Icone           =   "TMPU601.frx":008E
   End
End
Attribute VB_Name = "TMPU601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto

Private Sub AtualizaLista()
    Dim Sql As String
    Dim rs As VSRecordset
        
    Sql = "SELECT TAB_BAIRRO.TBA_NOME as Bairro, TAB_TIPO_LOGR.TTL_NOME as Logr, " & _
        " TAB_LOGRADOURO.tlg_nome as Nome, Tab_Geral.TGE_NOME as Lado, " & _
        " Tab_Valor_Terreno.tvl_valor as [Valor(R$)] FROM TAB_BAIRRO,TAB_TIPO_LOGR," & _
        " TAB_LOGRADOURO,Tab_Geral,Tab_Valor_Terreno WHERE " & _
        " Tab_Valor_Terreno.tvl_tlg_cod_logradouro = TAB_LOGRADOURO.tlg_cod_logradouro AND " & _
        " TAB_TIPO_LOGR.TTL_COD_TIP_LOGR = TAB_LOGRADOURO.tlg_ttl_cod_tip_logr AND " & _
        " TAB_BAIRRO.TBA_COD_BAIRRO = TAB_LOGRADOURO.tlg_tba_cod_bairro " & _
        " AND Tab_Valor_Terreno.tvl_lado =Tab_Geral.TGE_CODIGO " & _
        " AND Tab_Geral.TGE_CODIGO >0 and Tab_Geral.TGE_TIPO =8" & _
        " and tlg_tmu_cod_municipio =" & Temp.PegaParametro(Bdados, "MUNICIPIO") & " and TBA_TMU_COD_MUNICIPIO =" & Temp.PegaParametro(Bdados, "MUNICIPIO") & _
        " ORDER BY TAB_BAIRRO.TBA_NOME, TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome"
    Call MontaGrid(Bdados, lstBvt, Sql, 1400)
End Sub
Private Sub cboBairro_Click()
    Call cadastro.BuscaLogradouro(TipoLogr, cboBairro, cboTipoLogr)
End Sub

Private Sub cboTipoLogr_Click()
    Call cadastro.BuscaLogradouro(Rua, cboTipoLogr, cboLogr)
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
    Campos = "tvl_tlg_cod_logradouro,tvl_lado,tvl_valor"
    Valores = Bdados.PreparaValor(CodLogradouro, cboLado.ListIndex + 1, Bdados.Converte(txtValor, TCDuplo))
    If cboLado.ListIndex + 1 = 3 Then
        Bdados.DeletaDados "Tab_Valor_Terreno", "tvl_tlg_cod_logradouro=" & CodLogradouro
    Else
        Bdados.DeletaDados "Tab_Valor_Terreno", "tvl_tlg_cod_logradouro=" & CodLogradouro & " and tvl_lado=3"
    End If
    Call Bdados.GravaDados("Tab_Valor_Terreno", Valores, Campos, "tvl_tlg_cod_logradouro='" & CodLogradouro & "' and tvl_lado=" & cboLado.ListIndex + 1)
    AtualizaLista
    Informa "Transação completada."
    Edita.LimpaCampos Me
    cboBairro.SetFocus
End Sub

Private Sub Form_Activate()
    Call Edita.AtualizaCombo(Bdados, cboLogr, "Select tlg_nome From Tab_Logradouro where tlg_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO"))
    Call Edita.AtualizaCombo(Bdados, cboTipoLogr, "Select ttl_nome From Tab_Tipo_Logr")
    Call Edita.AtualizaCombo(Bdados, cboBairro, "Select tba_nome From Tab_Bairro where TBA_TMU_COD_MUNICIPIO=" & Temp.PegaParametro(Bdados, "MUNICIPIO"))
    Call Edita.AtualizaCombo(Bdados, cboLado, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =8 ORDER BY TGE_CODIGO ASC")
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

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValor_LostFocus()
    txtValor = Edita.FormataTexto(txtValor, Monetario, True)
End Sub


VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form CTRN101 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTRN101"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView grdConsultas 
      Height          =   2655
      Left            =   3600
      TabIndex        =   23
      Top             =   3840
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ActiveTabs.SSActiveTabs tabConsulta 
      Height          =   3120
      Left            =   3600
      TabIndex        =   18
      Top             =   660
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   5503
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "CTRN101.frx":0000
      Images          =   "CTRN101.frx":0081
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2700
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   4763
         _Version        =   131082
         TabGuid         =   "CTRN101.frx":0BC9
         Begin VB.CheckBox chkLimpar 
            Appearance      =   0  'Flat
            Caption         =   "Limpar antes de transferir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   90
            TabIndex        =   24
            Top             =   2280
            Value           =   1  'Checked
            Width           =   2805
         End
         Begin VTOcx.cmdVISUAL cmdLimpar 
            Height          =   345
            Left            =   3990
            TabIndex        =   9
            Top             =   2310
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Caption         =   "Limpar"
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   4210752
            CorFundo        =   -2147483633
            Icone           =   "CTRN101.frx":0BF1
         End
         Begin VTOcx.cmdVISUAL cmdBuscar 
            Height          =   345
            Left            =   5040
            TabIndex        =   10
            Top             =   2310
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Caption         =   "Buscar"
            Acao            =   5
            Enabled         =   0   'False
            CorBorda        =   8421504
            CorFrente       =   4210752
            CorFundo        =   -2147483633
            Icone           =   "CTRN101.frx":0F0B
         End
         Begin VTOcx.txtVISUAL txtConsulta 
            Height          =   2235
            Left            =   90
            TabIndex        =   7
            Top             =   0
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   3942
            Caption         =   "Consulta"
            Text            =   ""
            AlinhamentoRotulo=   1
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdIncluir 
            Height          =   345
            Left            =   2940
            TabIndex        =   8
            Top             =   2310
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Caption         =   "Incluir"
            Acao            =   8
            CorBorda        =   8421504
            CorFrente       =   4210752
            CorFundo        =   -2147483633
            Icone           =   "CTRN101.frx":1225
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   2700
         Left            =   -99969
         TabIndex        =   20
         Top             =   30
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   4763
         _Version        =   131082
         TabGuid         =   "CTRN101.frx":15BF
         Begin VTOcx.grdVISUAL grdResultado 
            Height          =   2925
            Left            =   30
            TabIndex        =   22
            Top             =   30
            Width           =   6075
            _ExtentX        =   10716
            _ExtentY        =   4339
            CorBorda        =   8421504
            CorFundo        =   -2147483633
            CorTitulo       =   8388608
            CorCaption      =   -2147483633
            CorDica         =   8388608
            OcultarRodape   =   -1  'True
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   17
      Top             =   6540
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   820
      CorFundo        =   14737632
      CorFrente       =   8421504
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8970
         TabIndex        =   14
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   4210752
         CorFrente       =   4210752
         CorFundo        =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   1138
      Icone           =   "CTRN101.frx":15E7
   End
   Begin VTOcx.fraVISUAL fraVISUAL3 
      Height          =   3105
      Left            =   0
      TabIndex        =   21
      Top             =   660
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5477
      Altura          =   4935
      Caption         =   " Conexão"
      CorTexto        =   16777215
      CorFaixa        =   12632064
      CorFundo        =   16774636
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdConectar 
         Height          =   345
         Left            =   840
         TabIndex        =   5
         Top             =   2700
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         Caption         =   "Conectar"
         Acao            =   8
         CorBorda        =   8421504
         CorFrente       =   4210752
         Icone           =   "CTRN101.frx":1901
      End
      Begin VTOcx.txtVISUAL txtServidor 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   750
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         Caption         =   "Servidor"
         Text            =   ""
      End
      Begin VTOcx.cboVISUAL cboTipoBanco 
         Height          =   315
         Left            =   420
         TabIndex        =   0
         Top             =   330
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtUsuario 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   1200
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         Caption         =   "Usuário"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtSenha 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1650
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         Caption         =   "Senha"
         Text            =   ""
         CaracterSenha   =   "*"
      End
      Begin VTOcx.txtVISUAL txtBanco 
         Height          =   315
         Left            =   270
         TabIndex        =   4
         Top             =   2100
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   556
         Caption         =   "Banco"
         Text            =   ""
      End
      Begin VTOcx.cmdVISUAL cmdDesconectar 
         Height          =   345
         Left            =   2070
         TabIndex        =   6
         Top             =   2700
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         Caption         =   "Desconectar"
         Acao            =   9
         Enabled         =   0   'False
         CorBorda        =   8421504
         CorFrente       =   4210752
         Icone           =   "CTRN101.frx":1C1B
      End
   End
   Begin VTOcx.txtVISUAL txtGrupo 
      Height          =   315
      Left            =   180
      TabIndex        =   15
      Top             =   4440
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      Caption         =   "Grupo"
      Text            =   ""
   End
   Begin VTOcx.cboVISUAL cboSistema 
      Height          =   315
      Left            =   30
      TabIndex        =   11
      Top             =   4020
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      Caption         =   "Sistema"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   345
      Left            =   1380
      TabIndex        =   12
      Top             =   5130
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      Caption         =   "Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   4210752
      Icone           =   "CTRN101.frx":1F35
   End
   Begin VTOcx.cmdVISUAL cmdRetirar 
      Height          =   345
      Left            =   2430
      TabIndex        =   13
      Top             =   5130
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
      Caption         =   "Retirar"
      Acao            =   2
      Enabled         =   0   'False
      CorBorda        =   8421504
      CorFrente       =   4210752
      Icone           =   "CTRN101.frx":22CF
   End
End
Attribute VB_Name = "CTRN101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Banco As Object

Private Sub cmdBuscar_Click()
    If Not Banco Is Nothing Then
        grdResultado.Caption = Util.ParseString(txtConsulta, "FROM ", 2)
        grdResultado.Caption = Util.ParseString(grdResultado.Caption, " WHERE ", 1)
        grdResultado.Preencher Banco, txtConsulta
        tabConsulta.Tabs(2).Selected = True
    End If
End Sub

Private Sub cmdConectar_Click()
    ConectarBanco
End Sub

Private Sub cmdDesconectar_Click()
    DesconectarBanco
End Sub

Private Sub cmdIncluir_Click()
    If txtConsulta <> "" Then
        IncluirConsulta txtConsulta
        LimparConsulta
        HabilitarExcluir
        txtConsulta.SetFocus
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimparConsulta
    txtConsulta.SetFocus
End Sub

Private Sub cmdRetirar_Click()
    If Not grdConsultas.SelectedItem Is Nothing Then
        grdConsultas.ListItems.Remove grdConsultas.SelectedItem.Index
    End If
    HabilitarExcluir
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Public Function ProximoGrupo() As Integer
    Dim sql As String, rs As Object
    
    sql = "SELECT MAX(TGT_COD_GRUPO)+1 FROM TAB_GRUPO_TRANSFERENCIA"
    If Bdados.AbreTabela(sql, rs) Then
        ProximoGrupo = Util.Nvl("" & rs(0), 1)
    End If
    Bdados.FechaTabela rs
End Function

Private Sub cmdSalvar_Click()
    Dim CodGrupo As String
    
    CodGrupo = GravarGrupo()
    GravarConsultas CodGrupo
    cboSistema = ""
    txtGrupo = ""
    grdConsultas.ListItems.Clear
    cmdLimpar_Click
End Sub

Private Sub Form_Load()
    cboTipoBanco.PreencherGeral Bdados, "TIPO BANCO"
    cboSistema.Preencher Bdados, "SELECT TSI_COD_SISTEMA FROM TAB_SISTEMA"
    PrepararGridConsultas
End Sub

Public Sub AcessoConexao(Valor As Boolean)
    cboTipoBanco.Enabled = Valor
    txtServidor.Enabled = Valor
    txtUsuario.Enabled = Valor
    txtSenha.Enabled = Valor
    txtBanco.Enabled = Valor
    cmdDesconectar.Enabled = Not Valor
End Sub

Public Sub AcessoConsulta(Valor As Boolean)
    cmdBuscar.Enabled = Valor
End Sub

Private Sub ConectarBanco()
    Screen.MousePointer = vbHourglass
    Set Banco = CreateObject("VSClass.VSDados")
    If Banco.AbreBanco(cboTipoBanco.Coluna(1).Valor - 1, txtServidor, txtUsuario, txtSenha, txtBanco) Then
        AcessoConexao False
        AcessoConsulta True
    Else
        Util.Erro "Não foi possível conectar ao banco de dados. Verifique os parâmetros informados!"
        AcessoConexao True
        AcessoConsulta False
        cboTipoBanco.SetFocus
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub DesconectarBanco()
    AcessoConexao True
    AcessoConsulta False
    grdResultado.Preencher Banco, ""
    Set Banco = Nothing
    cboTipoBanco.SetFocus
End Sub

Private Sub HabilitarExcluir()
    cmdRetirar.Enabled = grdConsultas.ListItems.Count > 0
End Sub

Private Sub PrepararGridConsultas()
    grdConsultas.ColumnHeaders.Clear
    grdConsultas.ColumnHeaders.Add , , "Codigo", 800
    grdConsultas.ColumnHeaders.Add , , "Consulta", 4700
    grdConsultas.ColumnHeaders.Add , , "Limpar", 300
End Sub

Private Sub IncluirConsulta(Consulta As String)
    Dim Item As ListItem
    Set Item = grdConsultas.ListItems.Add(, , grdConsultas.ListItems.Count + 1)
    Item.SubItems(1) = Consulta
    Item.SubItems(2) = CInt(chkLimpar)
End Sub

Private Sub LimparConsulta()
    txtConsulta = ""
    chkLimpar.Value = vbChecked
    grdResultado.Preencher Bdados, ""
End Sub

Private Function GravarGrupo() As String
    Dim Campos As String, Valores As String, CodGrupo As String
    
    Campos = "TGT_COD_GRUPO, TGT_GRUPO, TGT_TSI_COD_SISTEMA"
    CodGrupo = BuscarUltimoGrupo()
    Valores = Bdados.PreparaValor(CodGrupo, txtGrupo, cboSistema)
    Bdados.GravaDados "TAB_GRUPO_TRANSFERENCIA", Valores, Campos, "TGT_GRUPO ='" & txtGrupo & "'"
    GravarGrupo = CodGrupo
End Function

Private Sub GravarConsultas(CodGrupo As String)
    Dim Campos As String, Valores As String
    Dim Item As ListItem
    
    Campos = "TTT_TGT_COD_GRUPO, TTT_COD_TABELA, TTT_TABELA, TTT_LIMPAR_DESTINO"
    For Each Item In grdConsultas.ListItems
        Valores = Bdados.PreparaValor(CodGrupo, Item, Item.SubItems(1), Item.SubItems(2))
        Bdados.InsereDados "TAB_TABELA_TRANSFERENCIA", Valores, Campos
    Next
End Sub

Private Function BuscarUltimoGrupo() As String
    Dim sql As String
    sql = "SELECT MAX(TGT_COD_GRUPO)+1 FROM TAB_GRUPO_TRANSFERENCIA"
    If Bdados.AbreTabela(sql) Then
        BuscarUltimoGrupo = Util.Nvl("" & Bdados.Tabela.Fields(0), 1)
    End If
    Bdados.FechaTabela
End Function

Private Sub txtConsulta_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

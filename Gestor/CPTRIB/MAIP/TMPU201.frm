VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   1395
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   705
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   2461
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
      Begin VB.TextBox txtItem 
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
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Item"
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cboTipologia 
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
         ItemData        =   "TMPU201.frx":0000
         Left            =   1230
         List            =   "TMPU201.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Tipologia"
         Top             =   510
         Width           =   2865
      End
      Begin VB.ComboBox cboEstrutura 
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
         ItemData        =   "TMPU201.frx":0004
         Left            =   5310
         List            =   "TMPU201.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Estrutura"
         Top             =   510
         Width           =   2865
      End
      Begin VB.ComboBox cboPadrao 
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
         ItemData        =   "TMPU201.frx":0008
         Left            =   1230
         List            =   "TMPU201.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Padrao"
         Top             =   900
         Width           =   2865
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
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Valor"
         Top             =   900
         Width           =   1215
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   195
         TabIndex        =   9
         Top             =   930
         Width           =   1035
         _ExtentX        =   1826
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
         Caption         =   "Padrão"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   4710
         TabIndex        =   10
         Top             =   930
         Width           =   570
         _ExtentX        =   1005
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
         Caption         =   "Valor"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   4140
         TabIndex        =   11
         Top             =   570
         Width           =   1110
         _ExtentX        =   1958
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
         Caption         =   "Estrutura"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   570
         Width           =   1110
         _ExtentX        =   1958
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
         Caption         =   "Tipologia"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   4
         Left            =   600
         TabIndex        =   14
         Top             =   150
         Width           =   570
         _ExtentX        =   1005
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
         Caption         =   "Item"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
   End
   Begin MSComctlLib.ListView lstCUB 
      Height          =   2235
      Left            =   75
      TabIndex        =   12
      Top             =   2160
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   3942
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
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
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   15
      Top             =   -570
      Width           =   375
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   7170
      TabIndex        =   6
      Top             =   4455
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
      Left            =   5970
      TabIndex        =   5
      Top             =   4455
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
      TabIndex        =   16
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   1138
      Icone           =   "TMPU201.frx":000C
   End
   Begin Threed.SSCommand cmdImprimir 
      Height          =   435
      Left            =   4020
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3855
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "TMPU201.frx":0326
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TMPU201.frx":0342
      Caption         =   "&Imprimir"
      ButtonStyle     =   3
      PictureAlignment=   6
      BevelWidth      =   1
   End
End
Attribute VB_Name = "TMPU201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto

Sub AtualizaGrid()
    MontaGrid Bdados, lstCUB, "SELECT TCU_COD_ITEM AS ITEM,TCU_TCO_COD_COMPONENTE_TIPOLOGIA " & _
        " AS TIPOLOGIA,TCU_TCO_COD_COMPONENTE_ESTRUTURA AS ESTRUTURA," & _
        " TCU_TCO_COD_COMPONENTE_PADRAO AS PADRAO,TCU_VALOR_UNITARIO AS VALOR " & _
        " FROM TAB_CUB", 800, 1400
End Sub

Private Function PegaCodComponente(Componente As String, Grupo As String) As Integer
    Dim Sql As String
    Dim rs As VSRecordset
        
    Sql = "SELECT tco_cod_componente from Tab_Componente_AVANCADO Where " & _
    " tco_descricao_componente='" & Componente & "' and tco_grupo = (select " & _
    " tgc_cod_grupo from Tab_Grupo_Componente_AVANCADO where tgc_nome='" & Grupo & "')"
    If Bdados.AbreTabela(Sql, rs) Then
        PegaCodComponente = rs(0)
    End If
    Bdados.FechaTabela rs
End Function

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprimir_Click()
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TMPU102.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Titulo = "Componentes do Cadastro Imobiliário"
        .Arvore = False
        .Visualizar
        DoEvents
    End With
    Set Rpt = Nothing
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Tipologia As Integer
    Dim Estrutura As Integer
    Dim Padrao As Integer
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Tipologia = PegaCodComponente(cboTipologia.Text, "TIPOLOGIA")
    Estrutura = PegaCodComponente(cboEstrutura.Text, "ESTRUTURA")
    Padrao = PegaCodComponente(cboPadrao.Text, "PADRAO")
    Campos = "TCU_COD_ITEM,TCU_TCO_COD_COMPONENTE_TIPOLOGIA,TCU_TCO_COD_COMPONENTE_ESTRUTURA,TCU_TCO_COD_COMPONENTE_PADRAO,TCU_VALOR_UNITARIO"
    Valores = Bdados.PreparaValor(txtItem, Tipologia, Estrutura, Padrao, Bdados.Converte(txtValor, TCDuplo))
    If Bdados.GravaDados("TAB_CUB", Valores, Campos, "TCU_COD_ITEM='" & txtItem & "'") Then
        Informa "Transação completada."
        AtualizaGrid
        Edita.LimpaCampos Me
    End If
    txtItem.SetFocus
End Sub

Private Sub Form_Activate()
    Call Edita.AtualizaCombo(Bdados, cboEstrutura, "Select tco_descricao_componente From TAB_COMPONENTE_AVANCADO WHERE tco_grupo = (SELECT tgc_cod_grupo FROM Tab_Grupo_Componente_Avancado where TGC_NOME = 'ESTRUTURA') order by tco_cod_componente")
    Call Edita.AtualizaCombo(Bdados, cboPadrao, "Select tco_descricao_componente From TAB_COMPONENTE_AVANCADO WHERE tco_grupo = (SELECT tgc_cod_grupo FROM Tab_Grupo_Componente_Avancado where TGC_NOME ='PADRAO') order by tco_cod_componente")
    Call Edita.AtualizaCombo(Bdados, cboTipologia, "Select tco_descricao_componente From TAB_COMPONENTE_AVANCADO WHERE tco_grupo = (SELECT tgc_cod_grupo FROM Tab_Grupo_Componente_Avancado where TGC_NOME ='TIPOLOGIA') order by tco_cod_componente")
    AtualizaGrid
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub lstCUB_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    OrdenaGrid lstCUB, ColumnHeader
End Sub

Private Sub lstCUB_DblClick()
    txtItem = lstCUB.SelectedItem
    txtItem_LostFocus
End Sub

Private Sub lstCUB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Confirma("Deseja realmente excluir o item " & lstCUB.SelectedItem & "?") Then
            If Bdados.DeletaDados("TAB_CUB", "TCU_COD_ITEM ='" & lstCUB.SelectedItem & "'") Then
                Avisa "Dados eliminados com sucesso."
                AtualizaGrid
                LimpaCampos Me
                txtItem.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtItem_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "Select * from tab_cub where TCU_COD_ITEM ='" & txtItem & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        cboEstrutura.ListIndex = IIf(IsNull(rs!TCU_TCO_COD_COMPONENTE_ESTRUTURA), -1, rs!TCU_TCO_COD_COMPONENTE_ESTRUTURA - 1)
        cboPadrao.ListIndex = IIf(IsNull(rs!TCU_TCO_COD_COMPONENTE_PADRAO), -1, rs!TCU_TCO_COD_COMPONENTE_PADRAO)
        cboTipologia.ListIndex = rs!TCU_TCO_COD_COMPONENTE_TIPOLOGIA - 1
        txtValor = Format(rs!TCU_VALOR_UNITARIO, Const_Monetario)
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValor_LostFocus()
    txtValor = Edita.FormataTexto(txtValor, Monetario, True)
End Sub


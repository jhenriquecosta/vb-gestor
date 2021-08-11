VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TIMP103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   29
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "Timp103.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin ActiveTabs.SSActiveTabs tabTributo 
      Height          =   4590
      Left            =   120
      TabIndex        =   15
      Top             =   690
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   8096
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      Tabs            =   "Timp103.frx":2123
      Images          =   "Timp103.frx":21A0
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4170
         Index           =   0
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   7355
         _Version        =   131082
         TabGuid         =   "Timp103.frx":2AE8
         Begin VTOcx.txtVISUAL txtCodigoTipo 
            Height          =   315
            Left            =   960
            TabIndex        =   0
            Top             =   60
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            Caption         =   "Tipo"
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtTipo 
            Height          =   315
            Left            =   2040
            TabIndex        =   1
            Top             =   60
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtContaContabil 
            Height          =   315
            Left            =   60
            TabIndex        =   2
            Top             =   420
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            Caption         =   "Conta Contábil"
            Text            =   ""
         End
         Begin VTOcx.cboVISUAL cboReceita 
            Height          =   315
            Left            =   690
            TabIndex        =   3
            Top             =   780
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   556
            Caption         =   "Receita"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.grdVISUAL grdTipos 
            Height          =   2460
            Left            =   30
            TabIndex        =   18
            Top             =   1170
            Width           =   6420
            _ExtentX        =   11324
            _ExtentY        =   4339
            Caption         =   "Tipos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
         Begin VTOcx.cmdVISUAL cmdLimparTipo 
            Height          =   375
            Left            =   3510
            TabIndex        =   19
            Top             =   3690
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            Caption         =   "&Limpar"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdSalvarTipo 
            Height          =   375
            Left            =   4470
            TabIndex        =   4
            Top             =   3690
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            Caption         =   "&Salvar"
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdExcluirTipo 
            Height          =   375
            Left            =   5460
            TabIndex        =   20
            Top             =   3690
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "&Excluir"
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4170
         Index           =   1
         Left            =   -99969
         TabIndex        =   17
         Top             =   30
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   7355
         _Version        =   131082
         TabGuid         =   "Timp103.frx":2B10
         Begin VTOcx.grdVISUAL grdCategorias 
            Height          =   2490
            Left            =   30
            TabIndex        =   21
            Top             =   1170
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   4392
            Caption         =   "Categorias"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
         Begin VTOcx.txtVISUAL txtCodigoCategoria 
            Height          =   315
            Left            =   510
            TabIndex        =   22
            Top             =   60
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            Caption         =   "Categoria"
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCategoria 
            Height          =   315
            Left            =   2040
            TabIndex        =   23
            Top             =   60
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtContaCategoria 
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   420
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            Caption         =   "Conta Contábil"
            Text            =   ""
         End
         Begin VTOcx.cboVISUAL cboTipo 
            Height          =   315
            Left            =   960
            TabIndex        =   25
            Top             =   780
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   556
            Caption         =   "Tipo"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL1 
            Height          =   375
            Left            =   3510
            TabIndex        =   26
            Top             =   3690
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            Caption         =   "&Limpar"
            Acao            =   1
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL2 
            Height          =   375
            Left            =   4470
            TabIndex        =   27
            Top             =   3690
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            Caption         =   "&Salvar"
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdVISUAL3 
            Height          =   375
            Left            =   5460
            TabIndex        =   28
            Top             =   3690
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Caption         =   "&Excluir"
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1455
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   5670
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   2566
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
      Caption         =   "Tipo Categoria"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.cboVISUAL cboTipoTributo 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   556
         Caption         =   "Tipo Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.TextBox txtNomeImposto 
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   6
         Top             =   570
         Width           =   5955
      End
      Begin VB.TextBox txtCodImposto 
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
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   5
         Top             =   210
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   9
         Left            =   180
         TabIndex        =   12
         Top             =   240
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
         Caption         =   "Conta Orçamentária"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   315
         TabIndex        =   13
         Top             =   615
         Width           =   1560
         _ExtentX        =   2752
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
         Caption         =   "Nome da Categoria"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   1138
      Icone           =   "Timp103.frx":2B38
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   6855
      TabIndex        =   10
      Top             =   8460
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   2790
      TabIndex        =   8
      Top             =   7230
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   9
      Top             =   8460
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TIMP103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Tipo As cTipoTributo
Private Categoria As cCategoriaTributo


Dim CodCategoria As Double
Dim CodReceita As Double

Private Sub LimparTipo()
    txtCodigoTipo = ""
    txtTipo = ""
    txtContaContabil = ""
    cboReceita = ""
End Sub
Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    
    Select Case Index
        Case 0
            
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Sql = "Select TIC_COD_CATEGORIA from TAB_IMPOSTO_CATEGORIA where TIC_CONTA = " & CodCategoria
            If Not Bdados.AbreTabela(Sql, rs) Then
                Sql = "Select max(TIC_COD_CATEGORIA) +1 from TAB_IMPOSTO_CATEGORIA"
                If Bdados.AbreTabela(Sql, rs) Then
                    CodCategoria = rs(0)
                Else
                    CodCategoria = 1
                End If
            Else
                CodCategoria = rs!TIC_COD_CATEGORIA
            End If
            Valores = Bdados.PreparaValor(CodCategoria, txtCodImposto, txtNomeImposto, cboTipoTributo.Coluna(0).Valor)
            Campos = "TIC_COD_CATEGORIA,TIC_CONTA,TIC_NOME_CATEGORIA,TIC_TTT_COD_TIPO"
            Call Bdados.GravaDados("TAB_IMPOSTO_CATEGORIA", Valores, Campos, "TIC_COD_CATEGORIA='" & CodCategoria & "'")
            Call Util.Informa("Transação Completada.")
            Edita.LimpaCampos Me
            txtCodImposto.Enabled = True
'            lstImposto.Preencher Bdados, "SELECT TIC_CONTA AS CONTA,TIC_NOME_CATEGORIA AS [CATEGORIA],TTT_GRUPO  AS TIPO FROM TAB_IMPOSTO_CATEGORIA,TAB_TIPO_TRIBUTO WHERE  TTT_COD_TIPO=TIC_TTT_COD_TIPO", 1400
            txtCodImposto.SetFocus
        Case 3
            
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Sql = "Select TTT_COD_TIPO from TAB_TIPO_TRIBUTO where TTT_NUM_CONTA = " & CodReceita
            If Not Bdados.AbreTabela(Sql, rs) Then
                Sql = "Select max(TTT_COD_TIPO) +1 from TAB_TIPO_TRIBUTO"
                If Bdados.AbreTabela(Sql, rs) Then
                    CodReceita = rs(0)
                Else
                    CodReceita = 1
                End If
            Else
                CodReceita = rs!TTT_COD_TIPO
            End If
            'Valores = Bdados.PreparaValor(CodReceita, txtTipoTributo, txtCodConta, cboReceita.Coluna(1).Valor)
            Campos = "TTT_COD_TIPO,TTT_GRUPO,TTT_NUM_CONTA,TTT_RECEITA_TRIBUTARIA"
            Call Bdados.GravaDados("TAB_TIPO_TRIBUTO", Valores, Campos, "TTT_COD_TIPO='" & CodReceita & "'")
            Call Util.Informa("Transação Completada.")
            Edita.LimpaCampos Me
            txtCodImposto.Enabled = True
            grdTipos.Preencher Bdados, "select TTT_NUM_CONTA AS CONTA,TTT_GRUPO AS GRUPO,TGE_NOME AS TIPO FROM TAB_TIPO_TRIBUTO,TAB_GERAL WHERE TTT_RECEITA_TRIBUTARIA = TGE_CODIGO AND TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='TIPO RECEITA')", 1400
'            txtCodConta.SetFocus
        Case 1
            Unload Me
        Case 2
            Edita.LimpaCampos Me
            txtCodImposto.Enabled = True
            txtCodImposto.SetFocus
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdExcluirTipo_Click()
    If Not (grdTipos.SelectedItem Is Nothing) Then
        grdTipos_Click
        Tipo.preencherObjeto txtCodigoTipo, txtTipo, txtContaContabil, cboReceita
        If Tipo.Excluir() Then
            cmdLimparTipo_Click
            Tipo.preencherGrid grdTipos
        End If
    End If
End Sub

Private Sub cmdLimparTipo_Click()
    LimparTipo
    txtCodigoTipo = Tipo.buscarProximo()
    txtTipo.SetFocus
End Sub

Private Sub cmdSalvarTipo_Click()
    Tipo.preencherObjeto txtCodigoTipo, txtTipo, txtContaContabil, cboReceita
    If Tipo.Salvar() Then
        cmdLimparTipo_Click
        Tipo.preencherGrid grdTipos
    End If
End Sub



Private Sub Form_DblClick()
'    Dim Conta As New VsTFuncoes.ContaCorrente
'    Debug.Print Conta.GeraCodPagamento("11210201")
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    'lstImposto.Preencher Bdados, "SELECT TIC_CONTA AS CONTA,TIC_NOME_CATEGORIA AS [CATEGORIA],TTT_GRUPO  AS TIPO FROM TAB_IMPOSTO_CATEGORIA,TAB_TIPO_TRIBUTO WHERE  TTT_COD_TIPO=TIC_TTT_COD_TIPO", 1400
    'grdTipos.Preencher Bdados, "select TTT_NUM_CONTA AS CONTA,TTT_GRUPO AS GRUPO,TGE_NOME AS TIPO FROM TAB_TIPO_TRIBUTO,TAB_GERAL WHERE TTT_RECEITA_TRIBUTARIA = TGE_CODIGO AND TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='TIPO RECEITA')", 1400
    'cboTipoTributo.Preencher Bdados, "select TTT_COD_TIPO, TTT_GRUPO from TAB_TIPO_TRIBUTO", 1
    'cboReceita.PreencherGeral Bdados, "TIPO RECEITA"
    'AtualizaCabecalho lstImposto
    
    Set Tipo = New cTipoTributo
    Set Categoria = New cCategoriaTributo
    
    Tipo.Receita.preencherCombo cboReceita
    Tipo.preencherGrid grdTipos
    Categoria.preencherGrid grdCategorias
    
    txtCodigoTipo = Tipo.buscarProximo()
    txtCodigoCategoria = Categoria.buscarProximo()
End Sub

Private Sub lstImposto_Click()
'    txtCodImposto = lstImposto.SelectedItem
    txtCodImposto_LostFocus
End Sub

Private Sub lstImposto_DblClick()
'    If Confirma("Deseja excluir a categoria " & lstImposto.SelectedItem.SubItems(1) & "?") Then
'        If Bdados.DeletaDados("TAB_IMPOSTO_CATEGORIA", " TIC_CONTA =  '" & lstImposto.SelectedItem & "'") Then
'            Avisa "Categoria excluído com sucesso! "
'            grdTipos.Preencher Bdados, "select TTT_NUM_CONTA AS CONTA,TTT_GRUPO AS GRUPO,TGE_NOME AS TIPO FROM TAB_TIPO_TRIBUTO,TAB_GERAL WHERE TTT_RECEITA_TRIBUTARIA = TGE_CODIGO AND TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='TIPO RECEITA')", 1400
'            txtCodImposto.SetFocus
'            Edita.LimpaCampos Me
'        End If
'        grdTipos.SetFocus
'    End If
End Sub

Private Sub grdTipos_Click()
    If Not (grdTipos.SelectedItem Is Nothing) Then
        With grdTipos.SelectedItem
            txtCodigoTipo = .Text
            txtTipo = .SubItems(1)
            txtContaContabil = .SubItems(2)
            cboReceita = .SubItems(3)
        End With
    End If
End Sub

Private Sub grdTipos_DblClick()
    If Not (grdTipos.SelectedItem Is Nothing) Then
        Categoria.preencherGrid grdCategorias, grdTipos.SelectedItem.SubItems(1)
        tabTributo.Tabs(2).Selected = True
    End If
End Sub


Private Sub txtCodConta_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
'    If Trim(txtCodConta) = "" Then Exit Sub
'    CodReceita = txtCodConta
'    sql = "select TTT_GRUPO ,TTT_RECEITA_TRIBUTARIA FROM TAB_TIPO_TRIBUTO WHERE  TTT_NUM_CONTA =" & txtCodConta
    If Bdados.AbreTabela(Sql, rs) Then
'        txtTipoTributo = "" & rs!TTT_GRUPO
        cboReceita.SetarLinha Nvl("" & rs!TTT_RECEITA_TRIBUTARIA, 0), 1
    End If
End Sub

Private Sub txtCodImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCodImposto_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCodImposto) = "" Then Exit Sub
    CodCategoria = txtCodImposto
    Sql = "SELECT TIC_NOME_CATEGORIA ,TIC_TTT_COD_TIPO  FROM TAB_IMPOSTO_CATEGORIA WHERE  TIC_CONTA =" & txtCodImposto
    If Bdados.AbreTabela(Sql, rs) Then
        txtNomeImposto = "" & rs!TIC_NOME_CATEGORIA
        cboTipoTributo.SetarLinha Nvl("" & rs!TIC_TTT_COD_TIPO, 0), 0
    End If
End Sub

Private Sub txtNomeImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSiglaImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TIMP101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   2535
      Index           =   1
      Left            =   45
      TabIndex        =   11
      Top             =   660
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   4471
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
      Begin VB.TextBox txtSubConta 
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
         Left            =   8700
         TabIndex        =   20
         Tag             =   "Código do Tributo"
         Top             =   2100
         Width           =   2085
      End
      Begin VB.TextBox txtLei 
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
         MaxLength       =   200
         TabIndex        =   7
         Top             =   1710
         Width           =   8865
      End
      Begin VTOcx.cboVISUAL CboNatureza 
         Height          =   315
         Left            =   8670
         TabIndex        =   6
         Top             =   1320
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboCategoria 
         Height          =   315
         Left            =   1035
         TabIndex        =   5
         Top             =   1320
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         Caption         =   "Categoria"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.TextBox txtCorrelativo 
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
         Left            =   4590
         MaxLength       =   2
         TabIndex        =   2
         Top             =   570
         Width           =   645
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
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "Nome do Trubuto"
         Top             =   960
         Width           =   8865
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
         TabIndex        =   0
         Tag             =   "Código do Tributo"
         Top             =   210
         Width           =   1185
      End
      Begin VB.TextBox txtSiglaImposto 
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
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Sigla do Tributo"
         Top             =   570
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
         Index           =   12
         Left            =   510
         TabIndex        =   13
         Top             =   600
         Width           =   1410
         _ExtentX        =   2487
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
         Caption         =   "Sigla do Tributo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   420
         TabIndex        =   14
         Top             =   960
         Width           =   1470
         _ExtentX        =   2593
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
         Caption         =   "Nome do Tributo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   3510
         TabIndex        =   16
         Top             =   600
         Width           =   1080
         _ExtentX        =   1905
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
         Caption         =   "Correlativo:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   7920
         TabIndex        =   18
         Top             =   1380
         Width           =   780
         _ExtentX        =   1376
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
         Caption         =   "Natureza"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   960
         TabIndex        =   19
         Top             =   1770
         Width           =   1230
         _ExtentX        =   2170
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
         Caption         =   "Base Legal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboConvenio 
         Height          =   315
         Left            =   6900
         TabIndex        =   3
         Top             =   510
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         Caption         =   "Convênio Arrecadacão"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   4
         Left            =   6660
         TabIndex        =   21
         Top             =   2130
         Width           =   2070
         _ExtentX        =   3651
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
         Caption         =   "Sub Conta Orçamentária"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   750
      TabIndex        =   15
      Top             =   1680
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstImposto 
      Height          =   3000
      Left            =   45
      TabIndex        =   17
      Top             =   3225
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   5292
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   9810
      TabIndex        =   10
      Top             =   6240
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   8580
      TabIndex        =   9
      Top             =   6240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   7350
      TabIndex        =   8
      Top             =   6240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   1138
      Icone           =   "TIMP101.frx":0000
   End
End
Attribute VB_Name = "TIMP101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    
    Select Case Index
        Case 0
            txtSubConta.Tag = ""
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(txtCodImposto, txtSiglaImposto, txtNomeImposto, Bdados.Converte(txtCorrelativo, tctexto), cboCategoria.Coluna(0).Valor, CboNatureza.Coluna(1).Valor, txtLei, cboConvenio.Text, txtSubConta)
            Campos = "tip_cod_imposto,tip_sigla_imposto,tip_nome_imposto,TIP_COD_CORRELATIVO,TIP_TIC_COD_CATEGORIA,TIP_NATUREZA,tip_lei,TIP_TCB_CONVENIO,TIP_SUB_CONTA"
            
            Call Bdados.GravaDados("Tab_Imposto", Valores, Campos, "tip_cod_imposto='" & txtCodImposto & "'")
            Call Util.Informa("Transação Completada.")
            Edita.LimpaCampos Me
            txtCodImposto.Enabled = True
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                lstImposto.Preencher Bdados, "Select tip_cod_imposto as [Código Receita] , tip_sigla_Imposto as [Sigla],tip_nome_imposto as Tributo, TIP_COD_CORRELATIVO as Correlativo,TIC_NOME_CATEGORIA AS Categoria FROM Tab_Imposto LEFT OUTER JOIN TAB_IMPOSTO_CATEGORIA ON Tab_Imposto.TIP_TIC_COD_CATEGORIA = TAB_IMPOSTO_CATEGORIA.TIC_COD_CATEGORIA", 1400
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                lstImposto.Preencher Bdados, "Select tip_cod_imposto as Código_Receita , tip_sigla_Imposto as Sigla,tip_nome_imposto as Tributo, TIP_COD_CORRELATIVO as Correlativo,TIC_NOME_CATEGORIA AS Categoria FROM Tab_Imposto LEFT OUTER JOIN TAB_IMPOSTO_CATEGORIA ON Tab_Imposto.TIP_TIC_COD_CATEGORIA = TAB_IMPOSTO_CATEGORIA.TIC_COD_CATEGORIA", 1400
            End If
            txtCodImposto.SetFocus
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

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        lstImposto.Preencher Bdados, "Select tip_cod_imposto as [Código Receita] , tip_sigla_Imposto as [Sigla],tip_nome_imposto as Tributo, TIP_COD_CORRELATIVO as Correlativo,TIC_NOME_CATEGORIA AS Categoria ,tip_sub_conta as [Sub Conta] FROM Tab_Imposto LEFT OUTER JOIN TAB_IMPOSTO_CATEGORIA ON Tab_Imposto.TIP_TIC_COD_CATEGORIA = TAB_IMPOSTO_CATEGORIA.TIC_COD_CATEGORIA", 1400
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        lstImposto.Preencher Bdados, "Select tip_cod_imposto as Código_Receita , tip_sigla_Imposto as Sigla,tip_nome_imposto as Tributo, TIP_COD_CORRELATIVO as Correlativo,TIC_NOME_CATEGORIA AS Categoria ,tip_sub_conta as Sub_Conta FROM Tab_Imposto LEFT OUTER JOIN TAB_IMPOSTO_CATEGORIA ON Tab_Imposto.TIP_TIC_COD_CATEGORIA = TAB_IMPOSTO_CATEGORIA.TIC_COD_CATEGORIA", 1400
    End If
    cboCategoria.Preencher Bdados, "select TIC_COD_CATEGORIA, TIC_NOME_CATEGORIA from TAB_IMPOSTO_CATEGORIA", 1
    cboConvenio.Preencher Bdados, "SELECT DISTINCT tcb_convenio FROM TAB_CONTA_BANCARIA"
    AtualizaCabecalho lstImposto
    CboNatureza.PreencherGeral Bdados, "NATUREZA"
End Sub

Private Sub lstImposto_Click()
    txtCodImposto = lstImposto.SelectedItem
    txtCodImposto_LostFocus
    Dim Sql As String
        Sql = "SELECT TIP_NATUREZA,tip_lei FROM TAB_IMPOSTO WHERE tip_cod_imposto= '" & lstImposto.SelectedItem & "'"
        If Bdados.AbreTabela(Sql) Then
            If Not IsNull(Bdados.Tabela(0)) Then
                CboNatureza.SetarLinha Bdados.Tabela(0), 1
            Else
                CboNatureza.ListIndex = -1
            End If
            
            If Not IsNull(Bdados.Tabela(1)) Then
                txtLei = Bdados.Tabela(1)
            Else
                txtLei = ""
            End If
            
        End If
End Sub

Private Sub lstImposto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstImposto, ColumnHeader
End Sub

Private Sub lstImposto_DblClick()
    If Confirma("Deseja excluir o tributo " & lstImposto.SelectedItem.SubItems(1) & "?") Then
        Bdados.DeletaDados "Tab_Parametro_Imposto", " tpi_tip_cod_imposto =  '" & lstImposto.SelectedItem & "'"
        Bdados.DeletaDados "Tab_Imposto", "tip_cod_imposto =  '" & lstImposto.SelectedItem & "'"
        Avisa "Tributo excluído com sucesso! "
        lstImposto.Preencher Bdados, "Select tip_cod_imposto as [Código Receita] , tip_sigla_Imposto as [Sigla],tip_nome_imposto as Tributo, TIP_COD_CORRELATIVO as Correlativo,TIC_NOME_CATEGORIA AS Categoria from tab_imposto, TAB_IMPOSTO_CATEGORIA WHERE TIP_TIC_COD_CATEGORIA = TIC_COD_CATEGORIA", 1400
        txtCodImposto.Enabled = True
        Edita.LimpaCampos Me
        txtCodImposto.SetFocus
        'pEGO A NATUREZA
        
    End If
End Sub

Private Sub txtCodImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodImposto_LostFocus()
    On Error Resume Next
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "Select tip_sigla_imposto,tip_nome_imposto,tip_cod_correlativo,TIP_TIC_COD_CATEGORIA,TIP_TCB_CONVENIO,TIP_SUB_CONTA from Tab_Imposto Where tip_cod_imposto='" & txtCodImposto & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtSiglaImposto = rs(0)
        txtNomeImposto = rs(1)
        cboConvenio = "" & rs!TIP_TCB_CONVENIO
        txtCorrelativo = Trim("" & rs!tip_cod_correlativo)
        txtSubConta = "" & rs.Fields("TIP_SUB_CONTA")
        If Not IsNull(rs!TIP_TIC_COD_CATEGORIA) Then
            cboCategoria.SetarLinha rs!TIP_TIC_COD_CATEGORIA, 0
        Else
            cboCategoria.ListIndex = -1
        End If
        txtCodImposto.Enabled = False
    Else
        txtCodImposto.Enabled = True
        txtSiglaImposto = ""
        txtNomeImposto = ""
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtNomeImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSiglaImposto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "TMPU103.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1290
      Left            =   45
      TabIndex        =   8
      Top             =   690
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   2275
      Altura          =   1905
      Caption         =   " "
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   4290
         TabIndex        =   12
         Top             =   870
         Width           =   630
         _ExtentX        =   1111
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
         Caption         =   "Relação"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   4245
         TabIndex        =   11
         Top             =   450
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "Alíquota"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   285
         TabIndex        =   10
         Top             =   870
         Width           =   885
         _ExtentX        =   1561
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
         Caption         =   "Referência"
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
         TabIndex        =   9
         Top             =   450
         Width           =   900
         _ExtentX        =   1588
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
         Caption         =   "Nome Taxa"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.ComboBox cboComponenteRef 
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
         Left            =   1230
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Componente Referência"
         Top             =   810
         Width           =   2895
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
         Left            =   4980
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Valor"
         Top             =   405
         Width           =   765
      End
      Begin VB.ComboBox cboTaxa 
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
         Left            =   1230
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Componente"
         Top             =   390
         Width           =   2865
      End
      Begin VB.ComboBox cboUnidade 
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
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Unidade"
         Top             =   390
         Width           =   1215
      End
      Begin VB.ComboBox cboRelacao 
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
         Left            =   4980
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Componente Referência"
         Top             =   810
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1875
      Top             =   930
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   1
      Left            =   6030
      TabIndex        =   6
      Top             =   2040
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
      Left            =   4845
      TabIndex        =   5
      Top             =   2040
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
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1138
      Icone           =   "TMPU103.frx":000C
   End
End
Attribute VB_Name = "TMPU103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function PegaUnidade(Unidade As String) As Integer
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "SELECT TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =5 and TGE_NOME='" & Unidade & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        PegaUnidade = rs(0)
    End If
    Bdados.FechaTabela rs
End Function

Private Function PegaReferencia(Referencia As String) As Integer
    Dim Sql As String
    Dim rs As VSRecordset
    
    Sql = "SELECT tgc_cod_grupo from Tab_Grupo_Componente Where tgc_nome='" & Referencia & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        PegaReferencia = rs(0)
    End If
    Bdados.FechaTabela rs
End Function

Private Function PegaRelacao(Relacao As String) As Integer
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select tco_cod_componente from tab_componente where tco_grupo=("
    Sql = Sql & "SELECT tgc_cod_grupo from Tab_Grupo_Componente Where tgc_nome='" & Relacao & "') and tco_descricao_componente='SIM' and tco_tmu_cod_municipio =" & Temp.PegaParametro(Bdados, "MUNICIPIO")
    If Bdados.AbreTabela(Sql, rs) Then
        PegaRelacao = rs(0)
    End If
    Bdados.FechaTabela rs
End Function

Private Sub cabVisual_GotFocus()

End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim Campos As String
    Dim Taxa As String
    Dim Unidade As Integer
    Dim Referencia As String
    Dim Relacao As String
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Taxa = BuscaCodigo("Select tip_cod_imposto From Tab_Imposto where tip_nome_imposto ='" & cboTaxa.Text & "'")
            Unidade = PegaUnidade(cboUnidade.Text)
            Referencia = PegaReferencia(cboComponenteRef.Text)
            Relacao = PegaRelacao(cboRelacao.Text)
            Valores = Bdados.PreparaValor(Taxa, txtValor, Unidade, Referencia, Relacao)
            Campos = "tti_tip_cod_imposto,tti_aliquota,tti_tmo_cod_unidade,tti_tco_cod_componente,tti_tco_cod_componente_relacao"
            Call Bdados.GravaDados("Tab_Taxa_Iptu", Valores, Campos, "tti_tip_cod_imposto='" & Taxa & "'")
            Call Util.Informa("Transação Completada.")
            Edita.LimpaCampos Me
            cboTaxa.SetFocus
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    cboComponenteRef.Clear
    cboTaxa.Clear
    cboUnidade.Clear
    cboRelacao.Clear
    Call Edita.AtualizaCombo(Bdados, cboTaxa, "Select tip_nome_imposto From Tab_Imposto " & _
        " WHERE tip_cod_imposto in (Select tpi_tip_cod_imposto from Tab_Parametro_Imposto where tpi_tipo_tributo=2)")
    Call Edita.AtualizaCombo(Bdados, cboUnidade, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =5 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboComponenteRef, "Select tgc_nome From Tab_Grupo_Componente where tgc_cod_grupo in(101,103,108,109,110,111,112,113,1000)")
    Call Edita.AtualizaCombo(Bdados, cboRelacao, "Select tgc_nome From Tab_Grupo_Componente where tgc_cod_grupo in(37,38,39,40)")
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

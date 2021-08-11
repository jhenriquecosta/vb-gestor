VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1290
      Left            =   45
      TabIndex        =   8
      Top             =   690
      Width           =   8295
      _ExtentX        =   14631
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
         Left            =   4200
         TabIndex        =   12
         Top             =   450
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "Quando for"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   810
         TabIndex        =   11
         Top             =   885
         Width           =   360
         _ExtentX        =   635
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
         Caption         =   "Vale"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   4380
         TabIndex        =   10
         Top             =   885
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "Por Unid."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   450
         Width           =   1080
         _ExtentX        =   1905
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
         Caption         =   "Componente"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
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
         ItemData        =   "TMPU102.frx":0000
         Left            =   2970
         List            =   "TMPU102.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Unidade"
         Top             =   825
         Width           =   1125
      End
      Begin VB.ComboBox cboComponente 
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
         ItemData        =   "TMPU102.frx":002E
         Left            =   1230
         List            =   "TMPU102.frx":0030
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Componente"
         Top             =   390
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
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Valor"
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboValorComponente 
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
         ItemData        =   "TMPU102.frx":0032
         Left            =   5220
         List            =   "TMPU102.frx":0034
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Valor Componente"
         Top             =   390
         Width           =   2985
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
         ItemData        =   "TMPU102.frx":0036
         Left            =   5220
         List            =   "TMPU102.frx":0038
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Componente Referência"
         Top             =   825
         Width           =   2985
      End
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   7215
      TabIndex        =   7
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
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   6030
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
   Begin VTOcx.cmdVISUAL cmdImprimir 
      Height          =   375
      Left            =   4845
      TabIndex        =   6
      Top             =   2040
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
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
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1138
      Icone           =   "TMPU102.frx":003A
   End
End
Attribute VB_Name = "TMPU102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto

Private Function PegaCodComponente(Componente As String, Grupo As String) As Integer
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(Componente) = "" Then
        Sql = "SELECT tco_cod_componente from Tab_Componente Where " & _
        " tco_grupo in (select " & _
        " tgc_cod_grupo from Tab_Grupo_Componente where tgc_nome='" & Grupo & "') and tco_tmu_cod_municipio =" & Temp.PegaParametro(Bdados, "MUNICIPIO")
    Else
        Sql = "SELECT tco_cod_componente from Tab_Componente Where " & _
        " tco_descricao_componente='" & Componente & "' and tco_grupo in (select " & _
        " tgc_cod_grupo from Tab_Grupo_Componente where tgc_nome='" & Grupo & "') and tco_tmu_cod_municipio =" & Temp.PegaParametro(Bdados, "MUNICIPIO")
    End If
    If Bdados.AbreTabela(Sql, rs) Then
            PegaCodComponente = rs(0)
    End If
    Bdados.FechaTabela rs
End Function

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

Private Sub cboComponente_Click()
    Dim Sql As String
    Dim rs As VSRecordset
    cboValorComponente.Clear
    Sql = "Select tco_descricao_componente from tab_componente where tco_grupo in " & _
        "(select tgc_cod_grupo from Tab_Grupo_Componente where tgc_nome='" & cboComponente & "')"
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            cboValorComponente.AddItem "" & rs(0)
            rs.MoveNext
        Loop
        cboValorComponente.Tag = "Valor Componente"
    Else
        cboValorComponente.Tag = ""
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub cboValorComponente_Click()
    Dim CodComponente As Integer
    Dim CodReferencia As Integer
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Indice As Integer
    
    CodComponente = PegaCodComponente(cboValorComponente.Text, cboComponente.Text)
    Campos = "tco_valor,tco_unid_moneta,tco_cod_componente_fator_calculo"
    Sql = "Select " & Campos & " FROM tab_componente where tco_cod_componente=" & CodComponente
    If Bdados.AbreTabela(Sql, rs) Then
        txtValor = IIf(rs(0) = 0, "", rs(0))
        cboUnidade.ListIndex = IIf(rs(1) = 0, -1, rs(1) - 1)
        If IsNull(rs(2)) Then
            Indice = -1
        Else
            Indice = BuscaIndiceCombo(cboComponenteRef, "tab_componente", "tco_cod_componente", "tco_descricao_componente", rs(2))
        End If
        cboComponenteRef.ListIndex = Indice
    Else
        txtValor = ""
        cboUnidade.ListIndex = -1
        cboComponenteRef.ListIndex = -1
    End If
    Bdados.FechaTabela rs
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
    Dim CodComponente As Integer
    Dim CodReferencia As Integer
    Dim UnidadeMoneta As Integer
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    CodComponente = PegaCodComponente(cboValorComponente.Text, cboComponente.Text)
    CodReferencia = PegaReferencia(cboComponenteRef.Text)
    UnidadeMoneta = PegaUnidade(cboUnidade)
    Campos = "tco_valor,tco_unid_moneta,tco_cod_componente_fator_calculo"
    Valores = Bdados.PreparaValor(Bdados.Converte(CDbl(txtValor), TCDuplo), UnidadeMoneta, CodReferencia)
    Call Bdados.GravaDados("Tab_Componente", Valores, Campos, "tco_cod_componente=" & CodComponente & " and and tco_tmu_cod_municipio =" & Temp.PegaParametro(Bdados, "MUNICIPIO"))
    Informa "Transação completada."
    Edita.LimpaCampos Me
    cboComponente.SetFocus
End Sub

Private Sub Form_Activate()
    cboComponente.Clear
    cboValorComponente.Clear
    cboComponenteRef.Clear
    Call Edita.AtualizaCombo(Bdados, cboComponente, "Select tgc_nome From Tab_Grupo_Componente where tgc_cod_grupo not in(1,37,38,39,40)")
    Call Edita.AtualizaCombo(Bdados, cboUnidade, "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =5 ORDER BY TGE_CODIGO ASC")
    Call Edita.AtualizaCombo(Bdados, cboComponenteRef, "Select tgc_nome From Tab_Grupo_Componente where tgc_cod_grupo in(1000,100,101,103,104,105,108,109,110,200)")
    'cboComponenteRef.AddItem "VALOR FIXO"
    cboValorComponente.Clear
    Set cadastro = New VSImposto
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValor_LostFocus()
    txtValor = Edita.FormataTexto(txtValor, Monetario, True)
End Sub


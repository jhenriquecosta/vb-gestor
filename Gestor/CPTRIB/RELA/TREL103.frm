VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREL103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Dinâmico"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleMode       =   0  'User
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPrioridade 
      Height          =   315
      ItemData        =   "TREL103.frx":0000
      Left            =   2250
      List            =   "TREL103.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Prioridade"
      Top             =   5070
      Width           =   1890
   End
   Begin VB.TextBox txtTabelaVinculo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3090
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   7
      ToolTipText     =   "Número correspondente a que a atual está vinculada"
      Top             =   5610
      Width           =   540
   End
   Begin VB.TextBox txtNovo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5055
      Width           =   540
   End
   Begin VB.ComboBox cboTabelaFilho 
      Height          =   315
      Left            =   4935
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Tabela Filho"
      Top             =   915
      Width           =   4680
   End
   Begin VB.ComboBox cboTabelaPai 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "Tabela Pai"
      Top             =   915
      Width           =   4680
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   3
      Left            =   165
      TabIndex        =   8
      Top             =   645
      Width           =   840
      _ExtentX        =   1482
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
      Caption         =   "Tabela Pai"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   0
      Left            =   4935
      TabIndex        =   9
      Top             =   645
      Width           =   990
      _ExtentX        =   1746
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
      Caption         =   "Tabela Filho"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   1
      Left            =   1650
      TabIndex        =   10
      Top             =   5640
      Width           =   1395
      _ExtentX        =   2461
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
      Caption         =   "Tabela Vinculada"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin VTOcx.grdVISUAL lstJoin 
      Height          =   3675
      Left            =   180
      TabIndex        =   11
      Top             =   1290
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   4339
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TREL103.frx":002B
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8550
      TabIndex        =   5
      Top             =   5070
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdGravar 
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   5070
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdExcluir 
      Height          =   375
      Left            =   6090
      TabIndex        =   4
      Top             =   5070
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   2
      Left            =   1245
      TabIndex        =   13
      Top             =   5115
      Width           =   915
      _ExtentX        =   1614
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
      Caption         =   "Prioridade:"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
End
Attribute VB_Name = "TREL103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AtualizaJoin As Boolean

Private Sub cboTabelaPai_Click()
Dim Sql As String
    If AtualizaJoin Then
        Sql = " SELECT dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN AS [Join],PAI.TTA_ALIAS_USUARIO AS Pai,FILHO.TTA_ALIAS_USUARIO AS Filho, " & _
              " TJO_COD_JOIN_VINCULADO AS Vinculado,TJO_TIPO_JOIN AS Prioridade FROM dbo.TAB_TABELAS PAI INNER JOIN dbo.TAB_JOINS_TABELAS ON " & _
              " Pai.TTA_COD_TABELA = dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_PAI Inner Join " & _
              " dbo.TAB_TABELAS FILHO ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_FILHO = Filho.TTA_COD_TABELA "
        If cboTabelaPai <> "" Then Sql = Sql & " WHERE PAI.TTA_ALIAS_USUARIO = '" & cboTabelaPai & "'"
        lstJoin.Preencher Bdados, Sql, 1000, 3500, 3500, 0, 0
        Call BuscaUltimo
    End If
End Sub

Private Sub cmdGravar_Click()
    
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim Tabela As Integer
    Dim Pai As Integer
    Dim Filho As Integer
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    'If Not ExcluiJoinTabelas Then: Util.Avisa "Houve um erro, o join não foi criado.": Exit Sub
    
    Bdados.AbreTabela ("SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelaPai & "'")
    Pai = Bdados.Tabela(0)
    Bdados.AbreTabela ("SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelaFilho & "'")
    Filho = Bdados.Tabela(0)
    If Pai > 0 And Filho > 0 Then
        Valores = Bdados.PreparaValor(Val(txtNovo), Pai, Filho, IIf(Trim(txtTabelaVinculo) = "", 0, Val(txtTabelaVinculo)), cboPrioridade.ListIndex)
        Campos = "TJO_COD_JOIN,TJO_TTA_COD_TABELA_PAI,TJO_TTA_COD_TABELA_FILHO,TJO_COD_JOIN_VINCULADO,TJO_TIPO_JOIN"
        Condicao = "TJO_COD_JOIN = " & Val(txtNovo)
        If Bdados.GravaDados("TAB_JOINS_TABELAS", Valores, Campos, Condicao) Then
            Util.Informa "Join montado e gravados."
            Call cboTabelaPai_Click
            AtualizaJoin = True
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaJoin = True
    Call MostraTabelas
    Call cboTabelaPai_Click
    AtualizaCabecalho lstJoin
End Sub

Private Sub BuscaUltimo()
    Bdados.AbreTabela "SELECT MAX(TJO_COD_JOIN) FROM TAB_JOINS_TABELAS"
    txtNovo = Nvl("" & Bdados.Tabela(0), 0) + 1
End Sub

Private Sub MostraTabelas()
Dim Sql As String
    Sql = " SELECT TTA_ALIAS_USUARIO FROM TAB_TABELAS ORDER BY 1"
    Edita.AtualizaCombo Bdados, cboTabelaPai, Sql
    Edita.AtualizaCombo Bdados, cboTabelaFilho, Sql
End Sub

Private Sub lstJoin_DblClick()
    txtNovo = ""
    AtualizaJoin = False
    If lstJoin.ListItems.Count > 0 Then
        txtNovo = lstJoin.SelectedItem
        cboTabelaPai = lstJoin.SelectedItem.SubItems(1)
        cboTabelaFilho = lstJoin.SelectedItem.SubItems(2)
        cboPrioridade.ListIndex = Nvl(lstJoin.SelectedItem.SubItems(4), 0)
    End If
    AtualizaJoin = True
End Sub

Private Sub cmdExcluir_Click()
Dim CodTabela As Integer
Dim RsTabela As VSRecordset
Dim Escolha As Integer

Escolha = MsgBox("Excluir o campo selecionado?", vbQuestion + vbOKCancel, "Exclusão")

    If Escolha = vbOK Then
        If ExcluiJoinTabelas Then
            Util.Informa "Join entre as tabelas selecionadas foi apagado."
            Call cboTabelaPai_Click
        Else
            Util.Informa "Join entre as tabelas não pode ser apagado."
        End If
    End If
End Sub

Private Function ExcluiJoinTabelas() As Boolean
Dim CodTabela As Integer
Dim RsTabela As VSRecordset
    
    CodTabela = Val(txtNovo)
    ' CASCATA PARA ELIMINAR, CAMPOS, JOIN, E TABELAS
    If Bdados.AbreTabela("SELECT TJO_COD_JOIN FROM TAB_JOINS_TABELAS WHERE TJO_TTA_COD_TABELA_PAI = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO  ='" & cboTabelaPai & "') OR TJO_TTA_COD_TABELA_FILHO = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO  ='" & cboTabelaPai & "')", RsTabela) Then
        ' APAGAR JOIN ENTRE CAMPOS
        Bdados.Executa "DELETE FROM TAB_CAMPOS_JOIN WHERE TCJ_TJO_COD_JOIN =" & CodTabela  'RsTabela!TJO_COD_JOIN
    End If
    ' APAGAR JOIN ENTRE TABELAS
    If Bdados.DeletaDados("TAB_JOINS_TABELAS", "TJO_COD_JOIN = " & CodTabela) Then    'RsTabela!TJO_COD_JOIN
        ExcluiJoinTabelas = True
    End If
End Function

Private Sub txtNovo_DblClick()
    Call BuscaUltimo
End Sub

Private Sub txtTabelaVinculo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

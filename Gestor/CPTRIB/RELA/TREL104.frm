VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TREL104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Dinâmico"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleMode       =   0  'User
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSentido 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   930
      TabIndex        =   7
      Top             =   5055
      Width           =   540
   End
   Begin VB.ComboBox cboCampoPai 
      Height          =   315
      Left            =   165
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Campo da Tabela Pai"
      Top             =   1590
      Width           =   4680
   End
   Begin VB.ComboBox cboCampoFilho 
      Height          =   315
      Left            =   4935
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Campo da Tabela Filho"
      Top             =   1590
      Width           =   4680
   End
   Begin VB.ComboBox cboTabelaFilho 
      Height          =   315
      Left            =   4935
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Tabela Filho"
      Top             =   915
      Width           =   4680
   End
   Begin VB.ComboBox cboTabelaPai 
      Height          =   315
      Left            =   165
      Sorted          =   -1  'True
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
      Left            =   165
      TabIndex        =   10
      Top             =   1320
      Width           =   1470
      _ExtentX        =   2593
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
      Caption         =   "Campo Tabela Pai"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   2
      Left            =   4935
      TabIndex        =   11
      Top             =   1320
      Width           =   1620
      _ExtentX        =   2858
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
      Caption         =   "Campo Tabela Filho"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   4
      Left            =   165
      TabIndex        =   12
      Top             =   5085
      Width           =   645
      _ExtentX        =   1138
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
      Caption         =   "Sentido"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin VTOcx.grdVISUAL lstJoin 
      Height          =   2925
      Left            =   120
      TabIndex        =   13
      Top             =   2070
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   4339
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Threed.SSPanel lblSelecionado 
      Height          =   225
      Left            =   1875
      TabIndex        =   14
      Top             =   5085
      Width           =   2745
      _ExtentX        =   4842
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
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   1
      RoundedCorners  =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TREL104.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8550
      TabIndex        =   6
      Top             =   5130
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
      TabIndex        =   4
      Top             =   5130
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
      TabIndex        =   5
      Top             =   5130
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TREL104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodJoinAtual As Integer
Dim CodJoinCampos As Integer

Private Sub cboTabelaFilho_Click()
Dim Tabela As Integer
Dim Sql As String
Dim Escolha As Boolean

    
    If cboTabelaPai <> "" And cboTabelaFilho <> "" Then
        Sql = "SELECT COLUNA From Vis_Campos_Tabelas WHERE TABELA = (SELECT TTA_NOME FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO = '" & cboTabelaFilho & "')"
        Edita.AtualizaCombo Bdados, cboCampoFilho, Sql
        
        ' TRAS O JOIN DAS TABELAS ENVOLVIDAS
        Sql = " SELECT dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN AS [JOIN] FROM dbo.TAB_JOINS_TABELAS INNER JOIN " & _
              " dbo.TAB_TABELAS PAI ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_PAI = Pai.TTA_COD_TABELA " & _
              " Inner Join dbo.TAB_TABELAS FILHO ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_FILHO = Filho.TTA_COD_TABELA "
        Sql = Sql & " WHERE (PAI.TTA_ALIAS_USUARIO = '" & cboTabelaPai & "') AND (FILHO.TTA_ALIAS_USUARIO = '" & cboTabelaFilho & "') "
        If Bdados.AbreTabela(Sql) Then
            Tabela = IIf(Bdados.Tabela.EOF, Val(""), Bdados.Tabela(0))
            CodJoinAtual = IIf(Bdados.Tabela.EOF, 0, Bdados.Tabela(0))
        ' TRAS OS JOIN'S JA CRIADOS ANTERIORMENTE PARA AS TABELAS ENVOLVIDAS
            Sql = " SELECT dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN, dbo.TAB_CAMPOS_JOIN.TCJ_NOME_CAMPO_PAI, " & _
                  " dbo.TAB_CAMPOS_JOIN.TCJ_NOME_CAMPO_FILHO, dbo.TAB_CAMPOS_JOIN.TCJ_SENTIDO_RELACIONAMENTO SENTIDO FROM dbo.TAB_CAMPOS_JOIN INNER JOIN dbo.TAB_JOINS_TABELAS ON " & _
                  " dbo.TAB_CAMPOS_JOIN.TCJ_TJO_COD_JOIN = dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN "
            Sql = Sql & " WHERE dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN = " & Tabela
            lblSelecionado = ""
            lstJoin.Preencher Bdados, Sql
        End If
    End If

End Sub

Private Sub cboTabelaPai_Click()
Dim Sql As String
Dim Tabela As Integer

    cboTabelaFilho.Clear
    cboCampoPai.Clear
    cboCampoFilho.Clear
    lstJoin.Preencher Bdados, ""
    
    If cboTabelaPai <> "" Then
        If Bdados.AbreTabela("SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelaPai & "'") Then
            Tabela = Bdados.Tabela(0)
            Sql = " SELECT FILHO.TTA_ALIAS_USUARIO AS FILHO FROM dbo.TAB_JOINS_TABELAS INNER JOIN " & _
                  " dbo.TAB_TABELAS PAI ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_PAI = Pai.TTA_COD_TABELA Inner Join " & _
                  " dbo.TAB_TABELAS FILHO ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_FILHO = FILHO.TTA_COD_TABELA "
            Sql = Sql & " WHERE PAI.TTA_COD_TABELA = " & Tabela
            Edita.AtualizaCombo Bdados, cboTabelaFilho, Sql
            
            Sql = "SELECT COLUNA From Vis_Campos_Tabelas WHERE TABELA = (SELECT TTA_NOME FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO = '" & cboTabelaPai & "')"
            Edita.AtualizaCombo Bdados, cboCampoPai, Sql
        End If
    End If
End Sub


Private Sub cmdExcluir_Click()
Dim CodTabela As Integer
Dim RsTabela As VSRecordset
Dim Escolha As Boolean

Escolha = MsgBox("Excluir o join entre os campos selecionados?", vbQuestion + vbOKCancel, "Exclusão")
If Escolha = vbOK Then

    ' CASCATA PARA ELIMINAR JOIN ENTRE CAMPOS
    ' APAGAR JOIN ENTRE CAMPOS
    If Bdados.Executa("DELETE FROM TAB_CAMPOS_JOIN WHERE TCJ_TJO_COD_JOIN =" & CodJoinAtual & " AND TCJ_NOME_CAMPO_PAI='" & cboCampoPai & "' AND TCJ_NOME_CAMPO_FILHO = '" & cboCampoFilho & "'") Then
        Call cboTabelaFilho_Click
        Util.Informa "Relacionamento entre os campos escolhidos foi apagado com sucesso."
    Else
        Util.Avisa "Não foi possível apagar o relacionamento entre os campos."
    End If
End If
End Sub

Private Sub cmdGravar_Click()
    
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim TabelaPai As Integer
    Dim TabFilho As Integer
    Dim Sql As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    Sql = "SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelaPai & "'"
    If Bdados.AbreTabela(Sql) Then
        TabelaPai = Bdados.Tabela(0)
        Sql = "SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelaFilho & "'"
        If Bdados.AbreTabela(Sql) Then
            TabFilho = Bdados.Tabela(0)
            
            Valores = Bdados.PreparaValor(CodJoinAtual, cboCampoPai, cboCampoFilho, IIf(Trim(txtSentido) <> "", Val(txtSentido), 0))
            Campos = "TCJ_TJO_COD_JOIN,TCJ_NOME_CAMPO_PAI,TCJ_NOME_CAMPO_FILHO,TCJ_SENTIDO_RELACIONAMENTO"
            Condicao = "TCJ_TJO_COD_JOIN = " & CodJoinAtual & " And TCJ_NOME_CAMPO_PAI = '" & cboCampoPai & "'"
            If Bdados.GravaDados("TAB_CAMPOS_JOIN", Valores, Campos, Condicao) Then
                Util.Informa "Join montado e gravado."
                Call MostraJoin
                lblSelecionado = ""
            End If
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
    Call MostraTabelas
    AtualizaCabecalho lstJoin
End Sub

Private Sub MostraTabelas()
Dim Sql As String
    Sql = " SELECT TTA_ALIAS_USUARIO FROM TAB_TABELAS ORDER BY 1"
    Edita.AtualizaCombo Bdados, cboTabelaPai, Sql
    Edita.LimpaCampos Me
End Sub

Private Sub lstJoin_DblClick()
    If lstJoin.ListItems.Count > 0 Then
        CodJoinCampos = lstJoin.SelectedItem
        lblSelecionado = "Join " & CodJoinCampos & " selecionado"
        cboCampoPai = LCase(lstJoin.SelectedItem.SubItems(1))
        cboCampoFilho = LCase(lstJoin.SelectedItem.SubItems(2))
        txtSentido = LCase(lstJoin.SelectedItem.SubItems(3))
    End If
End Sub

Private Sub txtSentido_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub MostraJoin()
Dim Sql As String

    Sql = " SELECT dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN AS [JOIN] FROM dbo.TAB_JOINS_TABELAS INNER JOIN " & _
          " dbo.TAB_TABELAS PAI ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_PAI = Pai.TTA_COD_TABELA " & _
          " Inner Join dbo.TAB_TABELAS FILHO ON dbo.TAB_JOINS_TABELAS.TJO_TTA_COD_TABELA_FILHO = Filho.TTA_COD_TABELA "
    Sql = Sql & " WHERE (PAI.TTA_ALIAS_USUARIO = '" & cboTabelaPai & "') AND (FILHO.TTA_ALIAS_USUARIO = '" & cboTabelaFilho & "') "
    
    If Bdados.AbreTabela(Sql) Then
    ' TRAS OS JOIN'S JA CRIADOS ANTERIORMENTE PARA AS TABELAS ENVOLVIDAS
        Sql = " SELECT dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN, dbo.TAB_CAMPOS_JOIN.TCJ_NOME_CAMPO_PAI, " & _
              " dbo.TAB_CAMPOS_JOIN.TCJ_NOME_CAMPO_FILHO, dbo.TAB_CAMPOS_JOIN.TCJ_SENTIDO_RELACIONAMENTO SENTIDO FROM dbo.TAB_CAMPOS_JOIN INNER JOIN dbo.TAB_JOINS_TABELAS ON " & _
              " dbo.TAB_CAMPOS_JOIN.TCJ_TJO_COD_JOIN = dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN "
        Sql = Sql & " WHERE dbo.TAB_JOINS_TABELAS.TJO_COD_JOIN = " & IIf(Bdados.Tabela.EOF, Val(""), Bdados.Tabela(0))
        
        lstJoin.Preencher Bdados, Sql
    End If
End Sub

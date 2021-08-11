VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TREL101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Dinâmico"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleMode       =   0  'User
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fraDados 
      Height          =   1470
      Index           =   3
      Left            =   60
      TabIndex        =   8
      Top             =   5775
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   2593
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
      Begin VB.ComboBox cboTabelaVinculada 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   990
         Width           =   4545
      End
      Begin VB.ComboBox cboNomeTabela 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Tabela"
         Top             =   360
         Width           =   4545
      End
      Begin VB.TextBox txtCondicao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6030
         MaxLength       =   80
         TabIndex        =   4
         Top             =   1020
         Width           =   4095
      End
      Begin VB.TextBox txtApelido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6030
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Apelido"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtCodTabela 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "Cod Tabela"
         Top             =   375
         Width           =   990
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   23
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
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
         Caption         =   "Cod Tabela"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   1365
         TabIndex        =   10
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
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
         Caption         =   "Nome Tabela"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   6030
         TabIndex        =   11
         Top             =   105
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
         Caption         =   "Apelido"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   6030
         TabIndex        =   12
         Top             =   750
         Width           =   780
         _ExtentX        =   1376
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
         Caption         =   "Condição"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   1365
         TabIndex        =   13
         Top             =   750
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
   End
   Begin VTOcx.grdVISUAL lstTabelas 
      Height          =   5025
      Left            =   30
      TabIndex        =   14
      Top             =   720
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   4339
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TREL101.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   10140
      TabIndex        =   7
      Top             =   7350
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
      Left            =   8910
      TabIndex        =   5
      Top             =   7350
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
      Left            =   7680
      TabIndex        =   6
      Top             =   7350
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TREL101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Busca As Boolean

Private Sub cmdExcluir_Click()
Dim Escolha As Integer

    Escolha = MsgBox("Excluir a tabela selecionada?", vbQuestion + vbOKCancel, "Exclusão")
    If Escolha = vbOK Then
        If ExcluiTabela Then
            Util.Informa "Tabela, campos e relacionamentos apagados."
        Else
            Util.Avisa "Não foi possível efetuar a exclusão, a tabela não foi apagada."
        End If
        txtCodTabela = ""
        txtApelido = ""
        
    End If
End Sub

Private Sub cmdGravar_Click()
    
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim Vinculo As Integer
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Bdados.AbreTabela "SELECT TTA_ALIAS_USUARIO FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & txtApelido & "' AND TTA_COD_TABELA <> " & Val(txtCodTabela)
    If Not Bdados.Tabela.EOF Then
        Util.Avisa "Apelido já cadastrado, para continuar o cadastro verifique os dados inseridos."
        txtApelido.SetFocus
        Exit Sub
    End If
    
    If Bdados.AbreTabela("SELECT TTA_NOME FROM TAB_TABELAS WHERE TTA_COD_TABELA = " & Val(txtCodTabela)) Then
        ' SE O USUÁRIO ESCOLHEU MUDAR A TABELA E PERMANECEU O ALIASE, APAGA AS TABELAS, JOINS E CAMPOS RELACIONADOS
        If Bdados.Tabela(0) <> cboNomeTabela Then
            If Not ExcluiTabela Then Util.Informa "Não foi possível gravar alteração.": Exit Sub
        End If
    End If
    Bdados.AbreTabela "SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO = '" & cboTabelaVinculada & "'"
    Vinculo = IIf(Bdados.Tabela.EOF, 0, Bdados.Tabela(0))
    
    
    Valores = Bdados.PreparaValor(Val(txtCodTabela), UCase(cboNomeTabela), txtApelido, txtCondicao, Vinculo)
    Campos = "TTA_COD_TABELA,TTA_NOME,TTA_ALIAS_USUARIO,TTA_WHERE,TTA_TTA_COD_TABELA_ORIGEM"
    Condicao = "TTA_COD_TABELA = " & Val(txtCodTabela)
    If Bdados.GravaDados("TAB_TABELAS", Valores, Campos, Condicao) Then
        Util.Informa "Tabela gravada."
        lstTabelas.Preencher Bdados, "SELECT * FROM TAB_TABELAS ORDER BY 1"
        MostraTabelas
        txtCodTabela = ""
        txtApelido = ""
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
    KeyAscii = Edita.Maiuscula(KeyAscii)
    AtualizaCabecalho lstTabelas
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Busca = True
    Call MostraTabelas
End Sub

Private Sub MostraTabelas()
Dim Sql As String
' *****************************************************************
' MOSTRA AS TABELAS CADASTRADAS PARA O RELATÓRIO DINÂMICO
' *****************************************************************

    lstTabelas.Preencher Bdados, "SELECT * FROM TAB_TABELAS ORDER BY 1"
    
    Sql = " SELECT NAME FROM sysobjects WHERE xtype in ('U','V') AND (uid = 1) AND (name NOT IN ('dtproperties', " & _
            " 'TAB_JOINS_TABELAS', 'TAB_CAMPOS_JOIN', 'TAB_TABELAS_AUXILIARES_JOIN', 'TAB_CAMPOS_TABELAS', " & _
            " 'TAB_TABELAS')) order by name"
    Edita.AtualizaCombo Bdados, cboNomeTabela, Sql
    
    Sql = "SELECT TTA_ALIAS_USUARIO FROM TAB_TABELAS ORDER BY TTA_ALIAS_USUARIO ASC"
    Edita.AtualizaCombo Bdados, cboTabelaVinculada, Sql
    cboTabelaVinculada.AddItem " "
    Edita.LimpaCampos Me
End Sub

Private Sub BuscaUltimo()
    Bdados.AbreTabela "SELECT MAX(TTA_COD_TABELA) FROM TAB_TABELAS"
    txtCodTabela = Nvl("" & Bdados.Tabela(0), 0) + 1
    cboNomeTabela.SetFocus
End Sub

Private Sub lstTabelas_DblClick()
    Busca = False
    
    txtCodTabela = lstTabelas.SelectedItem
    cboNomeTabela = lstTabelas.SelectedItem.SubItems(1)
    txtApelido = lstTabelas.SelectedItem.SubItems(2)
    txtCondicao = lstTabelas.SelectedItem.SubItems(3)
    
    Bdados.AbreTabela "SELECT TTA_ALIAS_USUARIO FROM TAB_TABELAS WHERE TTA_COD_TABELA =" & Val(lstTabelas.SelectedItem.SubItems(4))
        cboTabelaVinculada = IIf(Not Bdados.Tabela.EOF, Bdados.Tabela(0), " ")
    Busca = True
End Sub

Private Sub txtCodTabela_DblClick()
    Call BuscaUltimo
End Sub

Private Function ExcluiTabela() As Boolean
Dim CodTabela As Integer
Dim RsTabela As VSRecordset
    
    CodTabela = Val(txtCodTabela)
    ' CASCATA PARA ELIMINAR, CAMPOS, JOIN, E TABELAS
        If Bdados.AbreTabela("SELECT TJO_COD_JOIN FROM TAB_JOINS_TABELAS WHERE TJO_TTA_COD_TABELA_PAI = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO  ='" & lstTabelas.SelectedItem.SubItems(2) & "') OR TJO_TTA_COD_TABELA_FILHO = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO  ='" & lstTabelas.SelectedItem.SubItems(2) & "')", RsTabela) Then
            ' APAGAR JOIN ENTRE CAMPOS
            While Not RsTabela.EOF
                Bdados.Executa "DELETE FROM TAB_CAMPOS_JOIN WHERE TCJ_TJO_COD_JOIN =" & RsTabela!TJO_COD_JOIN
                RsTabela.MoveNext
            Wend
            RsTabela.MoveFirst
        End If
        ' APAGAR CAMPOS DAS TABELAS
        Bdados.DeletaDados "TAB_CAMPOS_TABELAS", "TCP_TTA_COD_TABELA = " & CodTabela
        ' APAGAR JOIN ENTRE TABELAS
        
        While Not RsTabela.EOF
            Bdados.DeletaDados "TAB_JOINS_TABELAS", "TJO_COD_JOIN = " & RsTabela!TJO_COD_JOIN
            RsTabela.MoveNext
        Wend
            ' APAGAR ENFIM AS TABELAS DA TAB_TABELAS
        If Bdados.DeletaDados("TAB_TABELAS", "TTA_COD_TABELA = " & CodTabela) Then
            lstTabelas.Preencher Bdados, "SELECT * FROM TAB_TABELAS ORDER BY 1"
            ExcluiTabela = True
        End If
End Function

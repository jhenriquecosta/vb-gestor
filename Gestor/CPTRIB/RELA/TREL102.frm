VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TREL102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Dinâmico"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleMode       =   0  'User
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fraDados 
      Height          =   870
      Index           =   3
      Left            =   135
      TabIndex        =   9
      Top             =   4800
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1535
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
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   7890
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Campo"
         Top             =   345
         Width           =   1515
      End
      Begin VB.TextBox txtTamaho 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6795
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Cod Campo"
         Top             =   375
         Width           =   915
      End
      Begin VB.ComboBox cboNomeCampo 
         Height          =   315
         Left            =   1420
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Campo"
         Top             =   360
         Width           =   3300
      End
      Begin VB.TextBox txtApelido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4830
         MaxLength       =   30
         TabIndex        =   2
         Tag             =   "Apelido"
         Top             =   375
         Width           =   1815
      End
      Begin VB.TextBox txtCodCampo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   8
         Tag             =   "Cod Campo"
         Top             =   375
         Width           =   990
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   23
         Left            =   240
         TabIndex        =   10
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
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
         Caption         =   "Cod Campo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   1470
         TabIndex        =   11
         Top             =   90
         Width           =   600
         _ExtentX        =   1058
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
         Caption         =   "Campo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   4830
         TabIndex        =   12
         Top             =   90
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
         Left            =   6780
         TabIndex        =   14
         Top             =   90
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
         Caption         =   "Tamanho"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   7905
         TabIndex        =   15
         Top             =   75
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "Tipo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.ComboBox cboTabelas 
      Height          =   315
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "Tabela"
      Top             =   735
      Width           =   4545
   End
   Begin Threed.SSPanel lbl 
      Height          =   225
      Index           =   3
      Left            =   75
      TabIndex        =   13
      Top             =   750
      Width           =   555
      _ExtentX        =   979
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
      Caption         =   "Tabela"
      BorderWidth     =   1
      BevelOuter      =   0
      AutoSize        =   1
      Alignment       =   0
      RoundedCorners  =   0   'False
   End
   Begin VTOcx.grdVISUAL lstCampos 
      Height          =   3705
      Left            =   60
      TabIndex        =   16
      Top             =   1110
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   4339
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TREL102.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   5760
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
      Left            =   7290
      TabIndex        =   5
      Top             =   5760
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
      Left            =   6060
      TabIndex        =   6
      Top             =   5760
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TREL102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboNomeCampo_Click()
txtTamaho = ""
    If Bdados.AbreTabela("SELECT TTA_NOME FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelas & "'") Then
        If Bdados.AbreTabela("SELECT TAMANHO FROM Vis_Campos_Tabelas WHERE COLUNA='" & cboNomeCampo & "' AND TABELA ='" & Bdados.Tabela(0) & "'") Then
            txtTamaho = IIf(IsNull(Bdados.Tabela(0)), 12, Bdados.Tabela(0))
        End If
    End If
End Sub

Private Sub cmdExcluir_Click()
Dim CodTabela As Integer
Dim RsTabela As VSRecordset
Dim Escolha As Integer

Escolha = MsgBox("Excluir o campo selecionado?", vbQuestion + vbOKCancel, "Exclusão")

If Escolha = vbOK Then
        If ExcluiCampo Then
            Util.Informa "Campos e relacionamentos apagados."
            txtCodCampo = ""
            txtApelido = ""
            txtTamaho = ""
            Call cboTabelas_Click
        Else
            Util.Avisa "Não foi possível excluir a relação do campo e relacionamentos, o campo não foi apagado."
        End If
End If

End Sub

Private Sub cmdGravar_Click()
    
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim Tabela As Integer
    Dim Tipo As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
        
    If Bdados.AbreTabela("SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelas & "'") Then
        Tabela = Bdados.Tabela(0)
        
        If Bdados.AbreTabela("SELECT TCP_NOME FROM TAB_CAMPOS_TABELAS WHERE TCP_COD_CAMPO = " & Val(txtCodCampo)) Then
        ' SE O USUÁRIO ESCOLHEU MUDAR A CAMPO E PERMANECEU O ALIASE, APAGA OS CAMPOS E JOINS RELACIONADOS
            If LCase(Bdados.Tabela(0)) <> LCase(cboNomeCampo) Then
                If Not ExcluiCampo Then Util.Informa "Não foi possível gravar alteração.": Exit Sub
            End If
        End If
        
        Bdados.AbreTabela "SELECT TCP_COD_CAMPO,TCP_ALIAS_USUARIO FROM TAB_CAMPOS_TABELAS WHERE TCP_ALIAS_USUARIO='" & txtApelido & "' AND TCP_TTA_COD_TABELA=" & Tabela
        If Not Bdados.Tabela.EOF Then
            If Val(txtCodCampo) <> Bdados.Tabela(0) Then
                Util.Avisa "Apelido de campo já cadastrado para a tabela escolhida."
                txtApelido.SetFocus
                Exit Sub
            End If
        End If
        
        Bdados.AbreTabela "SELECT TTA_NOME FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelas & "'"
        If Bdados.AbreTabela("SELECT TIPO FROM Vis_Campos_Tabelas WHERE TABELA='" & Bdados.Tabela(0) & "' AND COLUNA='" & cboNomeCampo & "'") Then
            If cboTipo.ListIndex <= 0 Then
                Tipo = IIf(Bdados.AbreTabela("SELECT TGE_CODIGO FROM TAB_GERAL WHERE TGE_TIPO=716 AND TGE_NOME ='" & Tipo & "'"), Bdados.Tabela(0), tipTexto)
            Else
                Tipo = cboTipo.ListIndex
            End If
            
            Valores = Bdados.PreparaValor(Val(txtCodCampo), Tabela, UCase(cboNomeCampo), Trim(txtApelido), Val(Tipo), Val(txtTamaho))
            Campos = "TCP_COD_CAMPO,TCP_TTA_COD_TABELA,TCP_NOME,TCP_ALIAS_USUARIO,TCP_TIPO,TCP_TAMANHO"
            Condicao = "TCP_COD_CAMPO = " & Val(txtCodCampo)
            
            If Bdados.GravaDados("TAB_CAMPOS_TABELAS", Valores, Campos, Condicao) Then
                Util.Informa "Campos gravados."
                Call cboTabelas_Click
                txtCodCampo = ""
                txtApelido = ""
                txtTamaho = ""
                cboTipo.ListIndex = -1
                Call BuscaUltimo
            End If
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cboTabelas_Click()
If cboTabelas = "" Then Exit Sub
    lstCampos.Preencher Bdados, "SELECT TCP_COD_CAMPO AS COD, TCP_NOME AS NOME,TCP_ALIAS_USUARIO AS ALIAS,TCP_TAMANHO AS TAM,TCP_TIPO AS TIPO FROM TAB_CAMPOS_TABELAS WHERE TCP_TTA_COD_TABELA = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelas & "')", 1000, 3000, 2800, 1000, 1000
    Call MostraCampos
    Call BuscaUltimo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaCombo Bdados, cboTipo, "Select tge_nome from tab_geral where tge_Codigo > 0 AND tge_TIPO = (select tge_tipo from tab_geral where tge_nome ='TIPO CAMPO')  ORDER BY tge_tipo ASC"
    cboTipo.AddItem " ", 0
    Call MostraTabelas
    AtualizaCabecalho lstCampos
End Sub

Private Sub MostraCampos()
    On Error Resume Next
Dim Sql As String
' *****************************************************************
' MOSTRA OS CAMPOS DAS TABELAS CADASTRADAS PARA O RELATÓRIO DINÂMICO
' *****************************************************************
    If Bdados.AbreTabela("SELECT TTA_NOME FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelas & "'") Then
        Sql = " SELECT COLUNA FROM Vis_Campos_Tabelas WHERE TABELA = '" & Bdados.Tabela(0) & "'"
        Edita.AtualizaCombo Bdados, cboNomeCampo, Sql
    End If
End Sub

Private Sub BuscaUltimo()
    Bdados.AbreTabela "SELECT MAX(TCP_COD_CAMPO) FROM TAB_CAMPOS_TABELAS"
    txtCodCampo = Nvl("" & Bdados.Tabela(0), 0) + 1
    cboNomeCampo.SetFocus
End Sub

Private Sub lstCampos_DblClick()
    On Error Resume Next
    txtCodCampo = lstCampos.SelectedItem
    cboNomeCampo = LCase(lstCampos.SelectedItem.SubItems(1))
    txtApelido = lstCampos.SelectedItem.SubItems(2)
    txtTamaho = lstCampos.SelectedItem.SubItems(3)
    cboTipo.ListIndex = lstCampos.SelectedItem.SubItems(4)
End Sub

Private Sub MostraTabelas()
    Dim Sql As String
    Sql = " SELECT TTA_ALIAS_USUARIO FROM TAB_TABELAS ORDER BY 1"
    Edita.AtualizaCombo Bdados, cboTabelas, Sql
End Sub

Private Sub txtCodCampo_Click()
    Call BuscaUltimo
End Sub

Private Sub txtTamaho_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Function ExcluiCampo() As Boolean
Dim CodTabela As Integer
Dim RsTabela As VSRecordset
    
    CodTabela = Val(txtCodCampo)
    ' CASCATA PARA ELIMINAR, CAMPOS E JOIN
        If Bdados.AbreTabela("SELECT TJO_COD_JOIN FROM TAB_JOINS_TABELAS WHERE TJO_TTA_COD_TABELA_PAI = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO  ='" & cboTabelas & "') OR TJO_TTA_COD_TABELA_FILHO = (SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO  ='" & cboTabelas & "')", RsTabela) Then
            ' APAGAR JOIN ENTRE CAMPOS
            While Not RsTabela.EOF
                Bdados.Executa "DELETE FROM TAB_CAMPOS_JOIN WHERE TCJ_TJO_COD_JOIN =" & RsTabela!TJO_COD_JOIN
                RsTabela.MoveNext
            Wend
        End If
        ' APAGAR CAMPO DA TABELA
        If Bdados.AbreTabela("SELECT TTA_COD_TABELA FROM TAB_TABELAS WHERE TTA_ALIAS_USUARIO='" & cboTabelas & "'") Then
            If Bdados.DeletaDados("TAB_CAMPOS_TABELAS", "TCP_TTA_COD_TABELA = " & Bdados.Tabela(0) & " AND TCP_COD_CAMPO = " & CodTabela) Then ExcluiCampo = True
        End If
End Function

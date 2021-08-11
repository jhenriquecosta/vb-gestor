VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PMUS101 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PMUS101.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1425
      Left            =   60
      TabIndex        =   8
      Top             =   690
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   2514
      Altura          =   1905
      Caption         =   " Usuário"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtMatricula 
         Height          =   285
         Left            =   60
         TabIndex        =   2
         Tag             =   "Código"
         Top             =   1020
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         Caption         =   "Matricula"
         Text            =   ""
         MaxLen          =   15
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Tag             =   "Código"
         Top             =   360
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   503
         Caption         =   "Codigo"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtNome 
         Height          =   285
         Left            =   345
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   697
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   503
         Caption         =   "Nome"
         Text            =   ""
         Restricao       =   1
         MaxLen          =   200
      End
   End
   Begin Cabecalho.rodVISUAL rodVisual 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   9
      Top             =   7080
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   926
      CorFrente       =   4210752
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   405
         Left            =   4440
         TabIndex        =   10
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   6510
         TabIndex        =   5
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   5475
         TabIndex        =   4
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   405
         Left            =   3405
         TabIndex        =   3
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraEstrutura 
      Height          =   4830
      Left            =   60
      TabIndex        =   7
      Top             =   2175
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   8520
      Altura          =   5160
      Caption         =   " Lotação"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      Ocultavel       =   0   'False
      Begin MSComctlLib.TreeView treeUsuarios 
         Height          =   4380
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   7726
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   538
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlMenu"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   1138
      Icone           =   "PMUS101.frx":08CA
   End
End
Attribute VB_Name = "PMUS101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lotacao As pLotacao
Dim Usuario As pUsuario

Private Sub cmdExcluir_Click()
    If Trim(txtCodigo) = "" Then Exit Sub
    Set Usuario = New pUsuario
        Usuario.Excluir txtCodigo
    Set Usuario = Nothing
End Sub

Private Sub cmdLimpar_Click()
    Set Usuario = New pUsuario
    Set Lotacao = New pLotacao
    
'    Usuario.ExibirUsuarios treeUsuarios
        
    Set Lotacao = Nothing
    Set Usuario = Nothing
    Edita.LimpaCampos Me
    txtCodigo.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
        
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Set Usuario = New pUsuario
    With Usuario
        .Codigo = txtCodigo
        .Nome = txtNome
        .Matricula = txtMatricula
        If .Gravar(.Codigo) Then
            Util.Mensagem "Usuário gravado com sucesso."
            Call cmdLimpar_Click
        Else
        End If
    End With
    Set Usuario = Nothing
    Screen.MousePointer = 0
    Exit Sub
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
        
    cabVisual.Exibir BDados, Me.Name, App.Path
    rodVISUAL.Exibir BDados, Me.Name, App.Major, App.Minor, App.Revision
        
    Set Usuario = New pUsuario
'    Set Lotacao = New pLotacao
    
'    Lotacao.PreencherCombo cboLotacao
'    Lotacao.ExibirEstrutura treeUsuarios, True
'    Usuario.ExibirUsuarios treeUsuarios
        
'    Set Lotacao = Nothing
    Set Usuario = Nothing
    Edita.LimpaCampos Me
    
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub treeUsuarios_Collapse(ByVal Node As MSComctlLib.Node)
    If treeUsuarios.Nodes(Node.Index).Image = "OPEN" Then treeUsuarios.Nodes(Node.Index).Image = "CLOSE"
End Sub

Private Sub treeUsuarios_Expand(ByVal Node As MSComctlLib.Node)
    If treeUsuarios.Nodes(Node.Index).Image = "CLOSE" Then treeUsuarios.Nodes(Node.Index).Image = "OPEN"
End Sub

Private Sub treeUsuarios_NodeClick(ByVal Node As MSComctlLib.Node)
If treeUsuarios.SelectedItem Is Nothing Then Exit Sub
    If Util.ParseString(Node.Tag, ":", 1) = "USUARIO" Then
        txtCodigo = Util.ParseString(Node.Tag, ":", 2)
        txtNome = Node.Text
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    On Error GoTo Trata
    Set Usuario = New pUsuario
        Usuario.Buscar txtCodigo
        txtNome = Usuario.Nome
        txtMatricula = Usuario.Matricula
    Set Usuario = Nothing
    DoEvents
    Exit Sub
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

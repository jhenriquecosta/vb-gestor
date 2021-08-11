VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form PDEF101 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PDEF101.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   5130
      Left            =   60
      TabIndex        =   10
      Top             =   1845
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   9049
      Altura          =   1905
      Caption         =   " Estrutura Definida"
      CorTexto        =   16777215
      CorFaixa        =   8421504
      CorFundo        =   15724527
      Ocultavel       =   0   'False
      Begin MSComctlLib.TreeView treeLotacao 
         Height          =   4680
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   8255
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   538
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   6
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1095
      Left            =   60
      TabIndex        =   9
      Top             =   690
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1931
      Altura          =   1905
      Caption         =   " Definição"
      CorTexto        =   16777215
      CorFaixa        =   8421504
      CorFundo        =   15724527
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDescricao 
         Height          =   285
         Left            =   345
         TabIndex        =   3
         Tag             =   "Nome"
         Top             =   720
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   503
         Caption         =   "Descrição"
         Text            =   ""
         CorFundo        =   15724527
         MaxLen          =   200
      End
      Begin VTOcx.txtVISUAL txtSigla 
         Height          =   285
         Left            =   2152
         TabIndex        =   1
         Tag             =   "Sigla"
         Top             =   375
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         Caption         =   "Sigla"
         Text            =   ""
         CorFundo        =   15724527
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Tag             =   "Codigo"
         Top             =   375
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         Caption         =   "Codigo"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   15724527
      End
      Begin VTOcx.cboVISUAL cboHierarquia 
         Height          =   315
         Left            =   4260
         TabIndex        =   2
         Top             =   360
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   556
         Caption         =   "Hierarquia"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorFundo        =   15724527
      End
   End
   Begin Cabecalho.rodVISUAL rodVisual 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   7035
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   926
      CorFundo        =   12632256
      CorFrente       =   4210752
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   405
         Left            =   4410
         TabIndex        =   4
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   5460
         TabIndex        =   5
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   405
         Left            =   6510
         TabIndex        =   6
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   4210752
         CorFrente       =   4210752
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   1138
      Formulario      =   "[Nome]"
      Descricao       =   "[Descricao]"
      Icone           =   "PDEF101.frx":0342
   End
End
Attribute VB_Name = "PDEF101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dragNode As Node, hilitNode As Node
Dim Lotacao As pLotacao

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
        
    Set Lotacao = New pLotacao
        txtCodigo = Lotacao.BuscaProximaLotacao
        Lotacao.ExibirEstrutura treeLotacao
    Set Lotacao = Nothing
    txtSigla.Enabled = True
    txtSigla.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
    Set Lotacao = New pLotacao
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    If txtSigla = cboHierarquia Or CInt(txtCodigo) < cboHierarquia.Coluna(1).Valor Then
        Erro "Erro na hierarquia informada."
        cboHierarquia.SetFocus
        Exit Sub
    End If
    
    With Lotacao
        .Codigo = txtCodigo
        .Hierarquia = IIf(cboHierarquia = "", 0, cboHierarquia.Coluna(1).Valor)
        .Sigla = txtSigla
        .Descricao = txtDescricao
        If .Gravar(.Codigo) Then
            Util.Avisa "Lotação gravada com sucesso."
            Call cmdLimpar_Click
            .PreencherCombo cboHierarquia
        Else
            Util.Erro "Problemas ao gravar."
        End If
    End With
    
    Set Lotacao = Nothing
    Exit Sub
Trata:
    Erro ERR.Description
    Set Lotacao = Nothing
End Sub

Private Sub Form_Load()
    cabVisual.Exibir BDados, Me.Name, App.Path
    rodVisual.Exibir BDados, Me.Name

    Set Lotacao = New pLotacao
        Lotacao.PreencherCombo cboHierarquia
        Lotacao.ExibirEstrutura treeLotacao
    Set Lotacao = Nothing
End Sub

Private Sub treeLotacao_NodeClick(ByVal Node As MSComctlLib.Node)

    Edita.LimpaCampos Me
    txtCodigo = Util.ParseString(Node.Tag, ":", 3)
    txtSigla = Trim(Util.ParseString(Node.Text, "-", 1))
    cboHierarquia.SetarLinha Util.ParseString(Node.Tag, ":", 4), 1
    txtDescricao = Trim(Util.ParseString(Node.Text, "-", 2))
    txtSigla.Enabled = False
End Sub

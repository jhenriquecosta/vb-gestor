VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TDEC401.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   5
      Top             =   7305
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdPesquisar 
         Height          =   375
         Left            =   5460
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   90
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Declaracões"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   7230
         TabIndex        =   3
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8520
         TabIndex        =   4
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1560
      Left            =   30
      TabIndex        =   6
      Top             =   690
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   2752
      Altura          =   1905
      Caption         =   " Sujeito Passivo"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   2940
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   330
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6450
         TabIndex        =   9
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         Caption         =   "Status Declaracão"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   690
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         Caption         =   "Período"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   330
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   330
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Caption         =   "Tipo Declaracão"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL grdDec 
      Height          =   5025
      Left            =   30
      TabIndex        =   8
      Top             =   2280
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   8864
      CorBorda        =   16711680
      Caption         =   "Declaracões"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1138
      Icone           =   "TDEC401.frx":0342
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TDEC401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declaracao As VsTFuncoes.cDeclaracao
Private GerarIM As Boolean
Private AliqISSQN As Double, ISSQNFixo As Double
Private AliqISSST As Double, ISSSTFixo As Double

Private TotalImpostoST As Double
Private TotalBaseST As Double
Private TotalImpostoDevidoSaida As Double
Private TotalImpostoRetidoSaida As Double
Private TotalBaseSaida As Double
Private TotalICMSSujeito As Double
Private DeduzValores As Boolean
Private ContribuinteEndereco As String
Private ContribuinteAtividade As String
Dim Notas() As New NotaFiscal
Private Sub FormataRegistro(ByRef Inscricao As Object)
    Select Case Len(Inscricao.Text)
        Case 10
            Inscricao.Text = Imposto.FormataInscricao(Inscricao.Text, InscContrib)
        Case 11
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cpf)
        Case 14
            Inscricao.Text = Edita.FormataTexto(Inscricao, Cgc)
    End Select
End Sub

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub



Private Sub cmdPesquisar_Click()
    Dim Status As Integer
    Dim Tipo As Integer
    Set Declaracao = New VsTFuncoes.cDeclaracao
    If IsEmpty(cboStatus.Coluna(1).Valor) Then
        Status = 0
    Else
        Status = CInt(cboStatus.Coluna(1).Valor)
    End If
    If IsEmpty(cboTipo.Coluna(1).Valor) Then
        Tipo = 0
    Else
        Tipo = CInt(cboTipo.Coluna(1).Valor)
    End If
    Declaracao.CarregaGrid grdDec, txtIM, txtPeriodo, CInt(Tipo), CInt(Status)
    If grdDec.ListItems.Count = 0 Then Avisa "Nenhum registro encontrado."
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    txtIM.SetFocus
End Sub

Private Sub Form_Activate()
   ' cabVISUAL1.Exibir Bdados, Me.Tag, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Tag
    
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
    cboStatus.PreencherGeral Bdados, "STATUS DECLARACAO"
End Sub

Private Sub txtIM_LostFocus()
    
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        End If
    End If
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
End Sub


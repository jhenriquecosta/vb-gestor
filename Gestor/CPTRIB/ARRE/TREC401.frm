VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREC401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TREC401.frx":0000
   ScaleHeight     =   5730
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TREC401.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1065
      Left            =   90
      TabIndex        =   11
      Top             =   690
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   1879
      Altura          =   1905
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAgencia 
         Height          =   300
         Left            =   3330
         TabIndex        =   2
         Top             =   690
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         Caption         =   "Agencia  Arrecadação"
         Text            =   ""
         MaxLen          =   8
      End
      Begin VTOcx.cboVISUAL cboBanco 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   330
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   556
         Caption         =   "Agente Arrecadador"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtData 
         Height          =   300
         Left            =   390
         TabIndex        =   1
         Top             =   690
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   529
         Caption         =   "Data Arrecadação"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MaxLen          =   12
      End
      Begin VTOcx.cboVISUAL cboOcorrencia 
         Height          =   315
         Left            =   6360
         TabIndex        =   3
         Top             =   690
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   556
         Caption         =   "Ocorrência"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   930
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstArq 
      Height          =   3390
      Left            =   90
      TabIndex        =   8
      Top             =   1815
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   5980
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1138
      Icone           =   "TREC401.frx":2465
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   7845
      TabIndex        =   7
      Top             =   5325
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Sair"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdImprime 
      Height          =   375
      Left            =   6675
      TabIndex        =   6
      Top             =   5325
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdBusca 
      Height          =   375
      Left            =   5505
      TabIndex        =   4
      Top             =   5325
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmLimpar 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   5310
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TREC401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub PreencheRodape(Lista As Object)
    If Lista.ListItems.Count > 0 Then
        Lista.Mensagem = "Total rejeitado : R$" & Format(Lista.Colunas(2).Soma, Const_Monetario) & _
        Space(15) & " Legenda: JP = Já Pago  x  NE = Não Encontrado"
    Else
        Lista.Mensagem = " Legenda: JP = Já Pago  x  NE = Não Encontrado"
    End If
End Sub

Private Sub cmdBusca_Click()
    Dim Arq As New Arquivo
    If Not Arq.PesquisaDocumentos(lstArq, CStr(cboBanco.Coluna(0).Valor), txtAgencia, txtData, CStr(cboOcorrencia.Coluna(1).Valor)) Then Avisa "Nenhum registro encontrado."
    PreencheRodape lstArq
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdImprime_Click()
    Dim Aux As Byte
    Dim Formula As String
    Dim Paginas As Integer
    Dim SelecaoRpt As String
    
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    lstArq.ListItems.Clear
    PreencheRodape lstArq
    cboBanco.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboBanco.Preencher Bdados, "Select tar_cod_agente,tar_nome_agente from tab_agente_arrecadador", 1
    cboOcorrencia.PreencherGeral Bdados, "REJEICAO DOCUMENTO"
    PreencheRodape lstArq
End Sub

VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TDEC102.frx":0000
   ScaleHeight     =   6660
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   6135
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7710
         TabIndex        =   5
         Top             =   90
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5535
         TabIndex        =   4
         Top             =   75
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         Caption         =   "&Reabrir Declaracão"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8730
         TabIndex        =   6
         Top             =   75
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1138
      Icone           =   "TDEC102.frx":0342
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1080
      Left            =   0
      TabIndex        =   9
      Top             =   660
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Sujeito Passivo"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
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
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   30
         TabIndex        =   0
         Top             =   330
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Top             =   330
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   556
         Caption         =   "Razão Social"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Tag             =   "Periodo"
         Top             =   690
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         Caption         =   "Período"
         Text            =   ""
         TipoLetras      =   0
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisar 
         Height          =   375
         Left            =   8130
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Declaracões"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.grdVISUAL grdDec 
      Height          =   4245
      Left            =   0
      TabIndex        =   11
      Top             =   1770
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   7488
      CorBorda        =   16711680
      Caption         =   "Declaracões"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
      OcultarRodape   =   -1  'True
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "Geral"
      Visible         =   0   'False
      Begin VB.Menu mnuReabrir 
         Caption         =   "Reabrir"
      End
   End
End
Attribute VB_Name = "TDEC102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdPesquisar_Click()
'    If Not CriticaCampos(Me) Then Exit Sub
    Declaracao.CarregaGrid grdDec, txtIM, txtPeriodo, CInt(cboTipo.Coluna(1).Valor), decNaoAberta
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim i As Double
    If grdDec.ListItems.Count > 0 Then
        Campos = "TDC_STATUS"
        Valores = Bdados.PreparaValor(decAberta)
        For i = 1 To grdDec.ListItems.Count
            Call Bdados.AtualizaDados("TAB_DECLARACAO_CONTRIBUINTE", Valores, Campos, "TDC_NUM_DECLARACAO =" & grdDec.ListItems(i).SubItems(6))
        Next
        Avisa "Declaracão(ões) reaberta(s)."
        cmdPesquisar_Click
    Else
        Avisa "Nenhuma declaracão selecionada."
        txtIM.SetFocus
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
End Sub

Private Sub grdDec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuPopUp Me, grdDec, Button, mnuGeral, mnuReabrir, "Reabrir Declaracão"
End Sub

Private Sub mnuReabrir_Click()
    Dim Valores As String
    Dim Campos As String
    If grdDec.ListItems.Count > 0 Then
        Valores = Bdados.PreparaValor(decAberta)
        Campos = "TDC_STATUS"
        Call Bdados.AtualizaDados("TAB_DECLARACAO_CONTRIBUINTE", Valores, Campos, "TDC_NUM_DECLARACAO =" & grdDec.SelectedItem.SubItems(6))
        Avisa "Declaracão reaberta."
        cmdPesquisar_Click
    Else
        Avisa "Nenhuma declaracão selecionada."
        txtIM.SetFocus
    End If
End Sub

Private Sub txtIM_LostFocus()
    If Trim(txtIM) = "" Then
        txtRazao = ""
        Exit Sub
    End If
    If Not BuscarContribuinte(txtIM, txtRazao) Then
        Avisa "Contribuinte não encontrado."
        txtIM = "": txtRazao = ""
        txtIM.SetFocus
    End If
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
End Sub

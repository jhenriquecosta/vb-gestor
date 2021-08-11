VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Object = "{EA761AE1-8FDE-4340-8E6D-420E99B0C363}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdLivros 
      Height          =   3780
      Left            =   75
      TabIndex        =   7
      Top             =   1590
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4339
      Caption         =   "Livros"
      CorTitulo       =   32768
      CorCaption      =   16777215
   End
   Begin VTOcx.fraVISUAL fraRegra 
      Height          =   825
      Left            =   75
      TabIndex        =   6
      Top             =   690
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   1455
      Altura          =   1905
      Caption         =   " Definição de regras"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtQtdFolha 
         Height          =   300
         Left            =   2175
         TabIndex        =   1
         Tag             =   "Folhas"
         Top             =   390
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
         Caption         =   "Folhas por livro"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Tag             =   "Período"
         Top             =   390
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         Caption         =   "Período"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         MaxLen          =   4
         MinLen          =   4
      End
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   5385
      TabIndex        =   4
      Top             =   5430
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdRegistrar 
      Height          =   375
      Left            =   4125
      TabIndex        =   2
      Top             =   5430
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "R&egistrar"
      Acao            =   1
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdNovo 
      Height          =   375
      Left            =   2955
      TabIndex        =   3
      Top             =   5430
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   1138
      Icone           =   "TCIU105.frx":0000
   End
End
Attribute VB_Name = "TCIU105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
End Sub

Private Sub cmdRegistrar_Click()
Dim Sql As String
Dim sCampos As String
Dim sValores As String

    Sql = "SELECT TAM_LIVRO FROM TAB_AFORAMENTO_MANUTENCAO WHERE " & _
        " TAM_PERIODO ='" & txtPeriodo & "'"
    If Not Bdados.AbreTabela(Sql) Then
        sValores = Bdados.PreparaValor(txtPeriodo, 0, 0, 0, 0, txtQtdFolha, slaLivroAberto, Aplicacoes.Usuario)
        sCampos = "TAM_PERIODO,TAM_ORDEM,TAM_FICHA,TAM_LIVRO,TAM_FOLHA,TAM_FOLHA_TOTAL,TAM_STATUS,TAM_TUS_COD_USUARIO"
        If Bdados.InsereDados("TAB_AFORAMENTO_MANUTENCAO", sValores, sCampos) Then
            Util.Mensagem "Regra registrada."
            Edita.LimpaCampos Me
            Call Busca_Livros
        Else
            Util.Erro "Não foi possível criar a regra."
        End If
    Else
        Util.Informa "Já existe uma regra para os livros de aforamento para o ano de " & txtPeriodo & "."
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Call Busca_Livros
End Sub

Private Sub Busca_Livros()
Dim Sql As String

    Sql = "SELECT TAM_PERIODO AS PERÍODO, TAM_LIVRO AS LIVRO, TAM_FOLHA AS FOLHA, TAM_FOLHA_TOTAL AS DEFINIDO, TAM_ORDEM AS ORDEM, " & _
        " TAM_FICHA AS FICHA, TAM_STATUS AS STATUS FROM TAB_AFORAMENTO_MANUTENCAO ORDER BY PERÍODO DESC"
    
    grdLivros.Preencher Bdados, Sql
    grdLivros.Mensagem = "Status: 1 - Aberto      0 - Fechado"
    
    
    Sql = "SELECT MAX(TAM_PERIODO) " & _
    " FROM TAB_AFORAMENTO_MANUTENCAO WHERE TAM_STATUS = " & slaLivroAberto
    
    If Bdados.AbreTabela(Sql) Then
        fraRegra.Enabled = False
        cmdRegistrar.Enabled = False
        cmdNovo.Enabled = False
    End If
    
End Sub

Private Sub txtPeriodo_Validate(Cancel As Boolean)
    If (Val(txtPeriodo) < (Year(Date) - 5)) Or (Val(txtPeriodo) > Year(Date)) Then txtPeriodo = Year(Date)
    'If (Val(txtPeriodo) < (Year(Date) - 5)) Then txtPeriodo = Year(Date)
End Sub

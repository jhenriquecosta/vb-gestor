VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TOBR109 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TOBR109.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1138
      Icone           =   "TOBR109.frx":0342
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1140
      Left            =   30
      TabIndex        =   5
      Top             =   690
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   2011
      Altura          =   1905
      Caption         =   " Sujeito Passivo"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8370
         TabIndex        =   3
         Top             =   690
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdPesquisar 
         Height          =   375
         Left            =   6630
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   690
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   2940
         TabIndex        =   8
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
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   690
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         Caption         =   "Exercício"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
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
         Left            =   210
         TabIndex        =   0
         Top             =   330
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL grdDec 
      Height          =   5025
      Left            =   30
      TabIndex        =   7
      Top             =   1860
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   8864
      Caption         =   "Contribuintes"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TOBR109"
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
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Obrig As New Obrigacao
    Dim i As Byte
    Dim CodImposto As String
    
    If Trim(txtPeriodo) = "" Then
        Avisa "Período inválido."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    Sql = "SELECT TCE_TCI_IM as Contribuinte,TCE_VALOR_ESTIMADO as Valor FROM TAB_CONTRIBUINTE_ESTIMADO WHERE TCE_EXERCICIO = " & txtPeriodo
    If Trim(txtIM) <> "" Then
        Sql = Sql & " AND TCE_TCI_IM = '" & txtIM & "'"
    End If
    Screen.MousePointer = 11
    CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNEST))
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            Obrig.CriaObrigacao CodImposto, _
                    "01" & txtPeriodo, "12" & txtPeriodo, rs!Contribuinte, rs!Valor, etsCreditoOriginalAberto
            rs.MoveNext
        Loop While Not rs.EOF
    End If
    grdDec.Preencher Bdados, Sql
    Avisa "Processo finalizado!"
    Screen.MousePointer = 0
    If grdDec.ListItems.Count = 0 Then Avisa "Nenhum registro encontrado."
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    txtIM.SetFocus
End Sub

Private Sub txtIm_LostFocus()
    If Trim$(txtIM) <> "" Then
        txtIM = BuscaContribuinte(txtIM, txtRazao)
    End If
End Sub


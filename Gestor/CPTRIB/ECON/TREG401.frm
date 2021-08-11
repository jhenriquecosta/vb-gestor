VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREG401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TREG401.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   10
      Top             =   6465
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9045
         TabIndex        =   13
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Sair"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7830
         TabIndex        =   12
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6615
         TabIndex        =   11
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TREG401.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1138
      Icone           =   "TREG401.frx":2465
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   5820
      Left            =   15
      TabIndex        =   2
      Top             =   660
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   10266
      Altura          =   1905
      Caption         =   " Consulta"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDataFinalConsulta 
         Height          =   285
         Left            =   6930
         TabIndex        =   9
         Top             =   465
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         Caption         =   "Exercicio Final"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIMConsulta 
         Height          =   315
         Left            =   1395
         TabIndex        =   8
         Top             =   435
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataInicialConsulta 
         Height          =   285
         Left            =   4365
         TabIndex        =   7
         Top             =   465
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         Caption         =   "Exercicio Inicial"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   3990
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   435
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cboVISUAL cboProcedimentoConsulta 
         Height          =   315
         Left            =   3870
         TabIndex        =   5
         Top             =   795
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         Caption         =   "Procedimento"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtProcessoConsulta 
         Height          =   285
         Left            =   7380
         TabIndex        =   4
         Top             =   795
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   503
         Caption         =   "Processo"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.grdVISUAL grdVISUAL1 
         Height          =   4575
         Left            =   60
         TabIndex        =   3
         Top             =   1230
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   8070
      End
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TREG401"
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
        
  
        
    Sql = "SELECT TCE_TCI_IM AS Inscrição,"
    Sql = Sql & " TCI_NOME AS Nome ,"
    Sql = Sql & " TCE_EXERCICIO as Exercicio,"
    Sql = Sql & " TCE_BASE_CALCULO_ANUAL_UFM as Valor_Anual_UFM,"
    Sql = Sql & " TCE_VALOR_MENSAL as Valor_Mensal,"
    Sql = Sql & " TCE_BASE_CALCULO_ANUAL as Valor_Anual,"
    Sql = Sql & " TGE_NOME as Procedimento,"
    Sql = Sql & " TCE_PROCESSO as Processo ,"
    Sql = Sql & " TGE_CODIGO"
    Sql = Sql & " FROM TAB_CONTRIBUINTE_ESTIMADO,VIS_PROCEDIMENTO,TAB_CONTRIBUINTE"
    Sql = Sql & " WHERE TGE_CODIGO = TCE_STATUS "
    Sql = Sql & " AND TCE_TCI_IM = TCI_IM "
    
    If Trim(txtIMConsulta) <> "" Then
        Sql = Sql & " AND TCE_TCI_IM = '" & txtIMConsulta & "'"
    End If
    If Trim(txtDataInicialConsulta) <> "" Then
        Sql = Sql & " AND TCE_EXERCICIO >= " & txtDataInicialConsulta
    End If
    If txtDataFinalConsulta <> "" Then
        Sql = Sql & " AND TCE_EXERCICIO <= " & txtDataFinalConsulta
    End If
    
    If cboProcedimentoConsulta.ListIndex <> -1 Then
        Sql = Sql & " and TCE_STATUS  = '" & cboProcedimentoConsulta.Coluna(1).VALOR & "'"
    End If
    
    If txtProcessoConsulta <> "" Then
        Sql = Sql & " and TCE_PROCESSO  = '" & txtProcessoConsulta & "'"
    End If
    
    grdVISUAL1.Preencher Bdados, Sql, 1000, 4000, 1000, 2000, 2000, 2000, 2000, 1000, 1000
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    txtIM.SetFocus
End Sub

Private Sub txtIM_LostFocus()
    If Trim$(txtIM) <> "" Then
        txtIM = BuscaContribuinte(txtIM, txtRazao)
    End If
End Sub


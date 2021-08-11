VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREG401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TREG401"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Icone           =   "TREG401.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   6630
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   8100
         TabIndex        =   5
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4785
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   690
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Salvar Declaracão"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9195
         TabIndex        =   6
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   5955
      Left            =   15
      TabIndex        =   10
      Top             =   660
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   10504
      Altura          =   1905
      Caption         =   " Consulta"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin ActiveTabs.SSActiveTabs TabDados 
         Height          =   4800
         Left            =   90
         TabIndex        =   13
         Top             =   1065
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   8467
         _Version        =   131082
         TabCount        =   2
         TabOrientation  =   2
         BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TagVariant      =   ""
         Tabs            =   "TREG401.frx":282A
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   4410
            Left            =   30
            TabIndex        =   15
            Top             =   30
            Width           =   10020
            _ExtentX        =   17674
            _ExtentY        =   7779
            _Version        =   131082
            TabGuid         =   "TREG401.frx":28A9
            Begin VTOcx.grdVISUAL grdVISUAL1 
               Height          =   4290
               Left            =   75
               TabIndex        =   18
               Top             =   90
               Width           =   9870
               _ExtentX        =   17410
               _ExtentY        =   7567
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
            Height          =   4410
            Left            =   30
            TabIndex        =   14
            Top             =   30
            Width           =   10020
            _ExtentX        =   17674
            _ExtentY        =   7779
            _Version        =   131082
            TabGuid         =   "TREG401.frx":28D1
            Begin VTOcx.grdVISUAL grdDados 
               Height          =   2325
               Left            =   120
               TabIndex        =   16
               Top             =   2070
               Width           =   9765
               _ExtentX        =   17224
               _ExtentY        =   4101
            End
            Begin VTOcx.grdVISUAL grdContribuintes 
               Height          =   1950
               Left            =   105
               TabIndex        =   17
               Top             =   90
               Width           =   9810
               _ExtentX        =   17304
               _ExtentY        =   3440
               Caption         =   "Contribuintes"
            End
         End
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   8925
         TabIndex        =   12
         Top             =   645
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtProcessoConsulta 
         Height          =   285
         Left            =   6765
         TabIndex        =   4
         Top             =   690
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   503
         Caption         =   "Processo"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboProcedimentoConsulta 
         Height          =   315
         Left            =   3255
         TabIndex        =   3
         Top             =   690
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         Caption         =   "Procedimento"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   3375
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
      Begin VTOcx.txtVISUAL txtDataInicialConsulta 
         Height          =   285
         Left            =   3750
         TabIndex        =   1
         Top             =   360
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         Caption         =   "Exercicio Inicial"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIMConsulta 
         Height          =   315
         Left            =   780
         TabIndex        =   0
         Top             =   330
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataFinalConsulta 
         Height          =   285
         Left            =   6315
         TabIndex        =   2
         Top             =   360
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         Caption         =   "Exercicio Final"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         RetirarMascara  =   0   'False
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
Dim Modalidade As Integer
Dim String_Taxas As String
Dim Total_Taxas As Double
Dim atividade As New VsTEcon.atividade

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

Private Sub cmdBuscar_Click()
    Dim SQL As String
    
    SQL = "SELECT distinct(TCE_TCI_IM) AS Inscrição,"
    SQL = SQL & " TCI_NOME AS Nome"
    SQL = SQL & " FROM TAB_CONTRIBUINTE_ESTIMADO,VIS_PROCEDIMENTO,TAB_CONTRIBUINTE"
    SQL = SQL & " WHERE TGE_CODIGO = TCE_STATUS "
    SQL = SQL & " AND TCE_TCI_IM = TCI_IM "
    If Trim(txtIMConsulta) <> "" Then
        SQL = SQL & " AND TCE_TCI_IM = '" & txtIMConsulta & "'"
    End If
    If Trim(txtDataInicialConsulta) <> "" Then
        SQL = SQL & " AND TCE_EXERCICIO >= " & txtDataInicialConsulta
    End If
    If txtDataFinalConsulta <> "" Then
        SQL = SQL & " AND TCE_EXERCICIO <= " & txtDataFinalConsulta
    End If
    
    If cboProcedimentoConsulta.ListIndex <> -1 Then
        SQL = SQL & " and TCE_STATUS  = '" & cboProcedimentoConsulta.Coluna(1).VALOR & "'"
    End If
    
    If txtProcessoConsulta <> "" Then
        SQL = SQL & " and TCE_PROCESSO  = '" & txtProcessoConsulta & "'"
    End If
    
    grdContribuintes.Preencher Bdados, SQL
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIMConsulta
End Sub

Private Sub cmLimpar_Click()
    LimpaCampos Me
    txtIMConsulta.SetFocus
End Sub

Private Sub Form_Load()
    Dim SQL As String
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Tag
    Set Imposto = New VsTFuncoes.VSImposto
    
    cboProcedimentoConsulta.PreencherGeral Bdados, "PROCEDIMENTO ESTIMATIVA"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub fraVISUAL2_mudancaStatus()

End Sub


Private Sub grdContribuintes_DblClick()
  Dim SQL As String
        
    If grdContribuintes.ListItems.Count >= 1 Then
            
        SQL = "SELECT TCE_TCI_IM AS Inscrição,"
        SQL = SQL & " TCI_NOME AS Nome ,"
        SQL = SQL & " TCE_EXERCICIO as Exercicio,"
        SQL = SQL & " TCE_BASE_CALCULO_ANUAL_UFM as Valor_Anual_UFM,"
        SQL = SQL & " TCE_VALOR_MENSAL as Valor_Mensal,"
        SQL = SQL & " TCE_BASE_CALCULO_ANUAL as Valor_Anual,"
        SQL = SQL & " TGE_NOME as Procedimento,"
        SQL = SQL & " TCE_PROCESSO as Processo ,"
        SQL = SQL & " TGE_CODIGO,"
        SQL = SQL & " TCE_DATA_PROCESSO AS Data_Processo,"
        SQL = SQL & " TCE_DATA_PROCEDIMENTO AS Data_Procedimento"
        SQL = SQL & " FROM TAB_CONTRIBUINTE_ESTIMADO,VIS_PROCEDIMENTO,TAB_CONTRIBUINTE"
        SQL = SQL & " WHERE TGE_CODIGO = TCE_STATUS "
        SQL = SQL & " AND TCE_TCI_IM = TCI_IM "
        SQL = SQL & " AND TCE_TCI_IM = '" & grdContribuintes.SelectedItem & "'"
        If Trim(txtDataInicialConsulta) <> "" Then
            SQL = SQL & " AND TCE_EXERCICIO >= " & txtDataInicialConsulta
        End If
        grdDados.Preencher Bdados, SQL, 1000, 4000, 1000, 2000, 2000, 2000, 2000, 2000, 0, 2000, 2000
        
        'PEGO O HISTÓRICO...
        SQL = "SELECT TCE_TCI_IM AS Inscrição,"
        SQL = SQL & " TCI_NOME AS Nome ,"
        SQL = SQL & " TCE_EXERCICIO as Exercicio,"
        SQL = SQL & " TCE_BASE_CALCULO_ANUAL_UFM as Valor_Anual_UFM,"
        SQL = SQL & " TCE_VALOR_MENSAL as Valor_Mensal,"
        SQL = SQL & " TCE_BASE_CALCULO_ANUAL as Valor_Anual,"
        SQL = SQL & " TGE_NOME as Procedimento,"
        SQL = SQL & " TCE_PROCESSO as Processo ,"
        SQL = SQL & " TCE_DATA_PROCESSO AS Data_Processo,"
        SQL = SQL & " TCE_DATA AS Data,"
        SQL = SQL & " TCE_USUARIO AS Usuário,"
        SQL = SQL & " TCE_Motivo AS Motivo"
        SQL = SQL & " FROM TAB_CONTRIBUINTE_ESTIMADO_HIST,VIS_PROCEDIMENTO,TAB_CONTRIBUINTE"
        SQL = SQL & " WHERE TGE_CODIGO = TCE_STATUS "
        SQL = SQL & " AND TCE_TCI_IM = TCI_IM "
        SQL = SQL & " AND TCE_TCI_IM = '" & grdContribuintes.SelectedItem & "'"
        grdVISUAL1.Preencher Bdados, SQL
    End If
    
End Sub

Private Sub grdDados_DblClick()
    If grdDados.ListItems.Count >= 1 Then
        Dim SQL As String
        
        'PEGO O HISTÓRICO...
        SQL = "SELECT TCE_TCI_IM AS Inscrição,"
        SQL = SQL & " TCI_NOME AS Nome ,"
        SQL = SQL & " TCE_EXERCICIO as Exercicio,"
        SQL = SQL & " TCE_BASE_CALCULO_ANUAL_UFM as Valor_Anual_UFM,"
        SQL = SQL & " TCE_VALOR_MENSAL as Valor_Mensal,"
        SQL = SQL & " TCE_BASE_CALCULO_ANUAL as Valor_Anual,"
        SQL = SQL & " TGE_NOME as Procedimento,"
        SQL = SQL & " TCE_PROCESSO as Processo ,"
        SQL = SQL & " TCE_DATA_PROCESSO AS Data_Processo,"
        SQL = SQL & " TCE_DATA AS Data,"
        SQL = SQL & " TCE_USUARIO AS Usuário,"
        SQL = SQL & " TCE_Motivo AS Motivo"
        SQL = SQL & " FROM TAB_CONTRIBUINTE_ESTIMADO_HIST,VIS_PROCEDIMENTO,TAB_CONTRIBUINTE"
        SQL = SQL & " WHERE TGE_CODIGO = TCE_STATUS "
        SQL = SQL & " AND TCE_TCI_IM = TCI_IM "
        SQL = SQL & " AND TCE_TCI_IM = '" & grdDados.SelectedItem & "'"
        SQL = SQL & " AND TCE_EXERCICIO   = '" & grdDados.SelectedItem.SubItems(2) & "'"
        grdVISUAL1.Preencher Bdados, SQL
        TabDados.Tabs(2).Selected = True
    End If
    
End Sub

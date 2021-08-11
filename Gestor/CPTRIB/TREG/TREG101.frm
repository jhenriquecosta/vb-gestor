VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREG101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDEC108"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Icone           =   "TREG101.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   13
      Top             =   2505
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   90
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   7770
         TabIndex        =   8
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4770
         TabIndex        =   12
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
         Left            =   9060
         TabIndex        =   9
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1860
      Left            =   15
      TabIndex        =   14
      Top             =   645
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   3281
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDataProcesso 
         Height          =   285
         Left            =   7350
         TabIndex        =   3
         Top             =   690
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         Caption         =   "Data do Processo"
         Text            =   ""
         Formato         =   0
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtValorMensal 
         Height          =   300
         Left            =   6630
         TabIndex        =   5
         Top             =   1050
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   529
         Caption         =   "Valor Estimado Mensal R$"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.txtVISUAL txtValorAnual 
         Height          =   300
         Left            =   6735
         TabIndex        =   6
         Top             =   1395
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   "Valor Estimado Anual R$"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.txtVISUAL txtProcesso 
         Height          =   285
         Left            =   5145
         TabIndex        =   2
         Top             =   690
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   503
         Caption         =   "Processo"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   285
         Left            =   900
         TabIndex        =   4
         Top             =   1065
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   503
         Caption         =   "Base  Estimada Anualmente em  UFM"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboProcedimento 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   690
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         Caption         =   "Procedimento"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   3090
         TabIndex        =   15
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
      Begin VTOcx.txtVISUAL txtDataInicial 
         Height          =   285
         Left            =   3330
         TabIndex        =   1
         Top             =   690
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         Caption         =   "Exercicio"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   4
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3450
         TabIndex        =   10
         Top             =   330
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   45
         TabIndex        =   0
         Top             =   330
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   4050
      Left            =   15
      TabIndex        =   17
      Top             =   675
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   7144
      Altura          =   1905
      Caption         =   " Consulta"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtVISUAL1 
         Height          =   285
         Left            =   6765
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   8970
         TabIndex        =   18
         Top             =   570
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
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
Attribute VB_Name = "TREG101"
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

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub

Private Sub cmdFinaliza_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Codigo As String
    Dim Conta As New ContaCorrente
    
    '1 = INCLUSÃO
    '2 = CANCELAMENTO
    '3 = ALTERAÇÃO
    '4 = EXCLUSÃO
    Campos = "TCE_TCI_IM,TCE_EXERCICIO,TCE_BASE_CALCULO_ANUAL_UFM,TCE_STATUS,TCE_DATA_PROCEDIMENTO,TCE_PROCESSO,TCE_VALOR_MENSAL,TCE_BASE_CALCULO_ANUAL,TCE_DATA_PROCESSO,TCE_USUARIO"
    Valores = Bdados.PreparaValor(txtIM, txtDataInicial, Bdados.Converte(txtValor, TCMonetario), _
                    cboProcedimento.Coluna(1).Valor, Bdados.Converte(Date, TCDataHora), txtProcesso, Bdados.Converte(txtValorMensal, TCMonetario), Bdados.Converte(txtValorAnual, TCMonetario), Bdados.Converte(txtDataProcesso, TCDataHora), Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto))
    If Bdados.GravaDados("TAB_CONTRIBUINTE_ESTIMADO", Valores, Campos, "TCE_TCI_IM ='" & txtIM & "' and TCE_EXERCICIO = '" & txtDataInicial & "'") Then
        'GRAVO O PRIMEIRO REGISTRO NO HISTÓRICO...
        Campos = "TCE_COD_MUDANCA,TCE_TCI_IM,TCE_EXERCICIO,TCE_BASE_CALCULO_ANUAL_UFM,TCE_STATUS,TCE_DATA,TCE_PROCESSO,TCE_VALOR_MENSAL,TCE_BASE_CALCULO_ANUAL,TCE_DATA_PROCESSO,TCE_USUARIO"
        Codigo = Conta.GeraCodPagamento(90)
        Valores = Bdados.PreparaValor(Bdados.Converte(Codigo, tctexto), Bdados.Converte(txtIM, tctexto), Bdados.Converte(txtDataInicial, TCDataHora), Bdados.Converte(txtValor, TCMonetario), _
                    cboProcedimento.Coluna(1).Valor, Bdados.Converte(Date, TCDataHora), txtProcesso, Bdados.Converte(txtValorMensal, TCMonetario), Bdados.Converte(txtValorAnual, TCMonetario), Bdados.Converte(txtDataProcesso, TCDataHora), Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto))
        
        Bdados.GravaDados "TAB_CONTRIBUINTE_ESTIMADO_HIST", Valores, Campos, "TCE_COD_MUDANCA = '" & Codigo & "'"
        Avisa "Registro gravado com sucesso."
        cmLimpar_Click
    Else
        Avisa "Erro ao gravar registro."
    End If
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmLimpar_Click()
    LimpaCampos Me
    txtIM.SetFocus
End Sub

Private Sub Form_Load()
    Dim SQL As String
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    'rodVISUAL1.Exibir Bdados, Me.Tag
    Set Imposto = New VsTFuncoes.VSImposto

    cboProcedimento.PreencherGeral Bdados, "PROCEDIMENTO ESTIMATIVA"
    cboProcedimento.SetarLinha 1, 1
    cboProcedimento.Visible = False
    cboProcedimento.Enabled = False
End Sub
Private Sub Calcula()
    Dim SQL                          As String
    Dim Rs                           As VSRecordset
    Dim Aliquota                     As Double
    Dim Base                         As Double
'
   If txtValor = "" Then Exit Sub
    
    
   'Pego a aliquota para o ano...
      Aliquota = Aliquota_Atividade(txtIM)
      Base = txtValor / 12
      Base = Base * CCur(TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ","))
      txtValorMensal = Base + (Aliquota * Base / 100)
      txtValorAnual = (Base + (Aliquota * Base / 100)) * 12
End Sub
Public Function Aliquota_Atividade(Im As String) As Double
    Dim SQL                           As String
    Dim Rs                            As VSRecordset
    Dim RsAliquota                    As VSRecordset
    
    SQL = "Select tci_tae_cae from tab_contribuinte where tci_im = '" & Im & "'"
    If Bdados.AbreTabela(SQL, Rs) Then
        If Not IsNull(Rs.Fields("tci_tae_cae")) Then
            SQL = "Select * from tab_atividade_economica where tae_cae = '" & Rs.Fields("tci_tae_cae") & "'"
        Else
            Avisa "Atividade não definida no cadastro econômico."
            Exit Function
        End If
        If Bdados.AbreTabela(SQL, RsAliquota) Then
            Aliquota_Atividade = RsAliquota.Fields("TAE_ALIQUOTA_PJ")
        Else
            Avisa "Aliquota não definida no cadastro econômico."
        End If
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub txtIM_LostFocus()
    If Trim$(txtIM) <> "" Then
        txtIM = BuscaContribuinte(txtIM, txtRazao)
    End If
End Sub


Private Sub txtValor_Change()
    Calcula
End Sub

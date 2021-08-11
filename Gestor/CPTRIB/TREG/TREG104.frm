VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TREG104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TREG104"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Icone           =   "TREG104.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   19
      Top             =   6390
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6990
         TabIndex        =   15
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   5895
         TabIndex        =   14
         Top             =   90
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   8100
         TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   17
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
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1065
      Left            =   45
      TabIndex        =   20
      Top             =   4620
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1879
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDataProcesso 
         Height          =   285
         Left            =   7335
         TabIndex        =   10
         Top             =   690
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   503
         Caption         =   "Data do Processo"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   0
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtProcesso 
         Height          =   285
         Left            =   5220
         TabIndex        =   9
         Top             =   690
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   503
         Caption         =   "Processo"
         Text            =   ""
         Enabled         =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboProcedimento 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   690
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         Caption         =   "Procedimento"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataInicial 
         Height          =   285
         Left            =   3480
         TabIndex        =   8
         Top             =   690
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         Caption         =   "Exercicio"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         MaxLen          =   4
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   330
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   315
         Left            =   45
         TabIndex        =   5
         Top             =   330
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   5730
      Left            =   15
      TabIndex        =   22
      Top             =   660
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   10107
      Altura          =   1905
      Caption         =   " Consulta"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   285
         Left            =   1740
         TabIndex        =   11
         Top             =   5085
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   503
         Caption         =   "Base Estimada Anualmente em  UFM"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   2
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtValorAnual 
         Height          =   300
         Left            =   6750
         TabIndex        =   13
         Top             =   5385
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         Caption         =   "Valor Estimado Anual R$"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.txtVISUAL txtValorMensal 
         Height          =   300
         Left            =   6645
         TabIndex        =   12
         Top             =   5055
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   529
         Caption         =   "Valor Estimado Mensal R$"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.grdVISUAL grdVISUAL1 
         Height          =   2880
         Left            =   60
         TabIndex        =   25
         Top             =   1065
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   5080
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
         TabIndex        =   24
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
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   8970
         TabIndex        =   23
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
Attribute VB_Name = "TREG104"
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
    
    grdVISUAL1.Preencher Bdados, SQL, 1000, 4000, 1000, 2000, 2000, 2000, 2000, 1000, 0, 1500, 1500, 1500, 4000
End Sub

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub

Private Sub cmdExcluir_Click()
    If txtIM <> "" Then
        If Confirma("Deseja excluir o registro selecionado?", "CIAP") Then
            If Bdados.DeletaDados("TAB_CONTRIBUINTE_ESTIMADO", "TCE_TCI_IM ='" & txtIM & "' and TCE_EXERCICIO = '" & txtDataInicial & "'") Then
                If Confirma("Deseja Limpar o Histórico?", "CIAP") Then
'                    cmdFinaliza_Click
                    Bdados.DeletaDados "TAB_CONTRIBUINTE_ESTIMADO_HIST", "TCE_TCI_IM ='" & txtIM & "' and TCE_EXERCICIO = '" & txtDataInicial & "'"
                    
                End If
                Avisa "Operação concluída com sucesso."
                cmLimpar_Click
                cmdBuscar_Click
            End If
        End If
        End If
End Sub

Private Sub cmdFinaliza_Click()
    Dim Valores As String
    Dim Campos As String
    Dim i As Integer
    Dim Motivo As String
    
    Do
        Motivo = Entrada("Informe o Motivo.", "CIAP")
    Loop Until Motivo <> ""
    Campos = "TCE_TCI_IM,TCE_EXERCICIO,tce_base_calculo_anual_ufm,TCE_STATUS,TCE_DATA_PROCEDIMENTO,TCE_PROCESSO,TCE_VALOR_MENSAL,TCE_BASE_CALCULO_ANUAL,TCE_USUARIO,TCE_MOTIVO,TCE_DATA_PROCESSO"
    Valores = Bdados.PreparaValor(txtIM, txtDataInicial, Bdados.Converte(txtValor, TCMonetario), _
                    cboProcedimento.Coluna(1).VALOR, Bdados.Converte(Date, TCDataHora), txtProcesso, Bdados.Converte(txtValorMensal, TCMonetario), Bdados.Converte(txtValorAnual, TCMonetario), Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto), Bdados.Converte(Motivo, tctexto), Bdados.Converte(txtDataProcesso, TCDataHora))
    If Bdados.GravaDados("TAB_CONTRIBUINTE_ESTIMADO", Valores, Campos, "TCE_TCI_IM ='" & txtIM & "' and TCE_EXERCICIO = '" & txtDataInicial & "'") Then
        Avisa "Registro gravado com sucesso."
        cmLimpar_Click
        cmdBuscar_Click
    Else
        Avisa "Erro ao gravar registro."
    End If
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIMConsulta
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
    
    cboProcedimentoConsulta.PreencherGeral Bdados, "PROCEDIMENTO ESTIMATIVA"

    cboProcedimento.PreencherGeral Bdados, "PROCEDIMENTO ESTIMATIVA"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub grdVISUAL1_DblClick()
    If grdVISUAL1.ListItems.Count >= 1 Then
        txtIM = grdVISUAL1.SelectedItem
        txtIM_LostFocus
        cboProcedimento.SetarLinha grdVISUAL1.SelectedItem.SubItems(8), 1
        txtDataInicial = grdVISUAL1.SelectedItem.SubItems(2)
        txtValor = grdVISUAL1.SelectedItem.SubItems(3)
        txtProcesso = grdVISUAL1.SelectedItem.SubItems(7)
        txtDataProcesso = grdVISUAL1.SelectedItem.SubItems(9)
    End If
    
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
      Base = Base * CDbl(TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ","))
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
Private Sub txtIM_LostFocus()
    If Trim$(txtIM) <> "" Then
        txtIM = BuscaContribuinte(txtIM, txtRazao)
    End If
End Sub


Private Sub txtValor_Change()
    Calcula
End Sub

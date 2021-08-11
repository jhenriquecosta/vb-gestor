VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TOBR201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2340
      Left            =   90
      TabIndex        =   26
      Top             =   5235
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4128
      Altura          =   1905
      Caption         =   " Dados da Obrigração"
      CorTexto        =   0
      CorFaixa        =   8421504
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDescontoReal 
         Height          =   540
         Left            =   7440
         TabIndex        =   13
         Top             =   840
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   953
         Caption         =   "Desconto R$"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtParcelamento 
         Height          =   495
         Left            =   6360
         TabIndex        =   36
         Top             =   315
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         Caption         =   "Cod. Parcelamento"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtMotivo 
         Height          =   780
         Left            =   60
         TabIndex        =   14
         Top             =   1440
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   1376
         Caption         =   "Observação"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtDesconto 
         Height          =   540
         Left            =   6390
         TabIndex        =   12
         Top             =   810
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   953
         Caption         =   "Descto (%)"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         CorRotulo       =   192
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtCorrecao 
         Height          =   540
         Left            =   4980
         TabIndex        =   11
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   953
         Caption         =   "Correcão"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtVence 
         Height          =   510
         Left            =   4950
         TabIndex        =   7
         Tag             =   "Data Vencimento"
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         Caption         =   "Vencimento"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtTributo 
         Height          =   510
         Left            =   3570
         TabIndex        =   29
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   900
         Caption         =   "Tributo"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtJuros 
         Height          =   540
         Left            =   2160
         TabIndex        =   9
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   953
         Caption         =   "Juros"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   540
         Left            =   75
         TabIndex        =   8
         Tag             =   "Valor Obrigação"
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   953
         Caption         =   "Valor Original"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoRotulo=   1
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtMulta 
         Height          =   540
         Left            =   3570
         TabIndex        =   10
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   953
         Caption         =   "Multa"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   510
         Left            =   2190
         TabIndex        =   28
         Tag             =   "Período"
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   900
         Caption         =   "Periodo"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtInscricao 
         Height          =   510
         Left            =   90
         TabIndex        =   27
         Top             =   300
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
         Caption         =   "Contribuinte"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   8415
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1085
      Begin VTOcx.cmdVISUAL cmdObrig 
         Height          =   375
         Left            =   7560
         TabIndex        =   17
         Top             =   150
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         Caption         =   "&Alterar Valores"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10530
         TabIndex        =   19
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   9360
         TabIndex        =   18
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Critérios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   3
      Left            =   60
      TabIndex        =   22
      Top             =   675
      Width           =   11565
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
         TabIndex        =   0
         Top             =   150
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   390
         TabIndex        =   23
         Top             =   810
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   7470
         TabIndex        =   5
         Top             =   1500
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   529
         Caption         =   "Periodo Final"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   270
         TabIndex        =   4
         Top             =   1500
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         Caption         =   "Periodo Inicial"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   10200
         TabIndex        =   6
         Top             =   1500
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   690
         TabIndex        =   31
         Top             =   1140
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   690
         TabIndex        =   1
         Top             =   495
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3645
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   4125
         TabIndex        =   2
         Top             =   480
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   7605
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtDAM 
         Height          =   300
         Left            =   8070
         TabIndex        =   3
         Top             =   480
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Número DAM"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   24
         Top             =   1590
         Width           =   45
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1138
      Icone           =   "TOBR201.frx":0000
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   20
      Top             =   90
      Width           =   375
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   705
      Left            =   90
      TabIndex        =   30
      Top             =   7605
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1244
      Altura          =   1905
      Caption         =   " Dados complementares"
      CorTexto        =   0
      CorFaixa        =   8421504
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Tag             =   "Status"
         Top             =   300
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtTaxas 
         Height          =   300
         Left            =   75
         TabIndex        =   15
         Top             =   330
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   529
         Caption         =   "Taxas já inclusas"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   2610
      Left            =   60
      TabIndex        =   32
      Top             =   2640
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   4604
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtEnderecoContrib 
      Height          =   300
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   529
      Caption         =   "Periodo Inicial"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      MinLen          =   4
      AutoTAB         =   -1  'True
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "Principal"
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "reimprime"
      End
   End
End
Attribute VB_Name = "TOBR201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodObrigacao As String
Dim NovaData  As String
Dim Cobranca As New VSCobranca
Dim InscProprietario As String
Private Function CriticaCampos() As Boolean
    CriticaCampos = True
    If Not Edita.CriticaCampos(Me) Then
        CriticaCampos = False
        Exit Function
    End If
    If Len(txtPeriodoInicial) <> Len(txtPeriodoFinal) Then
        Avisa "Período inconsistente."
        txtPeriodoInicial.SetFocus
        CriticaCampos = False
        Exit Function
    End If
    If Len(txtPeriodoInicial) > 4 Then
        If Right(Trim(txtPeriodoInicial), 4) <> Right(Trim(txtPeriodoFinal), 4) Then
            Avisa "Período deve ser dentro do mesmo ano."
            txtPeriodoInicial.SetFocus
            CriticaCampos = False
        End If
    End If
End Function
Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Set Obrig = New Obrigacao
    If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), txtIM, , , txtPeriodoInicial, _
            txtPeriodoFinal, , , , txtImovel, , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), txtDAM) Then
        Avisa "Nenhum registro encontrado."
        cboImposto.SetFocus
    End If
End Sub
Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstObrig.ListItems.Clear
    cboImposto.SetFocus
End Sub
Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub
Private Sub cmdObrig_Click()
    Dim Obrig As New Obrigacao
    Dim Resultado As Boolean
    Dim Motivo As String
    
    Motivo = txtMotivo
    
    If Not CriticaCampos Then Exit Sub
    
    If Not Util.Confirma("Confirma a Alteração da Obrigação?") Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    If Obrig.AlteraObrigacaoTOBR201(CodObrigacao, txtVence, txtValor, txtMulta, txtJuros, txtTaxas, txtCorrecao, CInt(cboStatus.Coluna(1).Valor), Nvl(txtDesconto, 0), txtPeriodo, Motivo, txtParcelamento) Then
        'ALTERO NA TAB_COTAS_PARCELAMENTO DE ACORDO COM O NOVO PROCESSO DO PARCELAMENTO...
         If lstObrig.SelectedItem.SubItems(16) <> "" And lstObrig.SelectedItem.SubItems(16) <> "0" Then
            'checo se o parcelamento existe
            'É UM PARCELAMENTO
            Dim RsCotsa As VSRecordset
            If Bdados.AbreTabela("select TCP_STATUS_OBRIGACAO_PARCELA from TAB_COTAS_PARCELAMENTO where TCP_NUM_COTA = '" & lstObrig.SelectedItem & "'", RsCotsa) Then
                If Not IsNull(RsCotsa.Fields("TCP_STATUS_OBRIGACAO_PARCELA")) Then
                    Bdados.GravaDados "TAB_COTAS_PARCELAMENTO", CStr(cboStatus.Coluna(1).Valor), "TCP_STATUS_OBRIGACAO_PARCELA", "TCP_NUM_COTA = '" & lstObrig.SelectedItem & "'"
                End If
            End If
         End If
        Avisa "Registro gravado."
        cmdBuscar_Click
    Else
        Avisa "Problemas ao gravar registro."
    End If
    'BCP
    Dim c As String, v As String, d As String
        
    If Len(txtDescontoReal) > 0 Then
        d = txtDescontoReal
    Else
        d = 0
    End If
    c = "tcc_desconto_concedido"
     v = Bdados.PreparaValor(Bdados.Converte(CCur(txtDescontoReal), TCDuplo))
    v = Bdados.GravaDados("Tab_Conta_Contribuinte", v, c, "tcc_codigo_conta =" & CodObrigacao)
    'FIM BCP
    
    Screen.MousePointer = 0
End Sub
Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIM
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub
Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    txtDescontoReal = 0
End Sub
Private Sub lstObrig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lstObrig.SelectedItem Is Nothing Then
        If Button = 2 Then
            'BCP ANTES
            mnuReimprime.Caption = "Imprimir DAM da obrigação nº " & lstObrig.SelectedItem
            Me.PopupMenu mnuGeral
            
           
             
        End If
    End If
End Sub
Private Sub mnuReimprime_Click()
    Dim Cobranca As New VSCobranca
    
    If Not Cobranca.LiberaImpressaoDam(Nvl(lstObrig.SelectedItem.SubItems(15), 0)) Then Exit Sub
    With lstObrig.SelectedItem
        NovaData = Imposto.DataVencimentoNova(.SubItems(5))
        If Trim(NovaData) = "" Then Exit Sub
    End With
    
    If Trim(txtImovel) = "" Then
        ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela, , , txtIM, txtEnderecoContrib
    Else
        ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela, , , InscProprietario, txtEnderecoContrib
    End If
End Sub
Private Sub lstObrig_DblClick()
    Dim Obrig As New Obrigacao
    txtMotivo = ""
    If lstObrig.ListItems.Count = 0 Then Exit Sub
    CodObrigacao = lstObrig.SelectedItem
    txtInscricao = lstObrig.SelectedItem.SubItems(1)
    txtTributo = lstObrig.SelectedItem.SubItems(3)
    txtPeriodo = lstObrig.SelectedItem.SubItems(4)
    txtPeriodo = IIf(Len(Trim(txtPeriodo)) = 4, txtPeriodo, Right(txtPeriodo, 2) & Left(txtPeriodo, 4))
    txtVence = lstObrig.SelectedItem.SubItems(5)
    txtValor = Format(lstObrig.SelectedItem.SubItems(6), Const_Monetario)
    txtMulta = Format(lstObrig.SelectedItem.SubItems(8), Const_Monetario)
    txtJuros = Format(lstObrig.SelectedItem.SubItems(7), Const_Monetario)
    txtCorrecao = Format(IIf(Len(lstObrig.SelectedItem.SubItems(17)) = 0, 0, lstObrig.SelectedItem.SubItems(17)), Const_Monetario)
    txtTaxas = Format(Nvl(lstObrig.SelectedItem.SubItems(10), 0), Const_Monetario)
    txtParcelamento = lstObrig.SelectedItem.SubItems(16)
    cboStatus.SetarLinha Format(Nvl(lstObrig.SelectedItem.SubItems(15), -1)), 1
    txtDesconto = Format(Nvl(lstObrig.SelectedItem.SubItems(18), 0), Const_Monetario)
    Dim rs As VSRecordset
    If Bdados.AbreTabela("SELECT TOC_OBSERVACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & CodObrigacao, rs) Then
        txtMotivo = IIf(IsNull(rs(0)), "", rs(0))
    Else
        txtMotivo = ""
    End If
End Sub
Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIM) = 10 Or Len(txtIM) = 11 Then
            Ic = Imposto.FormataInscricao(txtIM, InscContrib)
        Else
            Ic = txtIM
        End If
    Else
            Ic = txtIM
    End If
    txtIM = BuscaContribuinte(Ic, txtRazao, txtEndereco)
End Sub
Private Sub txtImovel_LostFocus()
    Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, InscProprietario, etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIM.SetFocus
        End If
    End If
End Sub


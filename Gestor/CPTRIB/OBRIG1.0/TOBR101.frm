VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TOBR101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   16
      Top             =   7530
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5010
         TabIndex        =   11
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         Caption         =   "&Imprimir DAM"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdObrig 
         Height          =   375
         Left            =   6750
         TabIndex        =   8
         Top             =   90
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         Caption         =   "&Gerar Obrigação"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9870
         TabIndex        =   10
         Top             =   90
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
         Left            =   8685
         TabIndex        =   9
         Top             =   90
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
      Height          =   2925
      Index           =   3
      Left            =   60
      TabIndex        =   14
      Top             =   660
      Width           =   11100
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1845
         Left            =   150
         TabIndex        =   22
         Top             =   510
         Width           =   10935
         Begin VB.TextBox txtEndereco 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1395
            TabIndex        =   23
            Top             =   795
            Width           =   9450
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   300
            Left            =   585
            TabIndex        =   1
            Top             =   45
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
         Begin VTOcx.txtVISUAL txtRazao 
            Height          =   300
            Left            =   255
            TabIndex        =   24
            Top             =   435
            Width           =   10590
            _ExtentX        =   18680
            _ExtentY        =   529
            Caption         =   "Nome/Razão"
            Text            =   ""
            Enabled         =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.txtVISUAL txtPeriodoFinal 
            Height          =   300
            Left            =   2670
            TabIndex        =   4
            Tag             =   "Periodo Final"
            Top             =   1170
            Width           =   2400
            _ExtentX        =   4233
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
            Left            =   150
            TabIndex        =   3
            Tag             =   "Periodo Inicial"
            Top             =   1140
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
         Begin VTOcx.txtVISUAL txtOrigem 
            Height          =   300
            Left            =   270
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   529
            Caption         =   "Doc. Origem"
            Text            =   ""
            Requerido       =   0   'False
            MinLen          =   4
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.cboVISUAL cboStatus 
            Height          =   315
            Left            =   7260
            TabIndex        =   5
            Top             =   1170
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   556
            Caption         =   "Status"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
            Height          =   315
            Left            =   3540
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   30
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
         End
         Begin VTOcx.txtVISUAL txtFator 
            Height          =   300
            Left            =   2670
            TabIndex        =   7
            Top             =   1530
            Visible         =   0   'False
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            Caption         =   "Fator Multiplicador"
            Text            =   ""
            Restricao       =   2
            Requerido       =   0   'False
            MinLen          =   4
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtImovel 
            Height          =   300
            Left            =   6090
            TabIndex        =   2
            Top             =   30
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
            Left            =   9570
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   30
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   ""
            Acao            =   5
         End
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   150
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VB.Label lblGerado 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   5520
         TabIndex        =   17
         Top             =   2535
         Width           =   5280
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   8430
         TabIndex        =   15
         Top             =   2370
         Width           =   45
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   1138
      Icone           =   "TOBR101.frx":2123
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   12
      Top             =   90
      Width           =   375
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1920
      Left            =   75
      TabIndex        =   20
      Top             =   7635
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   3387
      Caption         =   "Taxas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   3840
      Left            =   60
      TabIndex        =   18
      Top             =   3615
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   6773
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.txtVISUAL txtEnderecoContrib 
      Height          =   300
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
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
   Begin VB.Menu mnuOpcao 
      Caption         =   "Opcao"
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "Reimprime"
      End
   End
End
Attribute VB_Name = "TOBR101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public String_Taxas    As String
Public Total_Taxas     As Double
Private ValorFixoTaxa As Double
Private ValorCalculadoTaxa As Double
Dim InscProprietario As String
Private Function CriticaCampos() As Boolean
    CriticaCampos = True
    
    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
        If cboImposto.Coluna(0).Valor = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ALVARA)) Or cboImposto.Coluna(0).Valor = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_tfs)) Then
            txtPeriodoInicial.Tag = ""
            txtPeriodoFinal.Tag = ""
        End If
    End If
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

Private Sub cboImposto_Click()
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select tpi_tipo_tributo,tpi_valor_taxa_fixa,tpi_tipo_inscricao,tpi_tipo_ic from tab_parametro_imposto where tpi_tip_cod_imposto ='" & cboImposto.Coluna(0).Valor & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        ValorFixoTaxa = IIf(IsNull(Rs!tpi_valor_taxa_fixa), 0, Rs!tpi_valor_taxa_fixa)
        If ValorFixoTaxa = 0 Then
            txtFator.Visible = False
        Else
            txtFator.Visible = True
        End If
    End If
    
    
End Sub

Private Sub cmdCancela_Click()
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    Dim CodBarra As New CodigoDeBarra
'    Dim LinhaDigitavel  As String
'    Dim LinhaBarra As String
'
'    Dim i As Integer
'    Sql = "SELECT TCO_INSCRICAO,TCO_TIP_COD_IMPOSTO, TCO_VALOR_PARCELA,TCO_PERIODO,TCO_DATA_VENCIMENTO," & _
'            "TCO_NUM_PARCELA,TCO_COD_OBRIGACAO_PARCELA FROM TAB_COTAS_OBRIGACAO " & _
'            "ORDER BY TCO_TPA_COD_PARCELAMENTO, TCO_VALOR_PARCELA"
'    If Bdados.AbreTabela(Sql, Rs) Then
'        Rs.MoveFirst
'        i = 0
'        Do
'            i = i + 1
'            LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(Trim(Rs!TCO_INSCRICAO), Trim(Rs!TCO_TIP_COD_IMPOSTO), _
'                 Format(Rs!TCO_VALOR_PARCELA, Const_Monetario), Rs!TCO_PERIODO, Picture, Rs!TCO_DATA_VENCIMENTO, Rs!TCO_NUM_PARCELA, Rs!TCO_COD_OBRIGACAO_PARCELA)
'            LinhaBarra = "convert(Char(50),'" & CodBarra.LinhaBarraGerada & "')"
'            Bdados.AtualizaDados "TAB_COTAS_OBRIGACAO", _
'                Bdados.PreparaValor(Bdados.Converte(LinhaDigitavel, tctexto), LinhaBarra), _
'                "TCO_LINHA_DIGITAVEL,TCO_LINHA_BARRA", _
'                "TCO_COD_OBRIGACAO_PARCELA =" & Rs!TCO_COD_OBRIGACAO_PARCELA
'            cmdCancela.Caption = i
'            DoEvents
'            Rs.MoveNext
'        Loop While Not Rs.EOF
'    End If
'    Avisa "Fim do Processo!!!!!!!"
'    Exit Sub
    
    Edita.LimpaCampos Me
    lstObrig.ListItems.Clear
    cboImposto.SetFocus
    lblGerado = ""
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdObrig_Click()
    Dim Obrig As New Obrigacao
    Dim Resultado As Boolean
    Dim Qtd As String
    Dim InsCad As String, Grupo As Byte 'criado para a utilizacao do grupo de inscruicao
    If Not CriticaCampos Then Exit Sub
    
    If Not Util.Confirma("Confirma a geração da obrigação") Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Grupo = IIf(Trim(txtImovel) = "", 0, 1)
    lstObrig.Preencher Bdados, ""
'    If cboImposto.Coluna(0).Valor = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNFIXO)) Then
'        Obrig.GeraExtratoUnificado = Confirma("Deseja gerar extrato unificado de débitos para o contribuinte?")
'    End If
    If Obrig.CriaObrigacao(CStr(cboImposto.Coluna(0).Valor), Edita.TiraPic(txtPeriodoInicial, "/"), _
                Edita.TiraPic(txtPeriodoFinal, "/"), txtIM, ValorCalculadoTaxa, CInt(cboStatus.Coluna(1).Valor), , , , , Grupo, lblGerado, , , txtOrigem, , txtImovel) Then
                'Raimundo Substituir a função MostraObrigacaoAbaixo por que o desempenho no Oracle está muito baixo...
        'Obrig.MostraObrigacaoGerada lstObrig, CStr(cboImposto.Coluna(0).Valor), txtIM, _
        '         , , , , txtPeriodoInicial, txtPeriodoFinal, Edita.TiraTudo(txtOrigem), txtImovel
        '=============================================================================================
        If txtIM <> "" Then
            Inscri = txtIM
        ElseIf txtImovel <> "" Then
            Inscri = txtImovel
        End If
        Obrig.MostraObrigacaoGerada lstObrig, CStr(cboImposto.Coluna(0).Valor), txtIM, _
            , etsCreditoOriginalAberto, , , _
             txtPeriodoInicial, txtPeriodoFinal, , txtImovel, , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)
        cboImposto.SetFocus
        Informa "Obrigação(ões) geradas com sucesso."
    Else
        Informa "Não foi possivel gerar a(s) obrigacão." & vbCrLf & "Verifique parametro na definição de tributos."
    End If
    txtPeriodoInicial.Tag = txtPeriodoInicial.Caption
    txtPeriodoFinal.Tag = txtPeriodoFinal.Caption
    Screen.MousePointer = 0
    ValorCalculadoTaxa = 0
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

Private Sub Form_Activate()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, True
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    Rem GrdTaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
End Sub

Private Sub Frame1_DblClick(Index As Integer)
'    Dim Obrig As New Obrigacao
'    Obrig.GBDam
'    Obrig.GBDARM
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo trata
    Screen.MousePointer = 11
    Dim i As Double
    For i = 1 To lstObrig.ListItems.Count
        With lstObrig.ListItems
            .Item(i).Selected = True
            'Pego As taxas
            Call Pega_taxas
'            ImprimeSelecionado lstObrig, txtRazao, txtEndereco, , , , String_Taxas, Total_Taxas
            If Trim(txtImovel) = "" Then
                ImprimeSelecionado lstObrig, txtRazao, txtEndereco, , , , String_Taxas, Total_Taxas, txtIM, txtEndereco
            Else
                ImprimeSelecionado lstObrig, txtRazao, txtEndereco, , , , String_Taxas, Total_Taxas, InscProprietario, txtEnderecoContrib
            End If
        End With
        DoEvents
    Next
    Avisa "Impressão concluída."
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    Screen.MousePointer = 0
    Erro Err.Description
End Sub

Private Sub grdVISUAL1_Click()

End Sub


Private Sub lstObrig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lstObrig.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuReimprime.Caption = "Imprimir DAM da obrigação nº " & lstObrig.SelectedItem
            Me.PopupMenu mnuOpcao
        End If
    End If
End Sub

Private Sub mnuReimprime_Click()
    Dim NovaData As String
    Dim Conta As New ContaCorrente
    Dim NovoJuro As Double
    Dim NovaMulta As Double
    Dim Correcao As Double
    Dim Sql As String
    Dim Rs_Dividas As VSRecordset
    Dim Rs_Carro As VSRecordset
    If lstObrig.SelectedItem Is Nothing Then Exit Sub
    
    With lstObrig.SelectedItem
        NovaData = Imposto.DataVencimentoNova(.SubItems(5))
        If Trim(NovaData) = "" Then Exit Sub
            Correcao = Conta.CalculaValoresCorrecaoAvulso(.SubItems(11), .SubItems(4), .SubItems(5), NovaData, .SubItems(6))
            NovoJuro = Conta.CalculaValoresJurosAvulsos(.SubItems(11), .SubItems(4), EtcCreditoTributario, NovaData, .SubItems(5), .SubItems(6) + Correcao)
            NovaMulta = Conta.CalculaValoresMultaAvulsos(.SubItems(11), .SubItems(4), EtcCreditoTributario, NovaData, .SubItems(5), .SubItems(6) + Correcao)
 '       End If
    End With
'    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData
    If Trim(txtImovel) = "" Then
        ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela, , , txtIM, txtEndereco
    Else
        ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela, , , InscProprietario, txtEnderecoContrib
    End If
End Sub

Private Sub txtFator_LostFocus()
    ValorCalculadoTaxa = ValorFixoTaxa * CDbl(Nvl(Trim(txtFator), 0))
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
Private Sub Pega_taxas()
    Dim i As Integer
    Dim Pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To GrdTaxas.ListItems.Count
        If GrdTaxas.ListItems(i).Checked Then
            Pos = InStr(GrdTaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(GrdTaxas.ListItems(i).SubItems(2))
        End If
    Next
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


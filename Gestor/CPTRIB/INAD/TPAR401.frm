VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.1#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   7215
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   5925
         TabIndex        =   12
         Top             =   135
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   661
         Caption         =   "&Imprimir parcelas"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9270
         TabIndex        =   5
         Top             =   135
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
         Left            =   8100
         TabIndex        =   4
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL grdCotas 
      Height          =   1950
      Left            =   60
      TabIndex        =   10
      Top             =   5220
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   3440
      Caption         =   "Cotas Geradas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.grdVISUAL grdOriginal 
      Height          =   1815
      Left            =   60
      TabIndex        =   14
      Top             =   3660
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   3201
      Caption         =   "Débitos Originais"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   6
      Top             =   -420
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   1725
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   645
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   3043
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Contribuinte"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   570
         TabIndex        =   15
         Top             =   930
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   465
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   195
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   705
         Left            =   9045
         TabIndex        =   3
         Top             =   900
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1244
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   300
         TabIndex        =   16
         Top             =   1290
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3000
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   555
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   5130
         TabIndex        =   2
         Top             =   555
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
         Left            =   8610
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   540
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtim 
         Height          =   300
         Left            =   30
         TabIndex        =   1
         Top             =   570
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Contribuinte"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
   End
   Begin VTOcx.grdVISUAL lstParc 
      Height          =   1515
      Left            =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   2672
      Caption         =   "Parcelamentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TPAR401.frx":2123
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1770
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   13
      Top             =   9000
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu mnumenu 
      Caption         =   "&menu"
      Visible         =   0   'False
      Begin VB.Menu mnuExcluir 
         Caption         =   "&Excluir"
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "&Reimprimir"
      End
   End
   Begin VB.Menu mnuCotas 
      Caption         =   "&Cotas"
      Visible         =   0   'False
      Begin VB.Menu mnuImprime 
         Caption         =   "&Cotas"
      End
   End
End
Attribute VB_Name = "TPAR401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Obrig As New Obrigacao
Dim String_Taxas As String
Dim Total_Taxas As Double

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstParc.ListItems.Clear
    grdCotas.ListItems.Clear
    grdOriginal.ListItems.Clear
    
    lstParc.Mensagem = ""
    grdCotas.Mensagem = ""
    grdOriginal.Mensagem = ""
    txtIm.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    Dim Cobranca As New VSCobranca
    Dim Razao As String
    Dim CpfCgc As String
    Dim EnderecoPessoa As String
    Dim EnderecoImovel As String
    Dim NomeTributo As String
    Dim i As Double
    EnderecoImovel = IIf(Trim(txtRazao) = "", txtEndereco, "")
    EnderecoPessoa = IIf(Trim(EnderecoPessoa) = "", txtEndereco, "")
    For i = 1 To grdCotas.ListItems.Count
        With grdCotas.ListItems
            .Item(i).Selected = True
            Cobranca.ImprimeDam Rpt, .Item(i), txtIm, txtRazao, CpfCgc, EnderecoPessoa, txtIm, EnderecoImovel, .Item(i).SubItems(9), .Item(i).SubItems(2), _
                .Item(i).SubItems(10), IIf(Len(.Item(i).SubItems(3)) = 4, .Item(i).SubItems(3), Right(.Item(i).SubItems(3), 2) & Left(.Item(i).SubItems(3), 4)), .Item(i).SubItems(5), _
                4, .Item(i).SubItems(4), lstParc.SelectedItem.SubItems(5), .Item(i).SubItems(6), 0, .Item(i).SubItems(7), 0, 0, "", "", PicBarra, , , , , , , , , , , tdiImpressora
        End With
        DoEvents
    Next
    Avisa "Impressão concluída."
End Sub

Private Sub cmdParcela_Click()
    Dim SQL As String
    Dim CCorrente As New ContaCorrente
    Dim CodImp As String
    Dim Rs As VSRecordset
    Dim RsCONTRIB As VSRecordset
    Dim UltParc As Double
     
    lstParc.ListItems.Clear
    grdCotas.ListItems.Clear
    grdOriginal.ListItems.Clear
    lstParc.Mensagem = ""
    grdCotas.Mensagem = ""
    lstParc.Mensagem = ""
    Screen.MousePointer = 11
    SQL = "SELECT TPA_NUM_PARCELAMENTO as Parcelamento,"
    SQL = SQL & " TPA_INSCRICAO AS Inscricao,TPA_PERIODO AS Periodo,"
    SQL = SQL & " TIP_SIGLA_IMPOSTO as Tributo,"
    SQL = SQL & " TPA_DATA_FINANCIAMENTO as Data, "
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        SQL = SQL & Bdados.Converte("TPA_VALOR_PARCELADO", TCDuplo) & " as Valor_Parcelado,"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        SQL = SQL & " TPA_VALOR_PARCELADO as Valor_Parcelado,"
    End If
    SQL = SQL & " TPA_NUM_COTAS as Cotas,TPA_STATUS_PARCELAMENTO as Situacão,"
    SQL = SQL & " TPA_PERIODO_INICIAL,TPA_PERIODO_FINAL "
    SQL = SQL & " FROM TAB_PARCELAMENTO,TAB_IMPOSTO "
    SQL = SQL & " WHERE  TPA_TIP_COD_IMPOSTO=TIP_COD_IMPOSTO "
    SQL = SQL & " AND TPA_STATUS_PARCELAMENTO <> " & stsParcelamentoCancelado
    
    If Trim(cboImposto) <> "" Then SQL = SQL & " and TPA_TIP_COD_IMPOSTO='" & cboImposto.Coluna(0).Valor & "'"
    If Trim(txtIm) <> "" Then
        SQL = SQL & " and (TPA_INSCRICAO='" & Trim(txtIm) & "' and TPA_TIPO_INSCRICAO = " & IIf(cboImposto.Coluna(0).Valor = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), 1, 2) & ")"
    ElseIf Trim(txtImovel) <> "" Then
        SQL = SQL & " and (TPA_INSCRICAO='" & Trim(txtImovel) & "' and TPA_TIPO_INSCRICAO =1)"
    End If
    lstParc.Preencher Bdados, SQL, 1200, 1200, 1000, 1200, 1200, 1500, 800, 0, 0, 0
    grdCotas.ListItems.Clear
    If lstParc.ListItems.Count > 0 Then
        lstParc.Mensagem = "Valor Total Parcelado: R$" & Format(lstParc.Colunas(6).Soma, Const_Monetario)
    Else
        Util.Avisa "Não existe parcelamento para esse contribuinte."
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboImposto.Preencher Bdados, "Select  tip_cod_imposto,TIP_sigla_IMPOSTO  " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto,tip_nome_imposto From TAB_IMPOSTO order by TIP_sigla_IMPOSTO asc", 1
    AtualizaCabecalho lstParc
    'Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
End Sub


Private Sub lstParc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstParc, ColumnHeader
End Sub

Private Sub grdCotas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lstParc.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuImprime.Caption = "Imprimir cota nº " & grdCotas.SelectedItem
            Me.PopupMenu mnuCotas
        End If
    End If
End Sub

Private Sub lstParc_Click()
    Dim SQL As String
    Dim Obrig As New Obrigacao
    Dim Insc As String
    
    If txtIm <> "" Then
        Insc = txtIm
    Else
        Insc = txtImovel
    End If
    If lstParc.ListItems.Count > 0 Then
        If Not Obrig.MostraObrigacaoGerada(grdOriginal, CStr(cboImposto.Coluna(0).Valor), txtIm, _
              , , , , _
            , , , txtImovel, , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), , lstParc.SelectedItem, " AND TOC_STATUS_OBRIGACAO = 4 ") Then
            Avisa "Nenhum registro encontrado."
        End If
        SQL = " Select TCp_NUM_COTA AS Documento,tcp_inscricao as Inscrição,"
        SQL = SQL & " TIP_SIGLA_IMPOSTO AS Tributo,"
        SQL = SQL & " TPA_PERIODO AS Periodo,"
        SQL = SQL & " TCp_DATA_VENCIMENTO AS Vencimento,TCp_NUM_PARCELA as Cota,"
        SQL = SQL & " TCp_VALOR_PARCELA As Valor, TCp_VALOR_JUROS As Juros, TCp_VALOR_PARCELA"
        SQL = SQL & " + TCp_VALOR_JUROS as Total,tip_cod_imposto as Imposto,"
        SQL = SQL & " tip_nome_imposto as Descrição,TGE_NOME as SITUACAO,TCP_STATUS_OBRIGACAO_PARCELA"
        SQL = SQL & " From tab_parcelamento, tab_cotas_parcelamento, tab_imposto,VIS_STATUS_OBRIGACAO"
        SQL = SQL & " where  tpa_num_parcelamento = TCp_TPA_COD_PARCELAMENTO and TCP_STATUS_OBRIGACAO_PARCELA = TGE_CODIGO and"
        SQL = SQL & " TPA_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
        SQL = SQL & " AND TCP_TPA_COD_PARCELAMENTO =   '" & lstParc.SelectedItem & "'   order by TCp_NUM_PARCELA"
        grdCotas.Preencher Bdados, SQL, 0, 1200, 1000, 800, 1200, 600, 800, 800, 1200, 0, 0, 2500, 0
        If grdCotas.ListItems.Count > 0 Then grdCotas.Mensagem = "Total Parcelamento: R$" & Format(grdCotas.Colunas(7).Soma, Const_Monetario) & " x Acréscimo na dívida original: R$" & Format(grdCotas.Colunas(9).Soma - grdCotas.Colunas(7).Soma, Const_Monetario)
    End If
End Sub

Private Sub lstParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then lstParc_Click
End Sub

Private Sub lstParc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lstParc.SelectedItem Is Nothing Then
        If Button = 2 Then
            MnuExcluir.Caption = "Cancelar parcelamento nº " & lstParc.SelectedItem
            mnuReimprime.Caption = "Reimprimir termo de parcelamento nº " & lstParc.SelectedItem
            Me.PopupMenu mnumenu
        End If
    End If
End Sub

Private Sub mnuExcluir_Click()
    TPAR201.Tag = Right(MnuExcluir.Caption, 8)
    TPAR201.Show
End Sub

Private Sub mnuImprime_Click()
    Dim Cobranca As New VSCobranca
    Dim Razao As String
    Dim CpfCgc As String
    Dim EnderecoPessoa As String
    Dim EnderecoImovel As String
    Dim NomeTributo As String
    Dim i As Double
    Dim NovaData As String
    Dim NovoJuro As Double
    Dim NovaMulta As Double
    
    Dim Conta As New ContaCorrente
    Dim obs As String
    Dim Taxa As Double
    If Not Cobranca.LiberaImpressaoDam(Nvl(grdCotas.SelectedItem.SubItems(12), 0)) Then Exit Sub
    NovaData = Imposto.DataVencimentoNova(grdCotas.SelectedItem.SubItems(4))
    If Trim(NovaData) = "" Then Exit Sub
    obs = Util.Entrada("Observacao", "Impressão de Parcela.")
    With grdCotas.SelectedItem
        NovoJuro = .SubItems(7) + Conta.CalculaValoresJurosAvulsos(.SubItems(9), .SubItems(3), EtcCreditoTributario, NovaData, .SubItems(4), .SubItems(6))
        NovaMulta = Conta.CalculaValoresMultaAvulsos(.SubItems(9), .SubItems(3), EtcCreditoTributario, NovaData, .SubItems(4), .SubItems(6))
        Cobranca.ImprimeDam Rpt, grdCotas.SelectedItem, lstParc.SelectedItem.SubItems(1), txtRazao, CpfCgc, txtEndereco, txtIm, txtEndereco, .SubItems(9), .SubItems(2), _
             .SubItems(10), IIf(Len(.SubItems(3)) = 4, .SubItems(3), Right(.SubItems(3), 2) & Left(.SubItems(3), 4)), .SubItems(5), _
            4, NovaData, lstParc.SelectedItem.SubItems(5), .SubItems(6), CStr(NovaMulta), CStr(NovoJuro), Taxa, 0, "", obs, PicBarra, , , , , , , , , , , tdiTela, etdNormal, String_Taxas
    End With
End Sub

Private Sub mnuReimprime_Click()
    Dim obs As String
'    If Not LiberaImpressaoDam(Nvl(lstObrig.SelectedItem.SubItems(15), 0)) Then Exit Sub
    obs = Entrada("Observção", "CIAP")
     With Rpt
            
               If Not .DefinirArquivo(Bdados, App.Path + "\TermoParcela.rpt") Then Exit Sub
               Bdados.GravaDados "TAB_PARCELAMENTO", Bdados.PreparaValor(obs), "TPA_OBSERVACAO", "TPA_NUM_PARCELAMENTO = '" & lstParc.SelectedItem & "'"
               .Formulas "NumParcelamento ", lstParc.SelectedItem
                .Formulas "Municipio ", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                .Formulas "Imposto ", CStr(cboImposto.Coluna(2).Valor)
                .Formulas "Inscricao", IIf(Trim(txtIm) = "", txtImovel, txtIm)
                .Formulas "Contribuinte", txtRazao
                .Formulas "Endereco", txtEndereco
                If IsNumeric(lstParc.SelectedItem.SubItems(5)) Then
                    .Formulas "ValorExtenso", VBA.UCase(Extenso(CDbl(lstParc.SelectedItem.SubItems(5)), "Reais", "Real"))
                End If
                .Formulas "VT_Periodo ", IIf(Len(CStr(lstParc.SelectedItem.SubItems(8))) = 4, CStr(lstParc.SelectedItem.SubItems(8)), Right(CStr(lstParc.SelectedItem.SubItems(8)), 2) & "/" & Left(CStr(lstParc.SelectedItem.SubItems(8)), 4)) & " a " & IIf(Len(CStr(lstParc.SelectedItem.SubItems(9))) = 4, CStr(lstParc.SelectedItem.SubItems(9)), Right(CStr(lstParc.SelectedItem.SubItems(9)), 2) & "/" & Left(CStr(lstParc.SelectedItem.SubItems(9)), 4))
                .Selecao = "{Tab_Parcelamento.TPA_NUM_PARCELAMENTO} = " & lstParc.SelectedItem
                If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "GDA")
                Else
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                End If
                .Titulo = "Termo de Parcelamento"
                .Arvore = False
                .Visualizar
        End With
        Set Rpt = Nothing
End Sub

Private Sub txtIm_LostFocus()
    
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco)
End Sub

Private Sub txtImovel_LostFocus()
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
End Sub

VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TPAR105"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1138
      Icone           =   "TPAR105.frx":0000
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   5460
      Left            =   0
      TabIndex        =   5
      Top             =   645
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   9631
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      Tabs            =   "TPAR105.frx":282A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5070
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   8943
         _Version        =   131082
         TabGuid         =   "TPAR105.frx":28BE
         Begin VB.Frame Frame1 
            Height          =   690
            Left            =   75
            TabIndex        =   22
            Top             =   4335
            Width           =   10095
            Begin VTOcx.txtVISUAL txtParcela 
               Height          =   330
               Left            =   8025
               TabIndex        =   26
               Top             =   240
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   582
               Caption         =   "Parcela"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtPeriodoInicial 
               Height          =   300
               Left            =   585
               TabIndex        =   23
               Tag             =   "Periodo "
               Top             =   255
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   529
               Caption         =   "Periodo "
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtVence 
               Height          =   300
               Left            =   2940
               TabIndex        =   24
               Tag             =   "Data Vencimento"
               Top             =   255
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   529
               Caption         =   "Vencimento"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   300
               Left            =   5400
               TabIndex        =   25
               Tag             =   "Valor Obrigação"
               Top             =   255
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               Caption         =   "Valor"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoTexto=   1
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
         End
         Begin VTOcx.grdVISUAL grdCotas 
            Height          =   4545
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   10110
            _ExtentX        =   17833
            _ExtentY        =   8017
            Caption         =   "Cotas Geradas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
            OcultarRodape   =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5070
         Left            =   -99969
         TabIndex        =   6
         Top             =   30
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   8943
         _Version        =   131082
         TabGuid         =   "TPAR105.frx":28E6
         Begin VTOcx.grdVISUAL grdOriginal 
            Height          =   2205
            Left            =   30
            TabIndex        =   7
            Top             =   3135
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   3889
            Caption         =   "Débitos Originais"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
            OcultarRodape   =   -1  'True
         End
         Begin Threed.SSFrame fra 
            Height          =   1725
            Index           =   0
            Left            =   30
            TabIndex        =   8
            Top             =   0
            Width           =   10170
            _ExtentX        =   17939
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
               TabIndex        =   9
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
               TabIndex        =   10
               Tag             =   "Tributo"
               Top             =   195
               Width           =   8490
               _ExtentX        =   14975
               _ExtentY        =   556
               Caption         =   "Tributo"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.txtVISUAL txtEndereco 
               Height          =   315
               Left            =   300
               TabIndex        =   11
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
               TabIndex        =   12
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
               TabIndex        =   13
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
               TabIndex        =   14
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
               TabIndex        =   15
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
            Height          =   1635
            Left            =   30
            TabIndex        =   16
            Top             =   1770
            Width           =   10185
            _ExtentX        =   17965
            _ExtentY        =   2884
            Caption         =   "Parcelamentos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
            OcultarRodape   =   -1  'True
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   6915
         TabIndex        =   20
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   5865
         TabIndex        =   19
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9135
         TabIndex        =   1
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
         Left            =   7965
         TabIndex        =   0
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
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   2
      Top             =   -420
      Width           =   375
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1770
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   4
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
Attribute VB_Name = "TPAR105"
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
    On Error Resume Next
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
  
End Sub

Private Sub cmdParcela_Click()
    Dim sql As String
    Dim CCorrente As New ContaCorrente
    Dim CodImp As String
    Dim rs As VSRecordset
    Dim RsCONTRIB As VSRecordset
    Dim UltParc As Double
     
    lstParc.ListItems.Clear
    grdCotas.ListItems.Clear
    grdOriginal.ListItems.Clear
    lstParc.Mensagem = ""
    grdCotas.Mensagem = ""
    lstParc.Mensagem = ""
    Screen.MousePointer = 11
    sql = "SELECT TPA_NUM_PARCELAMENTO as Parcelamento,"
    sql = sql & " TPA_INSCRICAO AS Inscricao,TPA_PERIODO AS Periodo,"
    sql = sql & " TIP_SIGLA_IMPOSTO as Tributo,"
    sql = sql & " TPA_DATA_FINANCIAMENTO as Data, "
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        sql = sql & Bdados.Converte("TPA_VALOR_PARCELADO", TCDuplo) & " as Valor_Parcelado,"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        sql = sql & " TPA_VALOR_PARCELADO as Valor_Parcelado,"
    End If
    sql = sql & " TPA_NUM_COTAS as Cotas,TPA_STATUS_PARCELAMENTO as Situacão,"
    sql = sql & " TPA_PERIODO_INICIAL,TPA_PERIODO_FINAL "
    sql = sql & " FROM TAB_PARCELAMENTO,TAB_IMPOSTO "
    sql = sql & " WHERE  TPA_TIP_COD_IMPOSTO=TIP_COD_IMPOSTO "
    sql = sql & " AND TPA_STATUS_PARCELAMENTO <> " & stsParcelamentoCancelado
    
    If Trim(cboImposto) <> "" Then sql = sql & " and TPA_TIP_COD_IMPOSTO='" & cboImposto.Coluna(0).Valor & "'"
    If Trim(txtIm) <> "" Then
        sql = sql & " and (TPA_INSCRICAO='" & Trim(txtIm) & "' and TPA_TIPO_INSCRICAO = " & IIf(cboImposto.Coluna(0).Valor = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), 1, 2) & ")"
    ElseIf Trim(txtImovel) <> "" Then
        sql = sql & " and (TPA_INSCRICAO='" & Trim(txtImovel) & "' and TPA_TIPO_INSCRICAO =1)"
    End If
    lstParc.Preencher Bdados, sql, 1200, 1200, 1000, 1200, 1200, 1500, 800, 0, 0, 0
    grdCotas.ListItems.Clear
    If lstParc.ListItems.Count > 0 Then
        'lstParc.Mensagem = "Valor Total Parcelado: R$" & Format(lstParc.Colunas(6).Soma, Const_Monetario)
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

Private Sub cmdSalvar_Click()
    Dim sql                      As String
    Dim rs                       As VSRecordset
    Dim i                        As Integer
    Dim RegistroApagados         As Integer
    Dim RegistroOriginais        As Integer
    Dim CodObrigacaoGerarda      As String
    Dim RegistroParcelaCancelada As Integer
    Dim Tipo                     As TipoInscricaoObrigacao
    Dim BAlterouParcelamento     As Boolean

    'PROCESSO PARA REUNIFICAÇÃO DE PARCELAS
    '1º APAGAR OS PARCELAMENTOS DA TAB_OBRIGAÇÃO_CONTRBUINTE
    '2º ALTERAR O STATUS DOS DÉBITOS ORIGINAIS PARA PARCELAMENTO CANCELADO
    '3º GERAR UM NOVO REGISTRO NA TAB_AOBRIGACAO_CONTRIBUINTE COM A SOMA DAS PARCELAS EM ABERTA
    '4º ALTERAR AS COTAS DA TAB_COTAS_PARCELAMENTOS PARA PARCELAMENTO CACELADOS
    '5º ALTERAR O STATUS DO PARCELAMENTO DA TAB_PARCELAMENTO PARA CANCELADO

    If grdCotas.ListItems.Count < 1 Then Exit Sub
    
    '1º
        'Bdados.AbreTrans
        For i = 1 To grdCotas.ListItems.Count
            If Bdados.DeletaDados("TAB_OBRIGACAO_CONTRIBUINTE", "TOC_COD_OBRIGACAO = " & grdCotas.ListItems(i)) Then
                RegistroApagados = RegistroApagados + 1
            End If
        Next
     '2º
        For i = 1 To grdOriginal.ListItems.Count
            If Bdados.GravaDados("TAB_OBRIGACAO_CONTRIBUINTE", stsParcelamentoCancelado, "TOC_STATUS_OBRIGACAO", "TOC_COD_OBRIGACAO = " & Bdados.Converte(grdOriginal.ListItems(i), tctexto)) Then
                RegistroOriginais = RegistroOriginais + 1
            End If
        Next
     '3º
        If txtIm <> "" Then
            Tipo = etiContribuinte
        Else
            Tipo = etiImovel
        End If
        CodObrigacaoGerarda = Obrig.CriaObrigacao(Imposto.BuscaCodImposto(lstParc.SelectedItem.SubItems(3)), lstParc.SelectedItem.SubItems(2), lstParc.SelectedItem.SubItems(2), lstParc.SelectedItem.SubItems(1), txtValor, etsCreditoOriginalAberto, etsCriaNova, txtVence, , , , , , , , 0, , Tipo)
     '4º
        For i = 1 To grdCotas.ListItems.Count
            If Bdados.GravaDados("TAB_COTAS_PARCELAMENTO", etsCreditoCancelado, "TCP_STATUS_OBRIGACAO_PARCELA", "TCP_NUM_COTA = " & grdCotas.ListItems(i)) Then
                RegistroParcelaCancelada = RegistroParcelaCancelada + 1
            End If
        Next
     '5º
        BAlterouParcelamento = Bdados.GravaDados("TAB_PARCELAMENTO", stsParcelamentoCancelado, "TPA_STATUS_PARCELAMENTO", "TPA_NUM_PARCELAMENTO  = " & lstParc.SelectedItem)
                
      'VALIDO OS PROCESSOS, SE UM FALHOU CANCELO TUDO
      '1º RegistroApagados
      '2º RegistroOriginais
      '3º CodObrigacaoGerarda
      '4º RegistroParcelaCancelada
      If RegistroApagados = grdCotas.ListItems.Count Then
        If RegistroOriginais = grdOriginal.ListItems.Count Then
            If CodObrigacaoGerarda <> "" Then
                If RegistroParcelaCancelada = grdCotas.ListItems.Count Then
                    If BAlterouParcelamento Then
                       ' Bdados.GravaTrans
                        Util.Avisa "Parcelamento Reunificado com sucesso."
                        cmdCancela_Click
                    Else
                        Erro "Erro na atualização do parcelamento."
                        'Bdados.CancelaTrans
                    End If
                Else
                    Erro "Erro na atualização das parcelas."
                    'Bdados.CancelaTrans
                End If
            Else
                Erro "Erro na reunificação de débitos."
                'Bdados.CancelaTrans
            End If
        Else
            Erro "Erro na atualização dos débitos originais."
            'Bdados.CancelaTrans
        End If
      Else
        Erro "Erro ao excluir parcelas."
        'Bdados.CancelaTrans
      End If
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    cboImposto.Preencher Bdados, "Select  tip_cod_imposto,TIP_sigla_IMPOSTO  " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto,tip_nome_imposto From TAB_IMPOSTO order by TIP_sigla_IMPOSTO asc", 1
    AtualizaCabecalho lstParc
    'Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    txtValor.Enabled = False
End Sub


Private Sub lstParc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstParc, ColumnHeader
End Sub

Private Sub fraVISUAL1_mudancaStatus()

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
    Dim sql As String
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
        sql = " Select TCp_NUM_COTA AS Documento,tcp_inscricao as Inscrição,"
        sql = sql & " TIP_SIGLA_IMPOSTO AS Tributo,"
        sql = sql & " TPA_PERIODO AS Periodo,"
        sql = sql & " TCp_DATA_VENCIMENTO AS Vencimento,TCp_NUM_PARCELA as Cota,"
        sql = sql & " TCp_VALOR_PARCELA As Valor, TCp_VALOR_JUROS As Juros, TCp_VALOR_PARCELA"
        sql = sql & " + TCp_VALOR_JUROS as Total,tip_cod_imposto as Imposto,"
        sql = sql & " tip_nome_imposto as Descrição,TGE_NOME as SITUACAO,TCP_STATUS_OBRIGACAO_PARCELA"
        sql = sql & " From tab_parcelamento, tab_cotas_parcelamento, tab_imposto,VIS_STATUS_OBRIGACAO"
        sql = sql & " where  tpa_num_parcelamento = TCp_TPA_COD_PARCELAMENTO and TCP_STATUS_OBRIGACAO_PARCELA = TGE_CODIGO and"
        sql = sql & " TPA_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
        sql = sql & " AND TCP_TPA_COD_PARCELAMENTO =   '" & lstParc.SelectedItem & "' AND TGE_CODIGO = 2 order by TCp_NUM_PARCELA"
        grdCotas.Preencher Bdados, sql, 0, 1200, 1000, 800, 1200, 600, 800, 800, 1200, 0, 0, 2500, 0
        If grdCotas.ListItems.Count >= 1 Then
            txtValor = grdCotas.Colunas(7).Soma
        End If
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
'    If Not LiberaImpressaoDam(Nvl(lstObrig.SelectedItem.SubItems(15), 0)) Then Exit Sub
     With Rpt
            
               If Not .DefinirArquivo(Bdados, App.Path + "\TermoParcela.rpt") Then Exit Sub
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
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
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

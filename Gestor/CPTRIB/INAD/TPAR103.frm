VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TPAR103"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1695
      Left            =   75
      TabIndex        =   24
      Top             =   5625
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   2990
      Altura          =   1905
      Caption         =   " Dados da Parcela"
      CorTexto        =   0
      CorFaixa        =   8421504
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Enabled         =   0   'False
         Height          =   975
         Left            =   45
         TabIndex        =   27
         Top             =   285
         Width           =   3000
         Begin VTOcx.txtVISUAL txtDocumento 
            Height          =   300
            Left            =   195
            TabIndex        =   28
            Top             =   195
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   529
            Caption         =   "Nº Documento"
            Text            =   ""
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtNumParcelamento 
            Height          =   300
            Left            =   0
            TabIndex        =   29
            Top             =   615
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   529
            Caption         =   "Nº Parcelamento"
            Text            =   ""
            Requerido       =   0   'False
            RetirarMascara  =   0   'False
            AutoTAB         =   -1  'True
         End
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8415
         TabIndex        =   26
         Top             =   375
         Width           =   1455
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Nova Parcela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6525
         TabIndex        =   25
         Top             =   375
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   5895
         TabIndex        =   6
         Tag             =   "Status"
         Top             =   690
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtParcela 
         Height          =   300
         Left            =   5820
         TabIndex        =   7
         Tag             =   "Parcela"
         Top             =   1095
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         Caption         =   "Parcela"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtdataVencimento 
         Height          =   300
         Left            =   3045
         TabIndex        =   5
         Tag             =   "Vencimento"
         Top             =   1095
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   529
         Caption         =   "Vencimento"
         Text            =   ""
         Formato         =   0
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   300
         Left            =   855
         TabIndex        =   4
         Tag             =   "Período"
         Top             =   1290
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Período"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   300
         Left            =   8130
         TabIndex        =   8
         Tag             =   "Valor"
         Top             =   1095
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         Caption         =   "Valor"
         Text            =   ""
         Formato         =   5
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VB.Shape Shape1 
         Height          =   360
         Left            =   6465
         Top             =   300
         Width           =   3510
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   17
      Top             =   7320
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   4920
         TabIndex        =   30
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   135
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   661
         Caption         =   "Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   6030
         TabIndex        =   9
         Top             =   135
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9270
         TabIndex        =   12
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
         Left            =   8130
         TabIndex        =   11
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
      Height          =   2205
      Left            =   60
      TabIndex        =   16
      Top             =   3660
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   3889
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
      TabIndex        =   21
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR103.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   13
      Top             =   -420
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   1725
      Index           =   0
      Left            =   60
      TabIndex        =   14
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   22
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
         TabIndex        =   23
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TPAR103.frx":2123
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1770
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   18
      Top             =   9000
      Visible         =   0   'False
      Width           =   795
   End
   Begin VTOcx.grdVISUAL lstParc 
      Height          =   1485
      Left            =   60
      TabIndex        =   31
      Top             =   2400
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   2619
      Caption         =   "Parcelamentos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
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
Attribute VB_Name = "TPAR103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Obrig As New Obrigacao
Dim String_Taxas As String
Dim Total_Taxas As Double
Public BMudou As Boolean
Private Rpt As New VSRelatorio
Private Sub cmdCancela_Click()
    On Error Resume Next
    Edita.LimpaCampos Me
    lstParc.ListItems.Clear
    grdCotas.ListItems.Clear
    
    
    lstParc.Mensagem = ""
    grdCotas.Mensagem = ""
    
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

Private Sub cmdExcluir_Click()
    If grdCotas.ListItems.Count >= 1 Then
        If Confirma("Confirma a exclusão?") = True Then
            If Bdados.DeletaDados("tab_cotas_parcelamento", "TCP_TPA_COD_PARCELAMENTO = '" & lstParc.SelectedItem & "' and TCP_NUM_COTA = '" & grdCotas.SelectedItem & "'") Then
                Util.Avisa "Operação realizada com sucesso."
                lstParc_Click
            End If
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Parcelamento As String
    Dim Selecao As String
    
    If txtIm = "" And txtImovel = "" Then
        Selecao = " {TAB_PARCELAMENTO.TPA_NUM_PARCELAMENTO} > 0"
    Else
    Selecao = " {TAB_PARCELAMENTO.TPA_NUM_PARCELAMENTO} > 0"
    
    Selecao = Selecao & " and {TAB_PARCELAMENTO.TPA_NUM_PARCELAMENTO} = " & lstParc.SelectedItem
    Selecao = Selecao & " and {TAB_COTAS_PARCELAMENTO.TCP_INSCRICAO} = '" & lstParc.SelectedItem.SubItems(1) & "'"
    End If
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TParcelamentos.rpt") Then Exit Sub
        .Formulas "VT_ENDERECO", txtEndereco
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Selecao = Selecao
        .Visualizar
    End With
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
    'grdOriginal.ListItems.Clear
    lstParc.Mensagem = ""
    grdCotas.Mensagem = ""
    lstParc.Mensagem = ""
'    If cboImposto.ListIndex = -1 Then
'        Util.Avisa "Selecione o imposto."
'        cboImposto.SetFocus
'        Exit Sub
'    End If
    Screen.MousePointer = 11
    sql = "SELECT TPA_NUM_PARCELAMENTO as Parcelamento,TPA_INSCRICAO AS Inscricao,TPA_PERIODO AS Periodo,TIP_SIGLA_IMPOSTO as Tributo," & _
        " TPA_DATA_FINANCIAMENTO as Data, " & _
        "" & _
        " TPA_VALOR_PARCELADO as Valor_Parcelado," & _
        " TPA_NUM_COTAS as Cotas,TPA_STATUS_PARCELAMENTO as Situacão,TPA_PERIODO_INICIAL,TPA_PERIODO_FINAL,tpa_tipo_inscricao as Tipo " & _
        " FROM TAB_PARCELAMENTO,TAB_IMPOSTO WHERE  TPA_TIP_COD_IMPOSTO=TIP_COD_IMPOSTO AND TPA_STATUS_PARCELAMENTO <> " & stsParcelamentoCancelado
    If Trim(cboImposto) <> "" Then sql = sql & " and TPA_TIP_COD_IMPOSTO='" & cboImposto.Coluna(0).VALOR & "'"
    If Trim(txtIm) <> "" Then
        sql = sql & " and (TPA_INSCRICAO='" & Trim(txtIm) & "' and TPA_TIPO_INSCRICAO =2)"
    ElseIf Trim(txtImovel) <> "" Then
        sql = sql & " and (TPA_INSCRICAO='" & Trim(txtImovel) & "' and TPA_TIPO_INSCRICAO =1)"
    End If
    lstParc.Preencher Bdados, sql, 1200, 1200, 1000, 1200, 1200, 1500, 800, 0, 0, 0
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

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Dim Tipo As TipoInscricaoObrigacao
    Dim Parcela As Integer
    Dim Conta As New ContaCorrente
    Dim ContaDocumento As String
    
    cboImposto.Tag = ""
    
'    If Valida_Valor = False Then Exit Sub
    txtDocumento.Tag = ""
    txtNumParcelamento.Tag = ""
    If CriticaCampos(Me) = False Then Exit Sub
    If lstParc.SelectedItem.SubItems(10) = etiImovel Then
        Tipo = etiImovel
    Else
        Tipo = etiContribuinte
    End If
    Condicao = "TCP_TPA_COD_PARCELAMENTO = '" & lstParc.SelectedItem & "' and TCP_NUM_COTA = '" & grdCotas.SelectedItem & "'"
 
    If OptTipo(0).Value = True Then
        'Checo se já existe uma parcela igual...
        For Parcela = 1 To grdCotas.ListItems.Count
            If grdCotas.ListItems(Parcela).SubItems(5) = txtParcela Then
                Util.Avisa "Nº de Parcela inválida."
                Exit Sub
            End If
        Next
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            ContaDocumento = Conta.correlativo("TRIB", "64", "COTAS PARCELAS")
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            ContaDocumento = Conta.GeraCodPagamento(64)
        End If
        Valores = Bdados.PreparaValor(txtPeriodo, Bdados.Converte(Date, TCDataHora), txtdataVencimento, cboStatus.Coluna(1).VALOR, txtParcela, txtValor, ContaDocumento, lstParc.SelectedItem, lstParc.SelectedItem.SubItems(1), Tipo, 0)
        Campos = "TCP_PERIODO,TCP_DATA_GERACAO,TCP_DATA_VENCIMENTO,TCP_STATUS_OBRIGACAO_PARCELA,TCp_NUM_PARCELA,TCP_VALOR_PARCELA,TCP_NUM_COTA,TCP_TPA_COD_PARCELAMENTO,TCP_INSCRICAO,tcp_tipo_inscricao,TCp_VALOR_JUROS"
        If Bdados.InsereDados("tab_cotas_parcelamento", Valores, Campos) Then
            'GRAVO NA TAB_OBRIGACAO_CONTRIBUINTE DE ACORDO COM O NOVO PROCESSO DO PARCELAMENTO...
            Dim Insc As String
            
            If txtIm <> "" Then
                Tipo = etiContribuinte
            Else
                Tipo = etiImovel
            End If
            
            Obrig.CriaObrigacao Imposto.BuscaCodImposto(lstParc.SelectedItem.SubItems(3)), txtPeriodo, txtPeriodo, lstParc.SelectedItem.SubItems(1), txtValor, etsCreditoOriginalAberto, etsCriaNova, txtdataVencimento, , , , , , , , 0, , Tipo
            Util.Avisa "Operação concluida com sucesso."
            lstParc_Click
            Limpa
        End If
    Else
        Valores = Bdados.PreparaValor(txtPeriodo, Bdados.Converte(Date, TCDataHora), txtdataVencimento, cboStatus.Coluna(1).VALOR, txtParcela, txtValor, txtDocumento, lstParc.SelectedItem)
        Campos = "TCP_PERIODO,TCP_DATA_GERACAO,TCP_DATA_VENCIMENTO,TCP_STATUS_OBRIGACAO_PARCELA,TCp_NUM_PARCELA,TCP_VALOR_PARCELA,TCP_NUM_COTA"
        If Bdados.AtualizaDados("tab_cotas_parcelamento", Valores, Campos, Condicao) Then
        'ALTERO NA TAB_OBRIGACAO_CONTRIBUINTE DE ACORDO COM O NOVO PROCESSO DE PARCELAMENTO...
            If Not Bdados.AbreTabela("Select toc_cod_obrigacao from tab_obrigacao_contribuinte where toc_cod_obrigacao = '" & grdCotas.SelectedItem & "'") Then
                Exit Sub
            End If
            Campos = "TOC_PERIODO,TOC_DATA_VENCIMENTO,"
            Campos = Campos & "TOC_VALOR_OBRIGACAO,TOC_STATUS_OBRIGACAO"
            Valores = Bdados.PreparaValor(Bdados.Converte(txtPeriodo, tctexto), Bdados.Converte(txtdataVencimento, TCDataHora), Bdados.Converte(txtValor, TCMonetario), Bdados.Converte(cboStatus.Coluna(1).VALOR, tctexto))
            Condicao = "TOC_COD_OBRIGACAO = " & Bdados.Converte(txtDocumento, tctexto)
            Bdados.GravaDados "TAB_OBRIGACAO_CONTRIBUINTE", Valores, Campos, Condicao
            Util.Avisa "Operação concluida com sucesso."
            lstParc_Click
            Limpa
        End If
    End If
End Sub
Private Sub Limpa()
   txtPeriodo = ""
    txtdataVencimento = ""
    cboStatus.ListIndex = -1
    txtParcela = ""
    txtValor = ""
    txtDocumento = ""
    txtNumParcelamento = ""
    
End Sub
Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboImposto.Preencher Bdados, "Select  tip_cod_imposto,TIP_sigla_IMPOSTO  " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto,tip_nome_imposto From TAB_IMPOSTO order by TIP_sigla_IMPOSTO asc", 1
    AtualizaCabecalho lstParc
    'Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
End Sub


Private Sub lstParc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstParc, ColumnHeader
End Sub

Private Sub grdCotas_DblClick()
    Dim sql As String
    Dim rs As VSRecordset
    
    If grdCotas.ListItems.Count >= 1 Then
        txtDocumento = grdCotas.SelectedItem
        txtDocumento.Enabled = False
        txtNumParcelamento = lstParc.SelectedItem
        txtNumParcelamento.Enabled = False
        txtPeriodo = grdCotas.SelectedItem.SubItems(3)
        txtdataVencimento = grdCotas.SelectedItem.SubItems(4)
        txtParcela = grdCotas.SelectedItem.SubItems(5)
        txtValor = grdCotas.SelectedItem.SubItems(6)
        BMudou = False
        cboStatus.SetarLinha grdCotas.SelectedItem.SubItems(12), 1
    End If
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
    If lstParc.ListItems.Count > 0 Then
        'Obrig.CarregaListaObrigacao grdOriginal, txtim, , , , lstParc.SelectedItem
        sql = " Select TCp_NUM_COTA AS Documento,tcp_inscricao as Inscrição,"
        sql = sql & " TIP_SIGLA_IMPOSTO AS Tributo,"
        sql = sql & " TPA_PERIODO AS Periodo,"
        sql = sql & " TCp_DATA_VENCIMENTO AS Vencimento,TCp_NUM_PARCELA as Cota,"
        sql = sql & " TCp_VALOR_PARCELA As Valor, TCp_VALOR_JUROS As Juros, TCp_VALOR_PARCELA"
        sql = sql & " + TCp_VALOR_JUROS as Total,tip_cod_imposto as Imposto,"
        sql = sql & " tip_nome_imposto as Descrição,TGE_NOME as SITUACAO,TCP_STATUS_OBRIGACAO_PARCELA,TCP_DATA_GERACAO"
        sql = sql & " From tab_parcelamento, tab_cotas_parcelamento, tab_imposto,VIS_STATUS_OBRIGACAO"
        sql = sql & " where  tpa_num_parcelamento = TCp_TPA_COD_PARCELAMENTO and TCP_STATUS_OBRIGACAO_PARCELA = TGE_CODIGO and"
        sql = sql & " TPA_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
        sql = sql & " AND TCP_TPA_COD_PARCELAMENTO =   '" & lstParc.SelectedItem & "'"
        txtIm = ""
        txtImovel = ""
        If lstParc.SelectedItem.SubItems(7) = etiContribuinte Then
            txtIm = Trim(lstParc.SelectedItem.SubItems(1))
            txtIm_LostFocus
        Else
            txtImovel = Trim(lstParc.SelectedItem.SubItems(1))
            txtImovel_LostFocus
        End If
        grdCotas.Preencher Bdados, sql, 0, 1200, 1000, 800, 1200, 600, 800, 800, 1200, 0, 0, 2500, 0
        If grdCotas.ListItems.Count > 0 Then grdCotas.Mensagem = "Total Parcelamento: R$" & Format(grdCotas.Colunas(7).Soma, Const_Monetario) & " x Acréscimo na dívida original: R$" & Format(grdCotas.Colunas(9).Soma - grdCotas.Colunas(7).Soma, Const_Monetario)
    End If
End Sub

Private Sub lstParc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then lstParc_Click
End Sub

Private Sub lstParc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not lstParc.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuExcluir.Caption = "Cancelar parcelamento nº " & lstParc.SelectedItem
            mnuReimprime.Caption = "Reimprimir termo de parcelamento nº " & lstParc.SelectedItem
            Me.PopupMenu mnumenu
        End If
    End If
End Sub

Private Sub mnuExcluir_Click()
    TPAR201.Tag = Right(mnuExcluir.Caption, 8)
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
        Cobranca.ImprimeDam Rpt, lstParc.SelectedItem, lstParc.SelectedItem.SubItems(1), txtRazao, CpfCgc, txtEndereco, txtIm, txtEndereco, .SubItems(9), .SubItems(2), _
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
                .Formulas "Imposto ", CStr(cboImposto.Coluna(2).VALOR)
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

Private Sub OptTipo_Click(Index As Integer)
    If Index = 0 Then
        txtParcela.Enabled = True
    Else
        txtParcela.Enabled = False
    End If
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

Public Function Valida_Valor() As Boolean
    Dim i As Integer
    Dim Soma As Double
    Valida_Valor = True
    For i = 1 To grdCotas.ListItems.Count
        Soma = Soma + CDbl(grdCotas.ListItems(i).SubItems(8))
    Next
    
    If CDbl(txtValor) + Soma > lstParc.SelectedItem.SubItems(6) And BMudou = True Then
        Util.Avisa "Valor inválido: " & vbCrLf & "Valor do Débito :" & lstParc.SelectedItem.SubItems(6) & vbCrLf & "Valor total da(s) parcela :" & Format(CDbl(txtValor) + Soma, "###,###,###,##0.00")
        Valida_Valor = False
    End If

End Function

Private Sub txtValor_Change()
    BMudou = True
End Sub

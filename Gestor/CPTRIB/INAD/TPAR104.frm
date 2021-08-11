VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TPAR104"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1138
      Icone           =   "TPAR104.frx":0000
   End
   Begin ActiveTabs.SSActiveTabs tabCND 
      Height          =   2970
      Left            =   15
      TabIndex        =   18
      Top             =   2280
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   5239
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TPAR104.frx":031A
      Images          =   "TPAR104.frx":03D0
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2550
         Index           =   0
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   4498
         _Version        =   131082
         TabGuid         =   "TPAR104.frx":1069
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   2370
            Left            =   0
            TabIndex        =   24
            Top             =   15
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   4180
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
               Left            =   390
               TabIndex        =   34
               Top             =   360
               Width           =   3000
               Begin VTOcx.txtVISUAL txtDocumento 
                  Height          =   300
                  Left            =   195
                  TabIndex        =   35
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
                  TabIndex        =   36
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
            Begin VTOcx.cmdVISUAL cmdSalvar 
               Height          =   375
               Left            =   6300
               TabIndex        =   33
               Top             =   1935
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   661
               Caption         =   "Salvar"
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdExcluir 
               Height          =   375
               Left            =   7380
               TabIndex        =   32
               Top             =   1935
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   661
               Caption         =   "Excluir"
               Acao            =   2
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.txtVISUAL txtValor 
               Height          =   300
               Left            =   6555
               TabIndex        =   31
               Tag             =   "Valor"
               Top             =   1500
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
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   300
               Left            =   1185
               TabIndex        =   30
               Tag             =   "Período"
               Top             =   1395
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   529
               Caption         =   "Período"
               Text            =   ""
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtdataVencimento 
               Height          =   300
               Left            =   825
               TabIndex        =   29
               Tag             =   "Vencimento"
               Top             =   1830
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
            Begin VTOcx.txtVISUAL txtParcela 
               Height          =   300
               Left            =   4245
               TabIndex        =   28
               Tag             =   "Parcela"
               Top             =   1500
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   529
               Caption         =   "Parcela"
               Text            =   ""
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboVISUAL1 
               Height          =   315
               Left            =   4320
               TabIndex        =   27
               Tag             =   "Status"
               Top             =   1095
               Width           =   4110
               _ExtentX        =   7250
               _ExtentY        =   556
               Caption         =   "Status"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
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
               Left            =   4920
               TabIndex        =   26
               Top             =   585
               Value           =   -1  'True
               Width           =   1755
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
               Left            =   6810
               TabIndex        =   25
               Top             =   585
               Width           =   1455
            End
            Begin VB.Shape Shape1 
               Height          =   360
               Left            =   4860
               Top             =   510
               Width           =   3510
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2550
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   4498
         _Version        =   131082
         TabGuid         =   "TPAR104.frx":1091
         Begin VTOcx.grdVISUAL lstObrig 
            Height          =   2505
            Left            =   15
            TabIndex        =   22
            Top             =   45
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   4419
            Caption         =   "Débitos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2550
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   4498
         _Version        =   131082
         TabGuid         =   "TPAR104.frx":10B9
         Begin VTOcx.grdVISUAL grdCotas 
            Height          =   2505
            Left            =   15
            TabIndex        =   23
            Top             =   45
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   4419
            Caption         =   "Parcelas"
            CorTitulo       =   32768
            CorCaption      =   -2147483634
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   5295
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   4800
         TabIndex        =   37
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   5910
         TabIndex        =   17
         Top             =   120
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8055
         TabIndex        =   3
         Top             =   120
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
         Left            =   6915
         TabIndex        =   2
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR104.frx":10E1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
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
      Height          =   1650
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   585
      Width           =   9255
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   825
         TabIndex        =   7
         Top             =   540
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   5925
         TabIndex        =   1
         Top             =   1260
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   1125
         TabIndex        =   10
         Top             =   900
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   705
         TabIndex        =   14
         Top             =   1260
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
      Begin VTOcx.txtVISUAL txtExercicioFinal 
         Height          =   300
         Left            =   3210
         TabIndex        =   15
         Top             =   1260
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
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Top             =   180
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
         Left            =   3645
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VB.Label LblObrigacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   9090
         TabIndex        =   38
         Top             =   225
         Width           =   45
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   8
         Top             =   1590
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   4
      Top             =   90
      Width           =   375
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4260
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   11
      Top             =   5490
      Visible         =   0   'False
      Width           =   795
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1620
      Left            =   45
      TabIndex        =   13
      Top             =   7320
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   2858
      Caption         =   "Taxas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "Geral"
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "Reimprime"
         Index           =   0
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Parcelamento Espontâneo"
         Index           =   2
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Parcelamento de Ofício"
         Index           =   3
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Cancelar"
         Index           =   5
      End
   End
End
Attribute VB_Name = "TPAR104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obrig As New Obrigacao
Dim Conta As New ContaCorrente
Dim Cobranca As New VSCobranca
    
Dim NovoJuro As String
Dim NovaMulta As String
Dim NovaData As String

Dim Rpt As New VSRelatorio

 Public String_Taxas    As String
 Public Total_Taxas     As Double
 Public BMudou As Boolean



Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Dim Inscri As String
    Set Obrig = New Obrigacao
    
    
    
    'If Trim(txtIm) <> "" Then Conta.ExecutaAtualizacao txtIm
    If Not Obrig.MostraObrigacaoGerada(lstObrig, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)), Inscri, , CInt(CboStatus.Coluna(1).Valor), , , txtExercicioInicial, txtExercicioFinal, , txtImovel) Then
        Avisa "Nenhum registro encontrado."
        txtImovel.SetFocus
    End If
End Sub

Private Sub cmdCancela_Click()
    On Error Resume Next
    Edita.LimpaCampos Me
    lstObrig.ListItems.Clear
    grdCotas.ListItems.Clear
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub





Private Sub cmdExcluir_Click()
    Dim condicao As String
    
   If grdCotas.ListItems.Count >= 1 Then
        condicao = "TCO_TOC_COD_OBRIGACAO = '" & lstObrig.SelectedItem & "' and TCO_COD_OBRIGACAO_PARCELA = '" & grdCotas.SelectedItem & "'"
        If Confirma("Confirma a exclusão?") = True Then
            If Bdados.DeletaDados("tab_cotas_OBRIGACAO", condicao) Then
                Util.Avisa "Operação realizada com sucesso."
                lstObrig_DblClick
            End If
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Parcelamento As String
    Dim Selecao As String
    Dim PeriodoInicio As String
    Dim PeriodoFim As String
    
    
    Selecao = " {TAB_COTAS_OBRIGACAO.TCO_TOC_COD_OBRIGACAO} > 0"
    
    If grdCotas.ListItems.Count >= 1 Then
        If LblObrigacao.Caption <> "" Then
            Selecao = Selecao + "  and {TAB_COTAS_OBRIGACAO.TCO_TOC_COD_OBRIGACAO}= " & lstObrig.SelectedItem
        End If
    End If
    If Confirma("Deseja informar período?", "Memsagem") = True Then
        PeriodoInicio = Entrada("Informe o período inicial.", "Aviso")
        PeriodoFim = Entrada("Informe o período final.", "Aviso")
        
        
        If Trim(PeriodoInicio) <> "" And Trim(PeriodoFim) <> "" Then
            Selecao = Selecao & " and  {TAB_COTAS_OBRIGACAO.TCO_PERIODO} >= " & Trim(PeriodoInicio) & " and {TAB_COTAS_OBRIGACAO.TCO_PERIODO} <= " & Trim(PeriodoFim)
        ElseIf Trim(PeriodoInicio) <> "" And Trim(PeriodoFim) = "" Then
            Selecao = Selecao & " and  {TAB_COTAS_OBRIGACAO.TCO_PERIODO} >= " & Trim(PeriodoInicio) & " and {TAB_COTAS_OBRIGACAO.TCO_PERIODO} <= " & Trim(PeriodoInicio)
        End If
    End If
    '{TAB_COTAS_OBRIGACAO.TCO_PERIODO}
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path + "\TParcelamentosObrigacao.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Selecao = Selecao
        .Visualizar
    End With

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
 Dim Valores As String
    Dim campos As String
    Dim condicao As String
    Dim Tipo As TipoInscricaoObrigacao
    Dim Parcela As Integer
    Dim CodigoParcela As String
    Dim Conta As New ContaCorrente
    Dim CodDocumento As String
    
    If CriticaCampos(Me) = False Then Exit Sub
    
    If Valida_Valor = False Then Exit Sub
    If grdCotas.ListItems.Count >= 1 Then
        If OptTipo(0).Value = True Then
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                CodigoParcela = Conta.correlativo("TRIB", "57", "COTAS LANCAMENTO")
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                CodigoParcela = Conta.GeraCodPagamento(57)
            End If
        Else
            CodigoParcela = grdCotas.SelectedItem
        End If
    Else
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            CodigoParcela = Conta.correlativo("TRIB", "57", "COTAS LANCAMENTO")
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            CodigoParcela = Conta.GeraCodPagamento(57)
        End If
    End If
    
    condicao = "TCO_TOC_COD_OBRIGACAO = '" & lstObrig.SelectedItem & "' and TCO_COD_OBRIGACAO_PARCELA = '" & CodigoParcela & "'"
    
    If OptTipo(0).Value = True Then
        'Checo se já existe uma parcela igual...
        For Parcela = 1 To grdCotas.ListItems.Count
            If grdCotas.ListItems(Parcela).SubItems(5) = txtParcela Then
                Util.Avisa "Nº de Parcela inválida."
                Exit Sub
            End If
        Next
        Valores = Bdados.PreparaValor(txtPeriodo, Bdados.Converte(Date, TCDataHora), txtdataVencimento, cboVISUAL1.Coluna(1).Valor, txtParcela, txtValor, CodigoParcela, lstObrig.SelectedItem, lstObrig.SelectedItem.SubItems(1), 0, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)))
        campos = "TCo_PERIODO,TCo_DATA_GERACAO,TCo_DATA_VENCIMENTO,TCo_STATUS_OBRIGACAO_PARCELA,TCO_NUM_PARCELA,TCo_VALOR_PARCELA,TCO_COD_OBRIGACAO_PARCELA ,TCO_TOC_COD_OBRIGACAO,TCo_INSCRICAO,TCo_VALOR_JUROS,TCO_TIP_COD_IMPOSTO"
        If Bdados.InsereDados("tab_cotas_OBRIGACAO", Valores, campos) Then
            Util.Avisa "Operação concluida com sucesso."
            lstObrig_DblClick
            Limpa
        End If
    Else
        Valores = Bdados.PreparaValor(txtPeriodo, Bdados.Converte(Date, TCDataHora), txtdataVencimento, cboVISUAL1.Coluna(1).Valor, txtParcela, txtValor, txtDocumento, 0, Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)))
        campos = "TCo_PERIODO,TCo_DATA_GERACAO,TCo_DATA_VENCIMENTO,TCo_STATUS_OBRIGACAO_PARCELA,TCO_NUM_PARCELA,TCo_VALOR_PARCELA,TCO_COD_OBRIGACAO_PARCELA ,TCo_VALOR_JUROS,TCO_TIP_COD_IMPOSTO"
        If Bdados.AtualizaDados("tab_cotas_OBRIGACAO", Valores, campos, condicao) Then
            Util.Avisa "Operação concluida com sucesso."
            lstObrig_DblClick
            Limpa
        End If
    End If
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Activate()
    If Left(Trim(Me.Tag), 1) = "C" Then
        
        
    ElseIf Left(Trim(Me.Tag), 1) = "I" Then
        txtImovel = Mid(Me.Tag, 2)
        txtImovel_LostFocus
    ElseIf Len(Trim(Me.Tag)) > 0 Then
        
        
    End If
    
    
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    CboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    cboVISUAL1.PreencherGeral Bdados, "STATUS OBRIGACAO"
    
    
End Sub




Private Sub grdCotas_DblClick()
    If grdCotas.ListItems.Count >= 1 Then
        txtDocumento = grdCotas.SelectedItem
        txtDocumento.Enabled = False
        txtNumParcelamento = lstObrig.SelectedItem
        txtNumParcelamento.Enabled = False
        txtNumParcelamento.Enabled = False
        txtPeriodo = grdCotas.SelectedItem.SubItems(3)
        txtdataVencimento = grdCotas.SelectedItem.SubItems(4)
        cboVISUAL1.SetarLinha grdCotas.SelectedItem.SubItems(12), 1
        txtParcela = grdCotas.SelectedItem.SubItems(5)
        txtValor = grdCotas.SelectedItem.SubItems(8)
        BMudou = False
        tabCND.Tabs(3).Selected = True
    End If
End Sub

Private Sub lstObrig_DblClick()
    Dim Sql As String
    Sql = " Select TCO_COD_OBRIGACAo_PARCELA AS Documento,tco_inscricao as Inscrição,"
    Sql = Sql & " TIP_SIGLA_IMPOSTO AS Tributo,tco_periodo as Ano,"
    Sql = Sql & " TCO_DATA_VENCIMENTO AS Vencimento,TCO_NUM_PARCELA as Cota,"
    Sql = Sql & " TCO_VALOR_PARCELA as Valor,TCO_VALOR_JUROS as Juros ,"
    Sql = Sql & " TCO_VALOR_PARCELA + TCO_VALOR_JUROS as Total,"
    Sql = Sql & " tip_cod_imposto as Imposto,tip_nome_imposto as Descrição,Tge_Nome as Situação,TCo_STATUS_OBRIGACAO_PARCELA as [Código Status]"
    Sql = Sql & " From TAB_COTAS_OBRIGACAO,Tab_imposto,vis_status_obrigacao"
    Sql = Sql & " Where"
    Sql = Sql & " TCO_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
    Sql = Sql & "  and tge_codigo = TCO_STATUS_OBRIGACAO_PARCELA"
    Sql = Sql & " AND TCO_TOC_COD_OBRIGACAO =  '" & lstObrig.SelectedItem & "' order by TCO_NUM_PARCELA"
    
    Limpa
    If grdCotas.Preencher(Bdados, Sql) Then
        tabCND.Tabs(2).Selected = True
        OptTipo(0).Enabled = True
        OptTipo(1).Enabled = True
    Else
        tabCND.Tabs(3).Selected = True
        OptTipo(0).Value = True
        OptTipo(0).Enabled = False
        OptTipo(1).Enabled = False
        txtPeriodo.SetFocus
    End If
    LblObrigacao = "Nº do Documento " & lstObrig.SelectedItem
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

Private Sub tabCND_Click()
    
    If tabCND.Tabs(1).Selected = True Then
        LblObrigacao.Caption = ""
    End If
End Sub

Private Sub txtImovel_LostFocus()
    Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            'txtIm.SetFocus
        End If
    End If
End Sub
Private Sub Limpa()
    txtPeriodo = ""
    txtdataVencimento = ""
    CboStatus.ListIndex = -1
    txtParcela = ""
    txtValor = ""
    txtDocumento = ""
    txtNumParcelamento = ""
    cboVISUAL1.ListIndex = -1
End Sub
Public Function Valida_Valor() As Boolean
    Dim i As Integer
    Dim Soma As Double
    Valida_Valor = True
    For i = 1 To grdCotas.ListItems.Count
        Soma = Soma + CDbl(grdCotas.ListItems(i).SubItems(8))
    Next
    
    If CDbl(txtValor) + Soma > lstObrig.SelectedItem.SubItems(6) And BMudou = True Then
        Util.Avisa "Valor inválido: " & vbCrLf & "Valor do Débito :" & lstObrig.SelectedItem.SubItems(6) & vbCrLf & "Valor total da(s) parcela :" & Format(CDbl(txtValor) + Soma, "###,###,###,##0.00")
        Valida_Valor = False
    End If
    
End Function

Private Sub txtValor_Change()
    BMudou = True
End Sub

VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TDEC101.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs TabDec 
      Height          =   4215
      Left            =   30
      TabIndex        =   18
      Top             =   3480
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   7435
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TDEC101.frx":0342
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3825
         Index           =   1
         Left            =   -99969
         TabIndex        =   19
         Top             =   30
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   6747
         _Version        =   131082
         TabGuid         =   "TDEC101.frx":03E8
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   4635
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   8176
            Altura          =   1905
            Caption         =   " Apuracão das Entradas - Apuração de imposto sobre as notas recebidas"
            CorFaixa        =   16711680
            Ocultavel       =   0   'False
            Begin VTOcx.fraFUTURO fraFUTURO1 
               Height          =   5325
               Index           =   0
               Left            =   0
               TabIndex        =   31
               Top             =   270
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   9393
               Caption         =   "Apuração do Imposto"
               Descricao       =   "Apuração de imposto sobre as notas emitidas"
               corFaixa        =   16384
               corFundo        =   14737632
               corTexto        =   16384
               Icone           =   "TDEC101.frx":0410
               Ocultavel       =   0   'False
               Altura          =   2000
               Begin VTOcx.fraVISUAL fraNormal 
                  Height          =   1350
                  Index           =   6
                  Left            =   0
                  TabIndex        =   33
                  Top             =   720
                  Width           =   11175
                  _ExtentX        =   19711
                  _ExtentY        =   2381
                  Altura          =   1905
                  Caption         =   " Substituicão Tributária"
                  CorTexto        =   0
                  CorFaixa        =   16711680
                  Ocultavel       =   0   'False
                  Begin VTOcx.fraVISUAL fraNormal 
                     Height          =   990
                     Index           =   1
                     Left            =   15
                     TabIndex        =   34
                     Top             =   315
                     Width           =   5685
                     _ExtentX        =   10028
                     _ExtentY        =   1746
                     Altura          =   1905
                     Caption         =   " Total de ISSQN Retido de Prestadores"
                     CorTexto        =   0
                     CorFaixa        =   16711680
                     Ocultavel       =   0   'False
                     Begin VTOcx.txtVISUAL txtItemDecl 
                        Height          =   315
                        Index           =   13
                        Left            =   3840
                        TabIndex        =   37
                        Top             =   600
                        Width           =   1725
                        _ExtentX        =   3043
                        _ExtentY        =   556
                        Caption         =   "Total"
                        Text            =   ""
                        Enabled         =   0   'False
                        Formato         =   5
                        Restricao       =   3
                        AlinhamentoTexto=   1
                        CorFundo        =   14737632
                     End
                     Begin VTOcx.txtVISUAL txtItemDecl 
                        Height          =   315
                        Index           =   10
                        Left            =   1170
                        TabIndex        =   36
                        Top             =   300
                        Width           =   2355
                        _ExtentX        =   4154
                        _ExtentY        =   556
                        Caption         =   "do Município"
                        Text            =   ""
                        Enabled         =   0   'False
                        Formato         =   5
                        Restricao       =   3
                        AlinhamentoTexto=   1
                        CorFundo        =   14737632
                     End
                     Begin VTOcx.txtVISUAL txtItemDecl 
                        Height          =   315
                        Index           =   11
                        Left            =   750
                        TabIndex        =   35
                        Top             =   630
                        Width           =   2775
                        _ExtentX        =   4895
                        _ExtentY        =   556
                        Caption         =   "Fora do Município"
                        Text            =   ""
                        Formato         =   5
                        Restricao       =   3
                        AlinhamentoTexto=   1
                        CorFundo        =   14737632
                     End
                  End
               End
               Begin VTOcx.grdVISUAL Grdtaxas 
                  Height          =   435
                  Left            =   60
                  TabIndex        =   32
                  Top             =   5565
                  Width           =   11190
                  _ExtentX        =   19738
                  _ExtentY        =   767
                  Caption         =   "Taxas"
                  OcultarRodape   =   -1  'True
                  CheckBox        =   -1  'True
                  Ordenavel       =   0   'False
               End
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3825
         Index           =   0
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   6747
         _Version        =   131082
         TabGuid         =   "TDEC101.frx":0CEA
         Begin VTOcx.grdVISUAL grdEntrada 
            Height          =   2190
            Left            =   0
            TabIndex        =   26
            Top             =   1560
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   3863
            CorBorda        =   16711680
            Caption         =   "Notas recebidas"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.fraVISUAL fraNormal 
            Height          =   1425
            Index           =   3
            Left            =   0
            TabIndex        =   27
            Top             =   60
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   2514
            Altura          =   1905
            Caption         =   " Nota Fiscal Recebida"
            CorTexto        =   0
            CorFaixa        =   16711680
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtAidf 
               Height          =   525
               Left            =   2610
               TabIndex        =   7
               Top             =   300
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   926
               Caption         =   "No. AIDF"
               Text            =   ""
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTNumNota 
               Height          =   525
               Left            =   90
               TabIndex        =   3
               Top             =   300
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   926
               Caption         =   "Nº Nota Fiscal"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTEmissao 
               Height          =   525
               Left            =   1380
               TabIndex        =   6
               Top             =   300
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   926
               Caption         =   "Data Emissão"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTValor 
               Height          =   525
               Left            =   90
               TabIndex        =   8
               Top             =   840
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   926
               Caption         =   "Vlr da Nota"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTImpostoDevido 
               Height          =   525
               Left            =   5100
               TabIndex        =   30
               Top             =   840
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   926
               Caption         =   "ISS devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTSaldo 
               Height          =   525
               Left            =   3660
               TabIndex        =   29
               Top             =   840
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   926
               Caption         =   "Base de Calculo"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.cmdVISUAL cmdAdicionarNotaST 
               Height          =   375
               Left            =   7020
               TabIndex        =   11
               ToolTipText     =   "Adicionar"
               Top             =   420
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   661
               Caption         =   "&Incluir"
               Acao            =   1
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.txtVISUAL txtSTIcms 
               Height          =   525
               Left            =   1380
               TabIndex        =   9
               Top             =   840
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   926
               Caption         =   "Vlr Sujeito ICMS"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTAliq 
               Height          =   525
               Left            =   2910
               TabIndex        =   5
               Top             =   840
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   926
               Caption         =   "Aliq(%)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTImpostoRetido 
               Height          =   525
               Left            =   6120
               TabIndex        =   10
               Top             =   840
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   926
               Caption         =   "ISS retido"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTSaldoDevedor 
               Height          =   525
               Left            =   7050
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   840
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   926
               Caption         =   "Saldo Devedor"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
               CorFundo        =   14737632
            End
            Begin VTOcx.txtVISUAL txtSTInscricao 
               Height          =   525
               Left            =   3660
               TabIndex        =   4
               Top             =   300
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   926
               Caption         =   "CPF/CNPJ Prestador"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorFundo        =   14737632
               AgruparValores  =   0   'False
               RetirarMascara  =   0   'False
            End
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   17
      Top             =   7860
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdFinaliza 
         Height          =   375
         Left            =   6720
         TabIndex        =   13
         Top             =   90
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   661
         Caption         =   "Finalizar Declaracão"
         Acao            =   1
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmLimpar 
         Height          =   375
         Left            =   8820
         TabIndex        =   14
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4770
         TabIndex        =   12
         Top             =   90
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Salvar Declaracão"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10110
         TabIndex        =   15
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1138
      Icone           =   "TDEC101.frx":0D12
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1380
      Left            =   30
      TabIndex        =   20
      Top             =   660
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   2434
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinte 
         Height          =   315
         Left            =   1890
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   495
         Left            =   60
         TabIndex        =   1
         Top             =   810
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   873
         Caption         =   "Período(mm/aaaa)"
         Text            =   ""
         AlinhamentoRotulo=   1
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   2250
         TabIndex        =   21
         Top             =   480
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   495
         Left            =   45
         TabIndex        =   0
         Top             =   300
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   873
         Caption         =   "Insc.Municipal"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   1800
         TabIndex        =   2
         Top             =   810
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   900
         Caption         =   "Tipo Declaracão"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
   End
   Begin VTOcx.grdVISUAL grdDec 
      Height          =   1515
      Left            =   30
      TabIndex        =   22
      Top             =   2070
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   2672
      CorBorda        =   16711680
      Caption         =   "Declaracões"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
      OcultarRodape   =   -1  'True
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "mnuGeral"
      Visible         =   0   'False
      Begin VB.Menu mnuDeletar 
         Caption         =   "Deletar"
      End
   End
End
Attribute VB_Name = "TDEC101"
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
Dim ClassGrid As New grdEditavel
Dim String_Taxas As String
Dim Total_Taxas As Double

Private Sub AtualizaTextoEntradas(Indice As Integer, DeduzValoresTotais As Boolean)
    If grdEntrada.ListItems.Count = 0 Then Exit Sub
    
    txtSTInscricao = grdEntrada.ListItems(Indice).Text
    txtSTNumNota = grdEntrada.ListItems(Indice).SubItems(1)
    txtSTEmissao = grdEntrada.ListItems(Indice).SubItems(2)
    txtSTValor = grdEntrada.ListItems(Indice).SubItems(3)
    txtSTIcms = grdEntrada.ListItems(Indice).SubItems(4)
    txtSTSaldo = grdEntrada.ListItems(Indice).SubItems(5)
    txtSTImpostoDevido = grdEntrada.ListItems(Indice).SubItems(6)
    txtSTImpostoRetido = grdEntrada.ListItems(Indice).SubItems(7)
    txtSTSaldoDevedor = grdEntrada.ListItems(Indice).SubItems(8)
    txtSTAliq = grdEntrada.ListItems(Indice).SubItems(10)
    txtAidf = grdEntrada.ListItems(Indice).SubItems(11)
    If DeduzValoresTotais Then
        If CDbl(Trim(txtSTImpostoRetido)) > 0 Then
            TotalBaseST = TotalBaseST - CDbl(Nvl(txtSTSaldo, 0))
        End If
        TotalImpostoST = TotalImpostoST - CDbl(Nvl(txtSTImpostoRetido, 0))
        grdEntrada.ListItems.Remove Indice
    End If
    txtItemDecl(100) = (100 * CDbl(txtItemDecl(13))) / (TotalBaseSaida + TotalBaseST)
    AtualizaApuracao
End Sub

Private Sub PreencheDeclaracao()
    On Error Resume Next
    Dim NumDec As String
    Dim i As Integer
    If UCase(Me.ActiveControl.Name) = "CMLIMPAR" Or UCase(Me.ActiveControl.Name) = "CMDSAIR" Then Exit Sub
    IniciaTotalizadores
    grdEntrada.ListItems.Clear
    
    DeduzValores = False
    If Declaracao.Buscar(txtIM, txtPeriodo, CInt(cboTipo.Coluna(1).Valor)) Then
        cboTipo.SetarLinha Declaracao.Tipo, 1
        Declaracao.PreencheCamposDeclaracao grdEntrada, Nothing, txtItemDecl(1), txtItemDecl(2), txtItemDecl(3), txtItemDecl(4)
        TotalBaseSaida = txtItemDecl(3)
        TotalICMSSujeito = txtItemDecl(4)
        
        DoEvents
        For i = 1 To grdEntrada.ListItems.Count
            AtualizaTextoEntradas i, DeduzValores
            cmdAdicionarNotaST_Click
        Next
        
        DoEvents
        AtualizaApuracao
        txtItemDecl_LostFocus 3
        TabDec.Tabs(3).Selected = True
        txtItemDecl(1).SetFocus
    Else
'        cboTipo.SetFocus
    End If
    DeduzValores = True
End Sub

Private Sub CarregaItensGrid(Lista As Object, CFOP As eTipoNotaOperacao)
    Dim i As Integer
    Dim j As Byte
    Dim Nota As NotaFiscal
    
    On Error GoTo TRATA
    
    For i = 1 To Lista.ListItems.Count
        Set Nota = New NotaFiscal
        If Trim(Lista.ListItems(i).ListSubItems(1).Text) <> "" And Trim(Lista.ListItems(i).ListSubItems(1).Text) <> "0,00" Then
            Nota.BaseCalculo = Lista.ListItems(i).ListSubItems(5).Text
            Nota.Data = Lista.ListItems(i).SubItems(2)
            Nota.ImpostoDevido = Lista.ListItems(i).SubItems(6)
            Nota.ImpostoRetido = Lista.ListItems(i).SubItems(7)
            Nota.Numero = Lista.ListItems(i).SubItems(1)
            Nota.Status = IIf(CFOP = etoSaida, Lista.ListItems(i).SubItems(9), 0)
            Nota.TipoOperacao = CFOP
            Nota.AIDF = Lista.ListItems(i).SubItems(3)
            Nota.DescricaoServico = Lista.ListItems(i).SubItems(11)
            Nota.ValorMaterialICMS = Lista.ListItems(i).SubItems(4)
            Nota.ValorTotal = Lista.ListItems(i).SubItems(3)
            Nota.Destinatario = Lista.ListItems(i)
            Nota.Aliquota = Nvl(Lista.ListItems(i).SubItems(10), 0)
        End If
        Declaracao.Notas.Adicionar Nota
        DoEvents
    Next
    Exit Sub
TRATA:
    If Err.Number = 35600 Then
        Exit Sub
    End If
End Sub

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

Private Sub AtualizaApuracao()
    On Error Resume Next
    'RESUMO NOTAS SAIDA
    txtItemDecl(3) = TotalBaseSaida
    txtItemDecl(4) = TotalICMSSujeito
    txtItemDecl(5) = TotalBaseSaida - TotalICMSSujeito
    txtItemDecl(6) = TotalImpostoDevidoSaida
    'NOTAS DE ENTRADA
    txtItemDecl(10) = TotalImpostoST
    txtItemDecl(11) = 0
    txtItemDecl(12) = TotalImpostoST
    'NOTAS EMITIDAS(SAIDAS)
    txtItemDecl(7) = TotalImpostoRetidoSaida
    txtItemDecl(8) = 0
    txtItemDecl(9) = TotalImpostoRetidoSaida
    'TOTAL RECOLHIMENTO
    txtItemDecl(13) = TotalImpostoDevidoSaida + TotalImpostoST - TotalImpostoRetidoSaida
    txtItemDecl(100) = (100 * CDbl(txtItemDecl(13))) / (TotalBaseSaida + TotalBaseST)
End Sub

Private Sub IniciaTotalizadores()
    TotalImpostoST = 0
    TotalBaseSaida = 0
    TotalBaseST = 0
    TotalImpostoDevidoSaida = 0
    TotalImpostoRetidoSaida = 0
    TotalICMSSujeito = 0
End Sub

Private Sub cboTipo_Click()
    Declaracao.CarregaGrid grdDec, txtIM, txtPeriodo, CInt(cboTipo.Coluna(1).Valor), , Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNSUBST))
End Sub

Private Sub cboTipo_LostFocus()
  If cboTipo.Coluna(1).Valor = 2 Then
        Avisa "Utilize o formulário de ENTREGA DE DECLARACÃO NEGATIVA."
        cboTipo.ListIndex = -1
        cboTipo.SetFocus
    Else
        PreencheDeclaracao
    End If
End Sub


Private Sub cmdAdicionarNotaST_Click()
    Dim Linha As Object
    On Error Resume Next
    If Trim$(txtSTInscricao) <> "" And Trim$(txtSTNumNota) <> "" Then
        If DeduzValores Then
            Set Linha = grdEntrada.ListItems.Add(, , txtSTInscricao)
            Linha.SubItems(1) = txtSTNumNota
            Linha.SubItems(2) = txtSTEmissao
            Linha.SubItems(3) = txtSTValor
            Linha.SubItems(4) = txtSTIcms
            Linha.SubItems(5) = txtSTSaldo
            Linha.SubItems(6) = txtSTImpostoDevido
            Linha.SubItems(7) = txtSTImpostoRetido
            Linha.SubItems(8) = txtSTSaldoDevedor
            Linha.SubItems(10) = txtSTAliq
            Linha.SubItems(9) = 0
            Linha.SubItems(11) = txtAidf
        End If
        If CDbl(Trim(txtSTImpostoRetido)) > 0 Then
            TotalBaseST = TotalBaseST + CDbl(Nvl(txtSTSaldo, 0))
        End If
        TotalImpostoST = TotalImpostoST + CDbl(Nvl(txtSTImpostoRetido, 0))
        txtItemDecl(100) = (100 * CDbl(txtItemDecl(13))) / (TotalBaseSaida + TotalBaseST)
        AtualizaApuracao
        txtSTInscricao = ""
        txtSTNumNota = ""
        txtSTEmissao = ""
        txtSTValor = ""
        txtSTIcms = ""
        txtSTSaldo = ""
        txtSTImpostoDevido = ""
        txtSTImpostoRetido = ""
        txtSTInscricao.SetFocus
        txtAidf = ""
'        grdEntrada.Mensagem = "Valor a recolher: " & Format(TotalImpostoST, Const_Monetario)
    Else
        Avisa "Informe os dados corretamente."
    End If
End Sub

Private Sub CmdConsultaContribuinte_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIM
    'blnConsultaIM = False
End Sub

Private Sub cmdFinaliza_Click()
    Dim NumDec As String
    If Confirma("Ao finalizar a declaracão, ela só poderá ser modificada através de uma retificadora. Deseja prosseguir?") Then
        If cboTipo.Coluna(1).ListIndex < 0 Then
            Avisa "Informe tipo da declaracao."
            cboTipo.SetFocus
            Exit Sub
        End If
        If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
            If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
                TabDec.Tabs(3).Selected = True
                txtItemDecl(3).SetFocus
                Exit Sub
            End If
        End If
        
        
        If Not Edita.CriticaCampos(Me) Then Exit Sub
        
        If txtIM = "" Then Exit Sub
    
        
        If cboTipo.Coluna(1).ListIndex < 0 Then
            Avisa "Informe tipo da declaracao."
            cboTipo.SetFocus
            Exit Sub
        End If
        If txtPeriodo = "" Then
            Util.Avisa "Informe o período."
            txtPeriodo.SetFocus
            Exit Sub
        End If
        If CDbl(Nvl(Trim(txtItemDecl(13)), 0)) = 0 Then
            If Not Confirma("Valor Devido igual a zero. Prosseguir?") Then
                TabDec.Tabs(3).Selected = True
                txtItemDecl(3).SetFocus
                Exit Sub
            End If
        End If
    
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNSUBST))
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.Tipo = cboTipo.Coluna(1).Valor
    Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
    Declaracao.Status = decFinalizada
    CarregarItens
    If Declaracao.Gravar() Then
        Avisa "Declaração gravada com sucesso."
        Declaracao.Salvar_Sem_Finalizar , , , , CDbl(Nvl(txtItemDecl(13), 0))
        cmLimpar_Click
        txtIM.SetFocus
    End If
    End If
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtIM = "" Then Exit Sub
    
    If cboTipo.Coluna(1).ListIndex < 0 Then
        Avisa "Informe tipo da declaracao."
        cboTipo.SetFocus
        Exit Sub
    End If
    If txtPeriodo = "" Then
        Util.Avisa "Informe o período."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.Data = Format(Date, "dd/mm/yyyy")
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Declaracao.Tipo = cboTipo.Coluna(1).Valor
    Declaracao.Status = decAberta
    Declaracao.CodTributo = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNSUBST))
    Declaracao.BaseGeral = TotalBaseSaida + TotalBaseST
    CarregarItens
    If Declaracao.Gravar() Then
        Avisa "Declaração gravada com sucesso."
        'Call Pega_taxas
        Declaracao.Salvar_Sem_Finalizar True, , , , CDbl(Nvl(txtItemDecl(13), 0))
        txtIM.SetFocus
    End If
End Sub

Private Sub cmLimpar_Click()
    Edita.LimpaCampos Me
    grdEntrada.ListItems.Clear
'    grdDec.ListItems.Clear
    TabDec.Tabs(1).Selected = True
    IniciaTotalizadores
    txtIM.SetFocus
End Sub

Private Sub AtualizaDEC(CodModalidade As Integer)
    Dim Sql As String
    
    Sql = "SELECT TCD_TMD_COD_DECLARACAO,TCD_COD_CAMPO as COD,TCD_CAMPO as ITEM,TCD_VALOR_CAMPO AS VALOR FROM TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & CodModalidade
End Sub
Private Sub Form_Load()
    Dim Sql As String
    IniciaTotalizadores
    
    cabVisual1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    AtualizaDEC 0
    Set Imposto = New VsTFuncoes.VSImposto
    DeduzValores = True
    PrepararGrid grdEntrada, 20
    Set Declaracao = New cDeclaracao
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
    
    Sql = "SELECT TCD_COD_CAMPO as Item ,TCD_CAMPO as Descricao, ' ' as Valor FROM " & _
        "TAB_CONTEUDO_DECLARACAO WHERE TCD_TMD_COD_DECLARACAO = " & 1
'    ClassGrid.CarregaGrid Grid, Sql
    Sql = "Select tip_cod_imposto,tip_nome_imposto from tab_imposto where tip_sigla_imposto like 'ISS%'"
    Grdtaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
    Set Atividade = Nothing
    IniciaTotalizadores
End Sub

Private Sub grdDec_DblClick()
    On Error Resume Next
    If grdDec.ListItems.Count >= 1 Then
        cmLimpar_Click
        txtIM = grdDec.SelectedItem
        txtPeriodo = Right(grdDec.SelectedItem.SubItems(1), 2) & "/" & Left(grdDec.SelectedItem.SubItems(1), 4)
        cboTipo.SetarLinha grdDec.SelectedItem.SubItems(2), 1
        cboTipo_LostFocus
    End If
End Sub

Private Sub grdEntrada_DblClick()
    On Error Resume Next
    If grdEntrada.ListItems.Count = 0 Then Exit Sub
    AtualizaTextoEntradas grdEntrada.SelectedItem.Index, True
    AtualizaApuracao
End Sub

Public Sub CarregarItens()
    Dim Controle As Object
    Dim Item As cItemDeclaracao
    
    Declaracao.Itens.Limpar
    Declaracao.Notas.Limpar
    'NOTAS DE ENTRADA(SUBSTITUICAO TRIBUTARIA)
    CarregaItensGrid grdEntrada, etoEntrada
    'NOTAS DE SAIDA(SERVICOS PRESTADOS)
    'APURACAO DO IMPOSTO
    For Each Controle In txtItemDecl
        If Trim(Controle.Text) <> "" And Trim(Controle.Text) <> "0,00" Then
            Set Item = New cItemDeclaracao
            Item.Numero = Controle.Index
            Item.Valor = Nvl(Controle.Text, 0)
            Declaracao.Itens.Adicionar Item
        End If
    Next
End Sub

Private Sub Grid_DblClick()
'ClassGrid.EditaCelula Grid, txtGrig
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then ClassGrid.TeclaPressionada Grid, txtGrig, KeyCode
End Sub

Private Sub txtGrig_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then ClassGrid.TextoKeyDown KeyCode, Grid, txtGrig
End Sub

Private Sub txtGrig_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtGrig_LostFocus()
'    txtGrig = Format(txtGrig, Const_Monetario)
End Sub

Private Sub txtIM_LostFocus()
'    Dim eContribuinte As New eContribuinte
    
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        End If
    End If
    
'    If Trim$(txtIM) <> "" Then
'        With eContribuinte
'            If .Buscar(txtIM, , False) Then
'                txtRazao = .Nome
'                txtIM.Enabled = False
'            Else
'                Avisa "Contribuinte não cadastrado."
'                Exit Sub
'            End If
'        End With
'    End If
End Sub

Private Sub txtItemDecl_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 12
            txtItemDecl(1) = txtItemDecl(12)
            If TotalBaseST > 0 Then
                txtItemDecl(0) = txtItemDecl(12) * 100 \ TotalBaseST
            Else
                txtItemDecl(0) = 0
            End If
    End Select
End Sub

Private Sub txtItemDecl_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4
            txtItemDecl(5) = CDbl(Nvl(txtItemDecl(3), 0)) - CDbl(Nvl(txtItemDecl(4), 0))
            CalcularImposto txtItemDecl(3), txtItemDecl(4), txtItemDecl(5), txtItemDecl(6), txtItemDecl(100)
        Case 7, 8
            txtItemDecl(8) = Nvl(txtItemDecl(8), 0)
            If CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0)) >= 0 Then
                txtItemDecl(7) = CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0))
            Else
                Avisa "Dados inválidos. Valor negativo encontrado." '& Nvl(txtItemDecl(9), 0) & " - " & Nvl(txtItemDecl(8), 0) & " = " & CDbl(CDbl(Nvl(txtItemDecl(9), 0)) - CDbl(Nvl(txtItemDecl(8), 0)))
                txtItemDecl(8).SetFocus
             End If
        Case 10, 11
        txtItemDecl(11) = Nvl(txtItemDecl(11), 0)
'        If CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0)) >= 0 Then
'            txtItemDecl(10) = CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0))
'        Else
'            Avisa "Dados inválidos. Valor negativo encontrado." '& Nvl(txtItemDecl(12), 0) & " - " & Nvl(txtItemDecl(11), 0) & " = " & CDbl(CDbl(Nvl(txtItemDecl(12), 0)) - CDbl(Nvl(txtItemDecl(11), 0)))
'            txtItemDecl(11).SetFocus
'        End If
    End Select
End Sub

Private Sub CalcularImposto(ByRef Total As Object, ByRef ICMS As Object, ByRef Tributavel As Object, ByRef Imposto As Object, ByRef Aliquota As Object)
    Total = Nvl(Trim$(Total.Text), 0)
    ICMS = Nvl(Trim$(ICMS.Text), 0)
   'Aliquota = AliqISSQN * 100
   'Tributavel = Total - ICMS
   'If AliqISSQN > 0 Then
   '    Imposto = Tributavel * AliqISSQN
   'Else
   '    Imposto = ISSQNFixo
   'End If
'    Aliquota = (CDbl(txtTMAliq) * (CDbl(txtItemDecl(5)) / 100))
    Tributavel = Total - ICMS
    If AliqISSQN >= 0 Then
        Imposto = Aliquota
    Else
        Imposto = ISSQNFixo
    End If
End Sub

Private Sub txtPeriodo_Change()
'    Dim Ativ As New Atividade
    
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
'    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
'    If Trim(txtIM) <> "" And cboTipo.ListIndex <> -1 Then PreencheDeclaracao
'    AliqISSQN = Ativ.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
    If CInt(Left(Trim(txtPeriodo), 2)) > 12 Or CInt(Left(Trim(txtPeriodo), 2)) < 1 Then
        Avisa "Periodo inválido."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    cboTipo_LostFocus
    If Trim$(txtSTEmissao) = "" Then Exit Sub
    If Right(Trim(txtSTEmissao), 7) <> Trim(txtPeriodo) Then
        Avisa "Periodo da NF incompativel com periodo da declaracão."
        txtPeriodo.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(Trim(txtPeriodo)) <> 6 Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
End Sub

Private Sub txtSTIcms_LostFocus()
    txtSTIcms = Nvl(Trim$(txtSTIcms), 0)
    txtSTSaldo = CDbl(Nvl(txtSTValor, 0)) - CDbl(txtSTIcms)
    If CDbl(Nvl(txtSTAliq, 0)) > 0 Then
        txtSTImpostoDevido = txtSTSaldo * CDbl(Nvl(txtSTAliq, 0)) / 100
    Else
        txtSTImpostoDevido = ISSSTFixo
    End If
End Sub

Private Function BuscarContribuinteProprio(ByRef Inscricao As Object, Optional ByRef Nome As Object, Optional ByRef Endereco As Object, _
                    Optional ByRef Bairro As Object, Optional ByRef Cep As Object, Optional ByRef Cidade As Object, Optional ByRef Uf As Object) As Boolean
    Dim Im As Boolean
    Im = False
    If Trim(Inscricao) = "" Then Exit Function
    Inscricao.Text = Edita.TiraTudo(Inscricao.Text)
    If Len(Inscricao.Text) = 10 Then Im = True
    FormataRegistro Inscricao
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, rs As VSRecordset
    Sql = "SELECT tci_im, TCI_CGC_CPF,tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            ",TAE_NOME FROM TAB_CONTRIBUINTE left join TAB_ATIVIDADE_ECONOMICA on TCI_TAE_CAE = TAE_CAE"
    If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        Sql = Sql & " where TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " where TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    
    If Bdados.AbreTabela(Sql, rs) Then
        If Im Or Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            Inscricao = "" & rs!tci_im
        Else
            Inscricao = "" & rs!TCI_CGC_CPF
        End If
        If Not Nome Is Nothing Then Nome = "" & rs!tci_nome
        If Not Endereco Is Nothing Then Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
        If Not Bairro Is Nothing Then Bairro = "" & rs!tci_bairro
        If Not Cep Is Nothing Then Cep = "" & rs!tci_cep
        If Not Cidade Is Nothing Then Cidade = "" & rs!tci_cidade
        If Not Uf Is Nothing Then Uf = "" & rs!tci_UF
        With Declaracao
            .tciNome = "" & rs!tci_nome
            .tciEndereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_numero & " " & rs!tci_complemento
            .tciBairro = "" & rs!tci_bairro
            .tciCEP = "" & rs!tci_cep
            .tciCidade = "" & rs!tci_cidade
            .tciUF = "" & rs!tci_UF
            .tciEndereco = .tciEndereco & " " & .tciBairro & " " & .tciCidade & "-" & rs!tci_UF
            .tciAtividade = "" & rs!TAE_NOME
        End With
        
        BuscarContribuinteProprio = True
    End If
    Bdados.FechaTabela rs
End Function

Private Sub PrepararGrid(Nome_Grid As Object, IndiceInicial As Byte)
    Nome_Grid.ColumnHeaders.Clear
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Inscricao": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Nota": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Emissão": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Valor da Nota": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Sujeito ICMS": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Tributavel": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Imposto Devido": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Imposto Retido": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Saldo Devedor": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Cancelada": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "Aliquota": IndiceInicial = IndiceInicial + 1
    Nome_Grid.ColumnHeaders.Add , "Item:" & IndiceInicial + 1, "AIDF": IndiceInicial = IndiceInicial + 1
    
End Sub


Private Sub txtSTImpostoRetido_LostFocus()
    txtSTImpostoRetido = Nvl(Trim$(txtSTImpostoRetido), 0)
    txtSTSaldoDevedor = CDbl(Nvl(txtSTImpostoRetido, 0))
End Sub

Private Sub txtSTInscricao_LostFocus()
    BuscarContribuinte txtSTInscricao, txtRazao
End Sub

Private Sub txtSTValor_LostFocus()
    txtSTSaldo = CDbl(Nvl(txtSTValor, 0)) - CDbl(Nvl(txtSTIcms, 0))
    If CDbl(Nvl(txtSTAliq, 0)) > 0 Then
        txtSTImpostoDevido = txtSTSaldo * CDbl(Nvl(txtSTAliq, 0)) / 100
    Else
        txtSTImpostoDevido = ISSSTFixo
    End If
    
End Sub

Private Sub txtTMAliq_Change()
 Calc_ISSQN
End Sub

Private Sub Calc_ISSQN()
'    AliqISSQN = (txtItemDecl(3) - ((CCur(txtTMAliq) * txtItemDecl(3)) / 100))
 '   txtItemDecl_LostFocus 3
End Sub
Private Sub Pega_taxas()
    Dim i As Integer
    Dim pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To Grdtaxas.ListItems.Count
        If Grdtaxas.ListItems(i).Checked Then
            pos = InStr(Grdtaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(Grdtaxas.ListItems(i).SubItems(1), pos) & " ]" & " - " & Format(Grdtaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(Grdtaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub


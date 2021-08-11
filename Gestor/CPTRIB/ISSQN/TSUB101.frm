VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TSUB101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabVisual 
      Height          =   5850
      Left            =   45
      TabIndex        =   29
      Top             =   690
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   10319
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      BeginProperty FontHotTracking {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TSUB101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5460
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   9631
         _Version        =   131082
         TabGuid         =   "TSUB101.frx":00DB
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   960
            Left            =   45
            TabIndex        =   34
            Top             =   1515
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   1693
            Altura          =   1905
            Caption         =   " Somente para créditos já recolhidos"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483626
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboAgente 
               Height          =   510
               Left            =   120
               TabIndex        =   4
               Top             =   330
               Width           =   10470
               _ExtentX        =   18468
               _ExtentY        =   900
               Caption         =   "Agente Arrecadador do(s) Tributo(s)"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   1425
            Left            =   45
            TabIndex        =   35
            Top             =   45
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   2514
            Altura          =   1905
            Caption         =   " Substituto Tributário"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483626
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdPesq 
               Height          =   360
               Index           =   0
               Left            =   10155
               TabIndex        =   2
               Top             =   465
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   635
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtEnderecoToma 
               Height          =   480
               Left            =   120
               TabIndex        =   3
               Top             =   840
               Width           =   10455
               _ExtentX        =   18441
               _ExtentY        =   847
               Caption         =   "Endereço Tomador"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   80
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtRazaoToma 
               Height          =   480
               Left            =   2265
               TabIndex        =   1
               Tag             =   "Nome"
               Top             =   330
               Width           =   7830
               _ExtentX        =   13811
               _ExtentY        =   847
               Caption         =   "Razão Social"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   80
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtImToma 
               Height          =   480
               Left            =   135
               TabIndex        =   0
               Tag             =   "CPF/CNPJ ou IM"
               Top             =   330
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   847
               Caption         =   "CPF/CNPJ ou IM"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.grdVISUAL lstsubstituto 
            Height          =   2895
            Left            =   30
            TabIndex        =   5
            Top             =   2520
            Width           =   10680
            _ExtentX        =   18838
            _ExtentY        =   5106
            CorBorda        =   32768
            Caption         =   "Lista de Pesquisa"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5460
         Left            =   30
         TabIndex        =   31
         Top             =   30
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   9631
         _Version        =   131082
         TabGuid         =   "TSUB101.frx":0103
         Begin VTOcx.fraVISUAL fraDeducao 
            Height          =   1920
            Left            =   45
            TabIndex        =   36
            Top             =   3480
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   3387
            Altura          =   1905
            Caption         =   " Deduções"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483626
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtDtRecolhe 
               Height          =   480
               Left            =   3645
               TabIndex        =   18
               Tag             =   "Data Emissao"
               Top             =   330
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Data Recolhimento"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDtVenc 
               Height          =   480
               Left            =   7200
               TabIndex        =   43
               Tag             =   "Data Emissao"
               Top             =   330
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Vencimento"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   480
               Left            =   5415
               TabIndex        =   42
               Tag             =   "Periodo"
               Top             =   330
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Período Referência"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
               MaxLen          =   7
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtTotalNota 
               Height          =   480
               Left            =   8970
               TabIndex        =   19
               Tag             =   "Total Notas"
               Top             =   330
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Total da Nota"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDtEmissao 
               Height          =   480
               Left            =   1890
               TabIndex        =   17
               Tag             =   "Data Emissao"
               Top             =   330
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Data Emissão"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNumNota 
               Height          =   480
               Left            =   120
               TabIndex        =   16
               Tag             =   "Nota Fiscal"
               Top             =   330
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "N° Nota Fiscal"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtTotalImposto 
               Height          =   480
               Left            =   8970
               TabIndex        =   41
               Tag             =   "Nota Fiscal"
               Top             =   840
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Total Devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtObs 
               Height          =   480
               Left            =   120
               TabIndex        =   21
               Top             =   1335
               Width           =   10470
               _ExtentX        =   18468
               _ExtentY        =   847
               Caption         =   "Observação"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   80
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtMulta 
               Height          =   480
               Left            =   3645
               TabIndex        =   40
               Tag             =   "Nota Fiscal"
               Top             =   840
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Multa"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtJuros 
               Height          =   480
               Left            =   5415
               TabIndex        =   39
               Tag             =   "Nota Fiscal"
               Top             =   840
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Juros"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtISS 
               Height          =   480
               Left            =   7200
               TabIndex        =   38
               Top             =   840
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "ISS Devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   30
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtMaterial 
               Height          =   480
               Left            =   120
               TabIndex        =   20
               Top             =   840
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Vl. Material ICMS"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   30
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtBase 
               Height          =   480
               Left            =   1890
               TabIndex        =   37
               Top             =   840
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   847
               Caption         =   "Base de Cálculo"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraPrestador 
            Height          =   1935
            Left            =   45
            TabIndex        =   33
            Top             =   45
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   3413
            Altura          =   1905
            Caption         =   " Prestador de Serviço"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483626
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdPesq 
               Height          =   360
               Index           =   1
               Left            =   10155
               TabIndex        =   8
               Top             =   465
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   635
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.cboVISUAL cboAtividade 
               Height          =   510
               Left            =   4620
               TabIndex        =   14
               Top             =   1335
               Width           =   5970
               _ExtentX        =   10530
               _ExtentY        =   900
               Caption         =   "Atividade Economica"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cboVISUAL cboUF_Rem 
               Height          =   510
               Left            =   9735
               TabIndex        =   12
               Top             =   825
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   900
               Caption         =   "UF"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtBairro_Rem 
               Height          =   480
               Left            =   4620
               TabIndex        =   10
               Top             =   840
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   847
               Caption         =   "Bairro"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCep_Rem 
               Height          =   480
               Left            =   8220
               TabIndex        =   11
               Top             =   840
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   847
               Caption         =   "CEP"
               Text            =   ""
               Formato         =   4
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   10
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtMunicipio_Rem 
               Height          =   480
               Left            =   120
               TabIndex        =   13
               Top             =   1350
               Width           =   4410
               _ExtentX        =   7779
               _ExtentY        =   847
               Caption         =   "Município"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   60
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtEndereco_Rem 
               Height          =   480
               Left            =   135
               TabIndex        =   9
               Top             =   840
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   847
               Caption         =   "Endereço"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   80
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtNome_Rem 
               Height          =   480
               Left            =   2265
               TabIndex        =   7
               Tag             =   "Nome"
               Top             =   330
               Width           =   7830
               _ExtentX        =   13811
               _ExtentY        =   847
               Caption         =   "Nome Empresarial"
               Text            =   ""
               AlinhamentoRotulo=   1
               MaxLen          =   80
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtImRem 
               Height          =   480
               Left            =   135
               TabIndex        =   6
               Tag             =   "CPF/CNPJ ou IM"
               Top             =   330
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   847
               Caption         =   "CPF/CNPJ ou IM"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.grdVISUAL lstPrestador 
            Height          =   1695
            Left            =   30
            TabIndex        =   15
            Top             =   2025
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   2990
            CorBorda        =   32768
            Caption         =   "Lista de Pesquisa"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            OcultarRodape   =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5460
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   9631
         _Version        =   131082
         TabGuid         =   "TSUB101.frx":012B
         Begin VTOcx.cmdVISUAL cmdImprime 
            Height          =   375
            Left            =   75
            TabIndex        =   23
            Top             =   4995
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   661
            Caption         =   "&Imprimir Comprovantes"
            Acao            =   4
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.grdVISUAL lstNotas 
            Height          =   5130
            Left            =   45
            TabIndex        =   22
            Top             =   45
            Width           =   10680
            _ExtentX        =   18838
            _ExtentY        =   9049
            CorBorda        =   32768
            Caption         =   "Lista de Pesquisa"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1138
      Icone           =   "TSUB101.frx":0153
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   28
      Top             =   6585
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   8565
         TabIndex        =   25
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9735
         TabIndex        =   26
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7395
         TabIndex        =   24
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5610
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   44
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "TSUB101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NovoRemetente As Boolean
Dim NovoDestino As Boolean
Dim Aliquota As Double
Dim Imposto As New VSImposto
Dim CodPagamento As Double
Dim CodImposto As String
Dim NomeImposto As String
Dim NumCGC As String
Dim NumIM As String
Dim CGCToma As String
Dim AtivToma As String
Dim ImToma As String
Dim Observacao As String
Dim Substituicao As cSubstituicao
Dim Nota As cNota

'Private Function GravaDadosBaixa() As Boolean
'    With Substituicao
'        .Nota.IM_CPF = txtimRem
'        .Nota.Cod_Imposto = CodImposto
'        .Nota.Data_Venc = txtDtVenc
'        .Nota.Periodo_Ref = txtPeriodo
'        .Nota.Data_Recolhimento = txtDtRecolhe
'        .Nota.ISS_Devido = txtISS
'        .Nota.Usuario = Aplicacoes.Usuario
'        .Nota.Cod_Pagamento = CodPagamento
'        GravaDadosBaixa = .GravaDadosBaixa(txtPeriodo, Date)
'    End With
'End Function

'O GUGA

Private Sub Form_Load()
    Set Substituicao = New cSubstituicao
    Set Nota = New cNota
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rod.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    '----preencher combos---------------------------------
    Nota.PreencherCboAtividade cboAtividade
    Substituicao.PreencherCboAgente cboAgente
    cboUF_Rem.PreencherGeral Bdados, "UF"
    '------------------------------------------------------------
    MontaGrid lstNotas
    
    HabilitaRemetente True
    BuscaAliquota
End Sub

Private Sub MontaGrid(Grid As Object)
    Grid.ColumnHeaders.Add , , "Cod. Pagamento"
    Grid.ColumnHeaders.Add , , "IM"
    Grid.ColumnHeaders.Add , , "Nome Empresarial", 3500
    Grid.ColumnHeaders.Add , , "Endereço", 3500
    Grid.ColumnHeaders.Add , , "CodImposto", 0
    Grid.ColumnHeaders.Add , , "NomeImposto", 0
    Grid.ColumnHeaders.Add , , "Período"
    Grid.ColumnHeaders.Add , , "Vencimento"
    Grid.ColumnHeaders.Add , , "Base de Cáculo"
    Grid.ColumnHeaders.Add , , "ISS Devido"
    Grid.ColumnHeaders.Add , , "Multa"
    Grid.ColumnHeaders.Add , , "Juros"
    Grid.ColumnHeaders.Add , , "Atividade", 3000
    Grid.ColumnHeaders.Add , , "N° Nota Fiscal"
    Grid.ColumnHeaders.Add , , "Material", 0
    Grid.ColumnHeaders.Add , , "Data Recol."
    Grid.ColumnHeaders.Add , , "Observação", 3000
    Grid.ColumnHeaders.Add , , "AUX", 0
End Sub

Public Sub GeraDam(Aliquota As Double)
    If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAM_SUBST_Barra.rpt") Then Exit Sub
    Dim Relatorio As New cRelatorio
    Screen.MousePointer = 11
    DoEvents
    
    With Relatorio
        .Nota.Cod_Pagamento = CStr(lstNotas.SelectedItem)
        .Nota.Data_Venc = lstNotas.SelectedItem.SubItems(7)
        .Nota.Nome_Empresa = lstNotas.SelectedItem.SubItems(2)
        .Nota.Endereco.Endereco = lstNotas.SelectedItem.SubItems(3)
        .Nota.Atividade = lstNotas.SelectedItem.SubItems(12)
        .Nota.Periodo_Ref = txtPeriodo
        .Nota.ISS_Devido = lstNotas.SelectedItem.SubItems(9)
        .Nota.Multa = lstNotas.SelectedItem.SubItems(10)
        .Nota.Juros = lstNotas.SelectedItem.SubItems(11)
        .Nota.Cod_Imposto = lstNotas.SelectedItem.SubItems(4)
        .Nota.Nome_Imposto = lstNotas.SelectedItem.SubItems(5) & " - " & NomeImposto
        .Nota.Nota_fiscal = lstNotas.SelectedItem.SubItems(13)
        .Nota.Base_Calculo = lstNotas.SelectedItem.SubItems(8)
        .OBS = lstNotas.SelectedItem.SubItems(17) & ". " & lstNotas.SelectedItem.SubItems(16)
        .GeraDam ImToma, CGCToma, txtRazaoToma, txtEnderecoToma, PicBarra
    End With
    
    Set Relatorio = Nothing
    Screen.MousePointer = 0
    DoEvents
End Sub

Private Sub IncluiNotaFiscal()
    Dim Lista As Object
    On Error Resume Next
    Set Lista = lstNotas.ListItems.Add(, , CodPagamento)
    Lista.SubItems(1) = txtImRem
    Lista.SubItems(2) = txtNome_Rem
    Lista.SubItems(3) = txtEndereco_Rem & " " & txtBairro_Rem & " " & txtCep_Rem & " " & CStr(cboUF_Rem.Coluna(0).Valor) & " " & txtMunicipio_Rem
    Lista.SubItems(4) = CodImposto
    Lista.SubItems(5) = Imposto.NomeTributo(ttr_ISSQNSUBST)
    Lista.SubItems(6) = txtPeriodo
    Lista.SubItems(7) = txtDtVenc
    Lista.SubItems(8) = txtBase
    Lista.SubItems(9) = txtISS
    Lista.SubItems(10) = txtMulta
    Lista.SubItems(11) = txtJuros
    Lista.SubItems(12) = CStr(cboAtividade.Coluna(0).Valor)
    Lista.SubItems(13) = txtNumNota
    Lista.SubItems(14) = txtMaterial
    Lista.SubItems(15) = txtDtRecolhe
    Lista.SubItems(16) = txtObs
    Lista.SubItems(17) = IIf(cboAgente.ListIndex >= 0, "Valor recolhido em " & lstNotas.SelectedItem.SubItems(15) & " no(a) " & cboAgente.Coluna(0).Valor, "")
    'Lista.SubItems(18) = Observacao = IIf(cboAgente.ListIndex >= 0, "Valor recolhido em " & lstNotas.SelectedItem.SubItems(15) & " no(a) " & cboAgente.Text & " por " & txtRazaoToma, Observacao)
End Sub

Sub BuscaAliquota()
    With Substituicao
        .BuscaAliquota Date, txtPeriodo
        Aliquota = .Nota.Aliquota
        CodImposto = .Nota.Cod_Imposto
        NomeImposto = .Nota.Nome_Imposto
    End With
End Sub

Sub HabilitaRemetente(Status As Boolean)
    txtNome_Rem.Enabled = Status
    txtEndereco_Rem.Enabled = Status
    txtBairro_Rem.Enabled = Status
    txtCep_Rem.Enabled = Status
    cboUF_Rem.Enabled = Status
    txtMunicipio_Rem.Enabled = Status
    cboAtividade.Enabled = Status
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    If lstNotas.SelectedItem Is Nothing Then Util.Informa "Selecione uma Nota.": Exit Sub
    GeraDam Aliquota
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    tabVisual.Tabs(1).Selected = True
    txtImToma.SetFocus
    lstPrestador.ListItems.Clear
    lstsubstituto.ListItems.Clear
    NumIM = ""
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    Dim Razao As String
    Screen.MousePointer = 11
    Razao = UCase(Trim(IIf(Index = 0, txtRazaoToma, txtNome_Rem)))
    '-------------------------------------------------------------------------------------
    Substituicao.PreencheGrid lstsubstituto, lstPrestador, Razao, Index
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Screen.MousePointer = 11
    If Not CriticaCampos(Me) Then: Screen.MousePointer = 0: Exit Sub
    Dim CodAgente As String
    With Substituicao
        .Nota.Usuario = Aplicacoes.Usuario
        .Nota.IM_CPF = txtImRem
        .Nota.Nome_Empresa = txtNome_Rem
        .Nota.Endereco.Endereco = txtEndereco_Rem
        .Nota.Endereco.Bairro = txtBairro_Rem
        .Nota.Endereco.CEP = txtCep_Rem
        .Nota.Endereco.municipio = txtMunicipio_Rem
        .Nota.Endereco.UF = CStr(cboUF_Rem.Coluna(0).Valor)
        .Nota.Cod_Imposto = CodImposto
        .Nota.Periodo_Ref = txtPeriodo
        .Nota.Data_Venc = txtDtVenc
        .Nota.ISS_Devido = txtISS
        .Nota.Cod_Pagamento = CodPagamento
        .Nota.Nota_fiscal = txtNumNota
        .Nota.Total_Nota = IIf(txtTotalNota = "", "0", txtTotalNota)
        .Nota.Valor_Material_ICMS = IIf(txtMaterial = "", "0", txtMaterial)
        .Nota.NumIM = NumIM
        .Nota.Data_emissao = txtDtEmissao
        .Nota.Data_Recolhimento = txtDtRecolhe
        If Not .Salvar(NovoRemetente, ImToma, Date) Then
            Util.Erro "Erro de gravação!"
            Exit Sub
        End If
        CodPagamento = .Nota.Cod_Pagamento
    End With
    IncluiNotaFiscal
    CodAgente = CStr(cboAgente.Coluna(1).Valor)
    Observacao = txtObs
    Edita.LimpaCampos Me
    txtImToma = Edita.TiraPic(ImToma, "-")
    If Trim(CodAgente) <> "" Then cboAgente.SetarLinha CodAgente, 1
    txtImToma_LostFocus
    txtImRem.SetFocus
    'Cobranca.ImprimeDam Rpt, CodPagamento, txtimRem, txtNome_Rem, "", txtEndereco_Rem, "", "", _
    CodImposto, Imposto.NomeTributo(ttr_ISSQN), NomeImposto, txtPeriodo, 0, 1, txtDtVenc, txtBase, txtISS, _
    txtMulta, txtJuros, 0, 0, cstr(cboAtividade.coluna(0).valor), "", PicBarra, txtNumNota, txtNumNota, txtMaterial
    NumIM = ""
    'Edita.LimpaCampos Me
    Screen.MousePointer = 0
    Util.Informa "Nota Fiscal gravada com sucesso."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Substituicao = Nothing
    Set Nota = Nothing
End Sub

Private Sub lstPrestador_Click()
    On Error Resume Next
    If lstPrestador.SelectedItem Is Nothing Then Exit Sub
    txtImRem = lstPrestador.SelectedItem
    txtimRem_LostFocus
End Sub

Private Sub lstsubstituto_Click()
    On Error Resume Next
    If lstsubstituto.SelectedItem Is Nothing Then Exit Sub
    txtImToma = lstsubstituto.SelectedItem
    txtImToma_LostFocus
End Sub

Private Sub lstsubstituto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lstsubstituto_Click
    cboAgente.SetFocus
End Sub

Private Sub txtBairro_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBase_Change()
    On Error Resume Next
    txtISS = CStr(CDbl(txtBase) * (Aliquota))
End Sub

Private Sub txtDtEmissao_LostFocus()
    If UCase(Me.ActiveControl.Name) = "CMDSALVAR" Or UCase(Me.ActiveControl.Name) = "CMDSAIR" Or UCase(Me.ActiveControl.Name) = "CMDNOVO" Then Exit Sub
    If IsNumeric(txtDtEmissao) Then txtDtEmissao = Edita.FormataTexto(txtDtEmissao, Data)
    If IsDate(txtDtEmissao) Then
        If CDbl(Year(txtDtEmissao) & Mid(txtDtEmissao, 4, 2) & Left(txtDtEmissao, 2)) > CDbl(Right(Date, 4) & Mid(Date, 4, 2) & Left(Date, 2)) Then
            Avisa "Data de emissão da nota não pode ser superior a atual."
            txtDtEmissao.SetFocus
            Exit Sub
        End If
        txtPeriodo = Format(Month(txtDtEmissao), "00") & Year(txtDtEmissao)
        txtPeriodo_LostFocus
    End If
End Sub

Private Sub txtDtRecolhe_LostFocus()
    If UCase(Me.ActiveControl.Name) = "CMDSALVAR" Or UCase(Me.ActiveControl.Name) = "CMDSAIR" Or UCase(Me.ActiveControl.Name) = "CMDNOVO" Then Exit Sub
    If IsNumeric(txtDtRecolhe) Then txtDtRecolhe = Edita.FormataTexto(txtDtRecolhe, Data)
    If IsDate(txtDtRecolhe) And IsDate(txtDtEmissao) Then
        If CDbl(Right(txtDtRecolhe, 4) & Mid(txtDtRecolhe, 4, 2) & Left(txtDtRecolhe, 2)) < CDbl(Right(txtDtEmissao, 4) & Mid(txtDtEmissao, 4, 2) & Left(txtDtEmissao, 2)) Then
            Avisa "Data de recolhimento do imposto não pode ser inferior a data de emissão da nota."
            txtDtRecolhe.SetFocus
            Exit Sub
        End If
    End If
    txtISS_Change
End Sub

Private Sub txtEndereco_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtimRem_LostFocus()
     If Trim(txtImRem) = "" Then Exit Sub
    If Len(txtImRem) = 10 And IsNumeric(txtImRem) Then
        txtImRem = Imposto.FormataInscricao(txtImRem, InscContrib)
    ElseIf Len(txtImRem) = 11 And IsNumeric(txtImRem) Then
        txtImRem = Edita.FormataTexto(txtImRem, Cpf)
    ElseIf Len(txtImRem) = 14 Then
        txtImRem = Edita.FormataTexto(txtImRem, Cgc)
    End If
    If Trim(txtImRem) = "" Then Exit Sub
    '-------------------------------------------------------------------------------
    Screen.MousePointer = 11
    DoEvents
    With Nota
        If .Buscar(txtImRem, txtImRem) Then
            HabilitaRemetente False
            If txtImRem = Const_ImAvulso Then
                txtEndereco_Rem = ""
                txtBairro_Rem = ""
                txtCep_Rem = ""
                cboUF_Rem.ListIndex = -1
                txtMunicipio_Rem = ""
                cboAtividade.ListIndex = -1
            Else
                txtEndereco_Rem = .Endereco.Endereco
                txtBairro_Rem = .Endereco.Bairro
                txtCep_Rem = .Endereco.CEP
                cboUF_Rem.SetarLinha .Endereco.UF, 0
                txtMunicipio_Rem = .Endereco.municipio
                cboAtividade.SetarLinha .Atividade, 1
            End If
            NumCGC = .NumCGC
            NumIM = .NumIM
            txtNome_Rem = .Nome_Empresa
            NovoRemetente = False
            txtNumNota.SetFocus
        Else
            HabilitaRemetente True
            NovoRemetente = True
            '----------------------------------------------------
            txtImRem.Tag = txtImRem
            LimpaCampos Me
            txtImRem = txtImRem.Tag
            txtImRem.Tag = "CPF/CNPJ ou IM"
            '----------------------------------------------------
            txtNome_Rem.SetFocus
        End If
    End With
    Screen.MousePointer = 0
End Sub

Private Sub txtImToma_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Impost As New VSImposto
    If Trim(txtImToma) = "" Then Exit Sub
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtImToma) = 14 Then
            txtImToma = Edita.FormataTexto(txtImToma, Cgc)
        ElseIf Len(txtImToma) = 11 And IsNumeric(txtImToma) Then
            txtImToma = Edita.FormataTexto(txtImToma, Cpf)
        Else
            txtImToma = Impost.FormataInscricao(txtImToma, InscContrib)
        End If
    End If
    With Substituicao
        If Not .BuscarTomador(txtImToma) Then
            txtImToma = ""
            ImToma = ""
            CGCToma = ""
            AtivToma = ""
            txtRazaoToma = ""
            txtEnderecoToma = ""
            Util.Avisa "Contribuinte Não cadastrado como Substituto Tributário ou em situação irregular."
            Screen.MousePointer = 0
            txtImToma.SetFocus
        Else
            ImToma = .Nota.IM_CPF
            CGCToma = .Nota.NumCGC
            AtivToma = .Nota.Atividade
            txtRazaoToma = .Nota.Nome_Empresa
            txtEnderecoToma = .Nota.Endereco.Endereco
        End If
    End With
End Sub

Private Sub txtISS_Change()
    Dim Conta As New ContaCorrente
    Dim DataHoje As Date
    If Trim(txtISS) = "" Then Exit Sub
    On Error Resume Next
    DataHoje = Date
    Date = txtDtRecolhe
    txtJuros = Format(Conta.CalculaValoresJurosAvulsos(CodImposto, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2)), EtcCreditoTributario, Format(Date, "dd/mm/yyyy"), txtDtVenc, txtISS), Const_Monetario)
    txtMulta = Format(Conta.CalculaValoresMultaAvulsos(CodImposto, CLng(Right(txtPeriodo, 4) & Left(txtPeriodo, 2)), EtcCreditoTributario, Format(Date, "dd/mm/yyyy"), txtDtVenc, txtISS), Const_Monetario)
    txtTotalImposto = Format(CDbl(txtISS) + CDbl(txtJuros) + CDbl(txtMulta), Const_Monetario)
    Date = DataHoje
End Sub

Private Sub txtMaterial_Change()
    On Error Resume Next
    txtBase = CDbl(Nvl(txtTotalNota, 0)) - CDbl(Nvl(txtMaterial, 0))
End Sub

Private Sub txtMaterial_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtMaterial_LostFocus()
    If CDbl(Nvl(txtTotalNota, 0)) < CDbl(Nvl(txtMaterial, 0)) Then
        Avisa "Valor não pode ser maior que Total em notas."
        txtMaterial = "0,00"
        txtMaterial.SetFocus
    End If
End Sub

Private Sub txtMunicipio_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNome_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPeriodo_Change()
    If Trim(txtPeriodo) = "" Then Exit Sub
    txtDtVenc = Imposto.BuscaDataVencimento(CodImposto, txtPeriodo)
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtPeriodo_LostFocus()
    If Trim(txtPeriodo) = "" Then Exit Sub
    If IsNumeric(txtPeriodo) Then
        If Len(txtPeriodo) = 6 Then
            txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
            txtDtVenc = Imposto.BuscaDataVencimento(CodImposto, txtPeriodo)
            If Trim(txtTotalNota) <> "" Then
                txtISS_Change
            End If
        Else
            Avisa "Período inválido."
            txtPeriodo.SetFocus
        End If
    End If
End Sub

Private Sub txtRazaoToma_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTotalNota_Change()
    On Error Resume Next
    txtBase = CDbl(Nvl(txtTotalNota, 0)) - CDbl(Nvl(txtMaterial, 0))
End Sub


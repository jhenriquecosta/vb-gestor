VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{467EEF11-5281-4102-AFD3-AD54F754C329}#1.5#0"; "VTControles.ocx"
Object = "{741D44DD-BF8E-4BC8-85FF-338C9BF39DFB}#1.0#0"; "Cabecalho.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RREM103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RREM103"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabCabecalho 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1138
      Icone           =   "RREM103.frx":0000
      ImagemFundo     =   "RREM103.frx":031A
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   6300
      Left            =   120
      TabIndex        =   11
      Top             =   780
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   11113
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "RREM103.frx":23F4C
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5910
         Left            =   -99969
         TabIndex        =   13
         Top             =   30
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   10425
         _Version        =   131082
         TabGuid         =   "RREM103.frx":23FDC
         Begin VB.Frame Frame1 
            Caption         =   "Informe o caminho dos arquivos de remessa.ret"
            Height          =   720
            Left            =   150
            TabIndex        =   26
            Top             =   75
            Width           =   9060
            Begin VB.TextBox txtCamminhoRemessa 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   90
               TabIndex        =   6
               Top             =   240
               Width           =   6615
            End
            Begin VTOcx.cmdVISUAL cmdConsultaArquivo 
               Height          =   405
               Left            =   6780
               TabIndex        =   7
               Top             =   195
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   714
               Caption         =   "Arquivo"
               Acao            =   5
               CorBorda        =   -2147483645
               CorFrente       =   -2147483630
               CorFoco         =   -2147483628
            End
            Begin VTOcx.cmdVISUAL cmdReceber 
               Height          =   405
               Left            =   7890
               TabIndex        =   38
               Top             =   195
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   714
               Caption         =   "&Receber"
               Acao            =   3
               CorBorda        =   -2147483645
               CorFrente       =   -2147483630
               CorFoco         =   -2147483628
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Dados da Recpção"
            Height          =   1725
            Left            =   150
            TabIndex        =   14
            Top             =   780
            Width           =   9075
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   1155
               Left            =   5580
               TabIndex        =   29
               Top             =   165
               Width           =   3270
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL TITULO:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   420
                  TabIndex        =   37
                  Top             =   60
                  Width           =   1245
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL JUROS:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   465
                  TabIndex        =   36
                  Top             =   585
                  Width           =   1185
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL GERAL:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   495
                  TabIndex        =   35
                  Top             =   840
                  Width           =   1170
               End
               Begin VB.Label LblTotalTitulo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0,00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   2580
                  TabIndex        =   34
                  Top             =   75
                  Width           =   360
               End
               Begin VB.Label LblTotalJuros 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0,00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   2580
                  TabIndex        =   33
                  Top             =   600
                  Width           =   360
               End
               Begin VB.Label LblTotalGeral 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0,00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   2580
                  TabIndex        =   32
                  Top             =   840
                  Width           =   360
               End
               Begin VB.Label LblTotalDesconto 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "0,00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   2580
                  TabIndex        =   31
                  Top             =   330
                  Width           =   360
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL DESCONTO:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   165
                  TabIndex        =   30
                  Top             =   330
                  Width           =   1500
               End
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   75
               Top             =   270
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSComctlLib.ProgressBar Progresso 
               Height          =   225
               Left            =   90
               TabIndex        =   15
               Top             =   1410
               Visible         =   0   'False
               Width           =   8910
               _ExtentX        =   15716
               _ExtentY        =   397
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.Label lblLote 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2400
               TabIndex        =   25
               Top             =   195
               Width           =   60
            End
            Begin VB.Label LblTotalRegistro 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2400
               TabIndex        =   24
               Top             =   1140
               Width           =   45
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Conta:"
               Height          =   195
               Left            =   1860
               TabIndex        =   23
               Top             =   645
               Width           =   495
            End
            Begin VB.Label LblConta 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2400
               TabIndex        =   22
               Top             =   660
               Width           =   45
            End
            Begin VB.Label Agencia 
               AutoSize        =   -1  'True
               Caption         =   "Agência:"
               Height          =   195
               Left            =   1740
               TabIndex        =   21
               Top             =   420
               Width           =   630
            End
            Begin VB.Label LblAgencia 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2400
               TabIndex        =   20
               Top             =   420
               Width           =   45
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Data Recepção:"
               Height          =   195
               Index           =   0
               Left            =   1215
               TabIndex        =   19
               Top             =   900
               Width           =   1155
            End
            Begin VB.Label LblDataRecepcao 
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   2400
               TabIndex        =   18
               Top             =   900
               Width           =   45
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Total Pagamentos:"
               Height          =   195
               Left            =   1020
               TabIndex        =   17
               Top             =   1155
               Width           =   1350
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Lote:"
               Height          =   195
               Index           =   1
               Left            =   1995
               TabIndex        =   16
               Top             =   195
               Width           =   375
            End
         End
         Begin VTOcx.grdVISUAL grdDocumentos 
            Height          =   3675
            Left            =   150
            TabIndex        =   27
            Top             =   2550
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   6482
            CorBorda        =   4210752
            Caption         =   "Arquivos"
            CorCaption      =   4210752
            CorDica         =   0
            OcultarRodape   =   -1  'True
            PictureBarra    =   "RREM103.frx":24004
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5910
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   10425
         _Version        =   131082
         TabGuid         =   "RREM103.frx":24020
         Begin VTOcx.cboVISUAL CboArquivos 
            Height          =   315
            Left            =   3660
            TabIndex        =   2
            Top             =   540
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            Caption         =   "Arquivos"
            Text            =   ""
            AutoFocaliza    =   0   'False
            TipoCampo       =   ""
            PictureFundo    =   "RREM103.frx":24048
         End
         Begin VTOcx.txtVISUAL txtDataInicial 
            Height          =   315
            Left            =   255
            TabIndex        =   0
            Top             =   180
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            Caption         =   "Data Recepção Inicial"
            Text            =   ""
            Formato         =   0
            PictureFundo    =   "RREM103.frx":24064
         End
         Begin VTOcx.grdVISUAL GrdItemPagamentos 
            Height          =   3015
            Left            =   90
            TabIndex        =   5
            Top             =   3210
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   5318
            CorBorda        =   4210752
            Caption         =   "Detalhes"
            CorCaption      =   4210752
            CorDica         =   0
            OcultarRodape   =   -1  'True
            PictureBarra    =   "RREM103.frx":24080
         End
         Begin VTOcx.grdVISUAL grdArquivos 
            Height          =   2580
            Left            =   90
            TabIndex        =   4
            Top             =   960
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   4551
            CorBorda        =   4210752
            Caption         =   "Arquivos"
            CorCaption      =   4210752
            CorDica         =   0
            OcultarRodape   =   -1  'True
            PictureBarra    =   "RREM103.frx":2409C
         End
         Begin VTOcx.txtVISUAL txtDataFinal 
            Height          =   315
            Left            =   375
            TabIndex        =   1
            Top             =   540
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            Caption         =   "Data Recepção Final"
            Text            =   ""
            Formato         =   0
            PictureFundo    =   "RREM103.frx":240B8
         End
         Begin VTOcx.cmdVISUAL cmdConsultar 
            Height          =   405
            Left            =   7935
            TabIndex        =   3
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
            Caption         =   "&Consultar"
            Acao            =   5
            CorBorda        =   -2147483645
            CorFrente       =   -2147483630
            CorFoco         =   -2147483628
         End
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   330
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   10
      Top             =   7140
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   953
      CorFundo        =   -2147483632
      CorFrente       =   -2147483633
      ImagemFundo     =   "RREM103.frx":240D4
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   405
         Left            =   6645
         TabIndex        =   39
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   7815
         TabIndex        =   8
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   8835
         TabIndex        =   9
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
End
Attribute VB_Name = "RREM103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Arquivo   As Arquivo
Public Cabecalho As Header
Dim Path         As String
Private Sub cmdConsultaArquivo_Click()
      With CommonDialog1
        .DialogTitle = "Selecione o arquivo retorno do bradesco"
        .Filter = "Arquivos do tipo retorno | *.ret"
        .ShowOpen
        If .FileName <> "" Then
            txtCamminhoRemessa = .FileName
        End If
    End With
End Sub

Private Sub cmdConsultar_Click()
    Dim sql As String
    
    sql = "SELECT trp_arquivo as Arquivo,"
    sql = sql & " tpr_numero_lote as Lote,"
    sql = sql & " tpr_data_recepcao as Recepção,"
    sql = sql & " tpr_usuario as Usuário,"
    sql = sql & " tpr_total_pagamentos as [Qtd Registros] "
    sql = sql & " FROM TAB_PAGAMENTO_RECEBIDO where 1 = 1"
    
    If CboArquivos <> "" Then
        sql = sql & " and trp_arquivo = '" & CboArquivos.Text & "'"
    End If
    
    If txtDataInicial <> "" And txtDataFinal <> "" Then
        sql = sql & " and tpr_data_recepcao >=" & Bdados.Converte(txtDataInicial, TCDataHora)
        sql = sql & " and tpr_data_recepcao <=" & Bdados.Converte(txtDataFinal, TCDataHora)
    ElseIf txtDataInicial <> "" And txtDataFinal = "" Then
        sql = sql & " and tpr_data_recepcao >=" & Bdados.Converte(txtDataInicial, TCDataHora)
        sql = sql & " and tpr_data_recepcao <=" & Bdados.Converte(txtDataInicial, TCDataHora)
    End If
    GrdItemPagamentos.ListItems.Clear
    If grdArquivos.Preencher(Bdados, sql) = False Then
        Avisa "Consulta sem resultados"
    End If
End Sub

Private Sub CmdImprimir_Click()
    Dim Condicao As String
    
    If TabDados.Tabs(1).Selected Then
        If grdArquivos.ListItems.Count >= 1 Then
            Condicao = "{TAB_PAGAMENTO_RECEBIDO.TRP_ARQUIVO} = '" & grdArquivos.SelectedItem & "'"
        Else
            Avisa "Selecione um arquivos"
        End If
    Else
        If grdDocumentos.ListItems.Count >= 1 Then
            Condicao = "{TAB_PAGAMENTO_RECEBIDO.TRP_ARQUIVO} = '" & Left(Right(txtCamminhoRemessa, 12), Len(Right(txtCamminhoRemessa, 12)) - 4) & "'"
        Else
            Avisa "Selecione um arquivos"
        End If
    End If
    
    With Rela
        If .DefinirArquivo(Bdados, App.Path & "\" & Me.Name & ".rpt") = False Then Exit Sub
        .Selecao = Condicao
        .Visualizar
    End With
End Sub

Private Sub cmdLimpar_Click()
    LblBanco = "": LblAgencia = "": LblDtGeracaoArquivo = "": LblTotalArrecardado = ""
    LblDataPagamentoDOC = "": LblTotalRegistroRemessa = "": LblPercento = ""
    lblAceitos = "": lblRejeitado = "": lblLote = "": lblTotalBaixado = "": lblTotalAberto = ""
    txtCamminhoRemessa = ""
    grdDocumentos.ListItems.Clear
    If TabDados.Tabs(1).Selected Then
        txtDataInicial.SetFocus
    Else
        txtCamminhoRemessa.SetFocus
    End If
End Sub

Private Sub cmdReceber_Click()
    Dim Campos As String
    Dim Valores As String
    Dim Condicao As String
    Dim Dados_Geral As String
    Dim TemTransa As Boolean
    Dim sql As String
    Dim Iregistro As Integer
    Dim Arquivo As Integer
    Dim Dados As String
    Dim NUMERO_LOTE As String
    Dim DATA_RECEPCAO As String
    Dim Usuario As String
    Dim TOTAL_PAGAMENTOS As String
    Dim TRP_ARQUIVO As String
    
    Dim TPI_TPR_NUMERO_LOTE As String
    Dim TPI_COD_PAGAMENTO As String
    Dim TPI_DATA_PAGAMENTO As String
    Dim TPI_VALOR_TITULO As String
    Dim TPI_DESCONTO As String
    Dim TPI_VALOR_PAGO As String
    Dim TPI_VENCIMENTO  As String
    Dim TPI_JUROS  As String
    Dim TPI_OCORRENCIA As String
    Dim TPI_MOTIVO_01 As String
    Dim TPI_MOTIVO_02 As String
    Dim TPI_MOTIVO_03 As String
    Dim TPI_MOTIVO_04 As String
    Dim TPI_MOTIVO_05 As String
    Dim TPI_DESPESA_COBRANCA As Currency
    Dim Documento_Pago As Boolean
    Dim Ocorrencia As String
    
    
    'Corrigo o arquivo, só para garantir que a leitura será um sucesso.
    'Arquivo = FreeFile
    'Open txtCamminhoRemessa For Input As #Arquivo
    'Do Until EOF(Arquivo)
    '    Line Input #Arquivo, Dados
    '    Dados_Geral = Dados_Geral & Dados
    'Loop
    'Close Arquivo
    
    
    LblTotalTitulo = "0,00"
    LblTotalDesconto = "0,00"
    LblTotalJuros = "0,00"
    LblTotalGeral = "0,00"
        
    If txtCamminhoRemessa = "" Then
        Avisa "Selecione o arquivo"
        cmdConsultaArquivo_Click
        Exit Sub
    End If
    
    If Pega_Dados_Lote(PegaConfiguracaoEscola(TEC_CONTA_CORRENTE_BRADESCO), , edlp_TLP_SITUACAO_LOTE) = sL_Lote_Aberto Then
        Avisa "Operação cancelada: O Lote " & Pega_Dados_Lote(PegaConfiguracaoEscola(TEC_CONTA_CORRENTE_BRADESCO)) & " deve ser fechado"
        Exit Sub
    End If
    
    
    DATA_RECEPCAO = Date
    Usuario = Aplica.Usuario
    TOTAL_PAGAMENTOS = 0
    lblLote = NUMERO_LOTE
    
    If Bdados.AbreTabela("SELECT TCB_AGENCIA,TCB_NUMERO FROM TAB_CONTA_BANCARIA WHERE TCB_COD_CONTA = " & PegaConfiguracaoEscola(TEC_CONTA_CORRENTE_BRADESCO)) Then
        LblAgencia = Bdados.Tabela("TCB_AGENCIA")
        LblConta = Bdados.Tabela("TCB_NUMERO")
    Else
        LblAgencia = ""
        LblConta = ""
    End If
    LblDataRecepcao = Date
    TRP_ARQUIVO = Left(Right(txtCamminhoRemessa, 12), Len(Right(txtCamminhoRemessa, 12)) - 4)
    'If Bdados.AbreTabela("Select * from tab_pagamento_recebido where TRP_ARQUIVO  = '" & TRP_ARQUIVO & "'") Then
    '    Avisa "Retorno já recebido"
    '    txtCamminhoRemessa.SetFocus
    '    Exit Sub
    'End If
    Bdados.AbreTrans
    TemTransa = True
    Arquivo = FreeFile
    Open txtCamminhoRemessa For Input As #Arquivo
    Do Until EOF(Arquivo)
        Line Input #Arquivo, Dados
        If Left(Dados, 1) = 0 Then 'heder
            Campos = "TRP_ARQUIVO,"
            Campos = Campos & " TPR_DATA_RECEPCAO,"
            Campos = Campos & " TPR_USUARIO,"
            Campos = Campos & " TPR_TOTAL_PAGAMENTOS"
            Valores = Bdados.PreparaValor(TRP_ARQUIVO, DATA_RECEPCAO, Usuario, TOTAL_PAGAMENTOS)
            Call Bdados.GravaDados("TAB_PAGAMENTO_RECEBIDO", Valores, Campos, "TRP_ARQUIVO = '" & TRP_ARQUIVO & "'")
        ElseIf Left(Dados, 1) = 1 Then
            TOTAL_PAGAMENTOS = TOTAL_PAGAMENTOS + 1
            Campos = "TPI_TPR_ARQUIVO,"
            Campos = Campos & " TPI_COD_PAGAMENTO,"
            Campos = Campos & " TPI_DATA_PAGAMENTO,"
            Campos = Campos & " TPI_VENCIMENTO,"
            Campos = Campos & " TPI_VALOR_TITULO,"
            Campos = Campos & " TPI_DESCONTO,"
            Campos = Campos & " TPI_JUROS,"
            Campos = Campos & " TPI_VALOR_PAGO,"
            Campos = Campos & " TPI_OCORRENCIA,"
            Campos = Campos & " TPI_MOTIVO_01,"
            Campos = Campos & " TPI_MOTIVO_02,"
            Campos = Campos & " TPI_MOTIVO_03,"
            Campos = Campos & " TPI_MOTIVO_04,"
            Campos = Campos & " TPI_MOTIVO_05,"
            Campos = Campos & " TPI_PAGAMENTO_EXISTE,TPI_DESPESA_COBRANCA"
                
            TPI_TPR_ARQUIVO = TRP_ARQUIVO
            TPI_TPR_NUMERO_LOTE = NUMERO_LOTE
            TPI_COD_PAGAMENTO = Left(Mid(Dados, 71, 12), Len(Mid(Dados, 71, 12)) - 1)
            'TPI_DATA_PAGAMENTO = IIf(InStr(Mid(Dados, 296, 6), "7") > 0, Left(Mid(Dados, 296, 6), 2) & "/" & Mid(Mid(Dados, 296, 6), 3, 2) & "/" & "20" & Right(Mid(Dados, 296, 6), 2), Date)
            TPI_DATA_PAGAMENTO = Left(Mid(Dados, 296, 6), 2) & "/" & Mid(Mid(Dados, 296, 6), 3, 2) & "/" & "20" & Right(Mid(Dados, 296, 6), 2)
            TPI_VENCIMENTO = IIf(InStr(Mid(Dados, 147, 6), "7") > 0, Left(Mid(Dados, 147, 6), 2) & "/" & Mid(Mid(Dados, 147, 6), 3, 2) & "/" & "20" & Right(Mid(Dados, 147, 6), 2), Date)
           
           
           TPI_VALOR_TITULO = FormatNumber(Left(Mid(Dados, 153, 13), Len(Mid(Dados, 153, 13)) - 2) & "," & Right(Mid(Dados, 153, 13), 2), 2)
            TPI_DESCONTO = FormatNumber(Left(Mid(Dados, 241, 13), Len(Mid(Dados, 241, 13)) - 2) & "," & Right(Mid(Dados, 241, 13), 2), 2)

            TPI_JUROS = FormatNumber(Left(Mid(Dados, 267, 13), Len(Mid(Dados, 267, 13)) - 2) & "," & Right(Mid(Dados, 267, 13), 2), 2)
            TPI_VALOR_PAGO = FormatNumber(Left(Mid(Dados, 254, 13), Len(Mid(Dados, 254, 13)) - 2) & "," & Right(Mid(Dados, 254, 13), 2), 2)
            
            TPI_OCORRENCIA = Mid(Dados, 109, 2)
            If TPI_OCORRENCIA = 19 Then
                TPI_MOTIVO_01 = Mid(Dados, 295, 1)
                If TPI_MOTIVO_01 = "A" Then 'A - Aceito
                    TPI_MOTIVO_01 = 1
                Else
                    TPI_MOTIVO_01 = 2 'D - Desprezado
                End If
                'LImpa os demais motivos
                TPI_MOTIVO_02 = ""
                TPI_MOTIVO_03 = ""
                TPI_MOTIVO_04 = ""
                TPI_MOTIVO_05 = ""
            Else
                TPI_MOTIVO_01 = Mid(Dados, 319, 2)
                TPI_MOTIVO_02 = Mid(Dados, 321, 2)
                TPI_MOTIVO_03 = Mid(Dados, 323, 2)
                TPI_MOTIVO_04 = Mid(Dados, 325, 2)
                TPI_MOTIVO_05 = Mid(Dados, 327, 2)
            End If
            TPI_DESPESA_COBRANCA = FormatNumber(Left(Mid(Dados, 176, 13), Len(Mid(Dados, 176, 13)) - 2) & "," & Right(Mid(Dados, 176, 13), 2), 2)

            Valores = Bdados.PreparaValor(TPI_TPR_ARQUIVO, _
            TPI_COD_PAGAMENTO, _
            TPI_DATA_PAGAMENTO, _
            TPI_VENCIMENTO, _
            Bdados.Converte(TPI_VALOR_TITULO, TCMonetario), _
            Bdados.Converte(TPI_DESCONTO, TCMonetario), _
            Bdados.Converte(TPI_JUROS, TCMonetario), _
            Bdados.Converte(TPI_VALOR_PAGO, TCMonetario), _
            TPI_OCORRENCIA, _
            TPI_MOTIVO_01, _
            TPI_MOTIVO_02, _
            TPI_MOTIVO_03, _
            TPI_MOTIVO_04, _
            TPI_MOTIVO_05, _
            Pagamento_Existe(TPI_COD_PAGAMENTO), TPI_DESPESA_COBRANCA)
            
            If Bdados.GravaDados("TAB_ITEM_PAGAMENTO", Valores, Campos, "TPI_TPR_ARQUIVO = '" & TPI_TPR_ARQUIVO & "' AND TPI_COD_PAGAMENTO  = " & TPI_COD_PAGAMENTO) = False Then
                Bdados.CancelaTrans
                Erro "Erro ao gravar item de pagamento"
                Exit Sub
            End If
            
        End If
    Loop
    
    If Bdados.Executa("UPDATE TAB_PAGAMENTO_RECEBIDO SET TPR_TOTAL_PAGAMENTOS = " & TOTAL_PAGAMENTOS & " WHERE TRP_ARQUIVO = '" & TRP_ARQUIVO & "'") = False Then
        Bdados.CancelaTrans
        Erro "Erro ao atualizar total pagamentos na tabela de pagamentos"
        Exit Sub
    End If
    
    LblTotalRegistro = TOTAL_PAGAMENTOS
    Close Arquivo
    
    sql = " SELECT Lote,"
    sql = sql & " Código, "
    sql = sql & " [Data Pagamento], "
    sql = sql & " [Data Vencimento], "
    sql = sql & " [Valor Titulo],"
    sql = sql & " Desconto,"
    sql = sql & " Juros,"
    sql = sql & " [Valor Pago],"
    sql = sql & " [Pagamento Existe],"
    sql = sql & " Ocorrência ,"
    sql = sql & " [Motivo 01], "
    sql = sql & " [Motivo 02] , "
    sql = sql & " [Motivo 03] , "
    sql = sql & " [Motivo 04],"
    sql = sql & " [Motivo 05],[Despesa Cobrança]"
    sql = sql & " From VIS_ITEM_PAGAMENTO"
    sql = sql & " where Arquivo = '" & TPI_TPR_ARQUIVO & "'"
    
    grdDocumentos.Preencher Bdados, sql, 0, 1000, 2300, 2300, 2300, 2000, 1500, 2300, 2500, 2300, 3000, 3000, 3000, 3000, 3000
    For Arquivo = 1 To grdDocumentos.ListItems.Count
        LblTotalTitulo = FormatNumber(CCur(LblTotalTitulo) + CCur(Nvl(0 & grdDocumentos.ListItems(Arquivo).SubItems(4), 0)), 2)
        LblTotalDesconto = FormatNumber(CCur(LblTotalDesconto) + CCur(Nvl(0 & grdDocumentos.ListItems(Arquivo).SubItems(5), 0)), 2)
        LblTotalJuros = FormatNumber(CCur(LblTotalJuros) + CCur(Nvl(0 & grdDocumentos.ListItems(Arquivo).SubItems(6), 0)), 2)
        LblTotalGeral = FormatNumber(CCur(LblTotalGeral) + CCur(Nvl(0 & grdDocumentos.ListItems(Arquivo).SubItems(7), 0)), 2)
        
        If grdDocumentos.ListItems(Arquivo).SubItems(9) <> "" Then
            Ocorrencia = Format(Trim(Left(grdDocumentos.ListItems(Arquivo).SubItems(9), InStr(grdDocumentos.ListItems(Arquivo).SubItems(9), "-") - 1)), "00")
            If Ocorrencia = "06" Or Ocorrencia = "09" Or Ocorrencia = "10" Or Ocorrencia = "15" Or Ocorrencia = "17" Then
                Documento_Pago = True
            End If
        End If
    Next
    'Bdados.GravaTrans
    'Bdados.AbreTrans
    If Documento_Pago Then
        'Verifico se o o lote já foi aberto para esse arquivo, caso tenha sido aberto pego o numero na tab_pagamento_recebido.
        Dim Rs_Consulta_Lote As VSRecordset
        If Bdados.AbreTabela("SELECT TPR_NUMERO_LOTE FROM TAB_PAGAMENTO_RECEBIDO WHERE TRP_ARQUIVO = '" & TRP_ARQUIVO & "'", Rs_Consulta_Lote) Then
            If Nvl("" & Rs_Consulta_Lote("TPR_NUMERO_LOTE"), 0) > 0 Then
                NUMERO_LOTE = Rs_Consulta_Lote.Fields("TPR_NUMERO_LOTE")
                lblLote = Rs_Consulta_Lote.Fields("TPR_NUMERO_LOTE")
            Else
                NUMERO_LOTE = Gera_Lote_Pagamento
                lblLote = NUMERO_LOTE
            End If
        End If
        
        'GRAVO OS PAGAMENSTO NA TAB_BAIXA_RECEBIMENTO
        
        For Arquivo = 1 To grdDocumentos.ListItems.Count
            Ocorrencia = Format(Trim(Left(grdDocumentos.ListItems(Arquivo).SubItems(9), InStr(grdDocumentos.ListItems(Arquivo).SubItems(9), "-") - 1)), "00")
            If Ocorrencia = "06" Or Ocorrencia = "09" Or Ocorrencia = "10" Or Ocorrencia = "15" Or Ocorrencia = "17" Then
                Campos = " TBR_TCR_CODIGO,"
                Campos = Campos & " TBR_ORDEM,"
                Campos = Campos & " TBR_OPERACAO,"
                Campos = Campos & " TBR_VALOR_PAGO,"
                Campos = Campos & " TBR_MULTA,"
                Campos = Campos & " TBR_JUROS,"
                Campos = Campos & " TBR_DESCONTO,"
                Campos = Campos & " TBR_SUB_TOTAL,"
                Campos = Campos & " TBR_FORMA_PAGAMENTO,"
                Campos = Campos & " TBR_NUMERO_DOCUMENTO,"
                Campos = Campos & " TBR_DATA_PAGAMENTO,"
                Campos = Campos & " TBR_TCB_CONTA,"
                Campos = Campos & " TBR_USUARIO,"
                Campos = Campos & " TBR_TIPO_BAIXA,"
                Campos = Campos & " TBR_TLP_LOTE_PAGAMENTO"
                
                Valores = Bdados.PreparaValor(grdDocumentos.ListItems(Arquivo).SubItems(1), _
                1, _
                1, _
                grdDocumentos.ListItems(Arquivo).SubItems(4), _
                0, _
                grdDocumentos.ListItems(Arquivo).SubItems(6), _
                0, _
                grdDocumentos.ListItems(Arquivo).SubItems(7), _
                4, _
                grdDocumentos.ListItems(Arquivo) & "-" & grdDocumentos.ListItems(Arquivo).SubItems(1), _
                grdDocumentos.ListItems(Arquivo).SubItems(2), _
                Pega_Dados_Lote(, NUMERO_LOTE, edlp_TLP_TCB_CONTA), _
                Aplica.Usuario, _
                2, _
                NUMERO_LOTE)
                'Verifico se o pagamento existe, caso não exista não jogo lixo na tab_baixa_recebimento.
                If Bdados.AbreTabela("Select * from tab_conta_receber where tcr_cod_conta = " & grdDocumentos.ListItems(Arquivo).SubItems(1)) Then
                    If Bdados.GravaDados("TAB_BAIXA_RECEBIMENTO", Valores, Campos, "TBR_TCR_CODIGO = " & Bdados.Tabela.Fields("TCR_COD_CONTA")) = False Then
                        Bdados.CancelaTrans
                        Erro "Erro ao gravar baixa"
                        Exit Sub
                    End If
                    'Atulizo o lote
                    'Atualiza_Valor_Lote NUMERO_LOTE, "+" & grdDocumentos.ListItems(Arquivo).SubItems(7)
                    Atualiza_Lote_Automatico NUMERO_LOTE
                             
                    'Altero o valor pago e o saldo devedo da nota
                    Atualiza_Debitos grdDocumentos.ListItems(Arquivo).SubItems(1), grdDocumentos.ListItems(Arquivo).SubItems(2)
                    'Se O Saldo devedor for maior que 0, o status da conta continuará em aberto.
                    If PegaSaldoDevedorDebito(Bdados.Tabela.Fields("TCR_COD_CONTA")) > 0 Then
                        MudaStatusDebito Bdados.Tabela.Fields("TCR_COD_CONTA"), esrAberto
                    Else
                        MudaStatusDebito Bdados.Tabela.Fields("TCR_COD_CONTA"), esrQuitado
                    End If
                    
                    'If Bdados.GravaDados("TAB_CONTA_RECEBER", Bdados.PreparaValor(esrAberto, grdDocumentos.ListItems(Arquivo).SubItems(7), 0, grdDocumentos.ListItems(Arquivo).SubItems(6), 0, grdDocumentos.ListItems(Arquivo).SubItems(5)), "TCR_STATUS,TCR_VALOR_PAGO,TCR_SALDO_DEVEDOR,TCR_JUROS,TCR_MULTA,TCR_DESCONTO", "TCR_COD_CONTA = " & grdDocumentos.ListItems(Arquivo).SubItems(1)) = False Then
                    '    Bdados.CancelaTrans
                    '    Erro "Erro ao atualizar totais da nota"
                    '    Exit Sub
                    'End If
                    'Digo que o pagamento existe
                    Bdados.Executa "Update tab_item_pagamento set TPI_PAGAMENTO_EXISTE = 1 where  TPI_TPR_ARQUIVO  = '" & TPI_TPR_ARQUIVO & "' and TPI_TPR_NUMERO_LOTE = " & NUMERO_LOTE & " and  TPI_COD_PAGAMENTO = " & grdDocumentos.ListItems(Arquivo).SubItems(1)
                    grdDocumentos.ListItems(Arquivo).SubItems(8) = "Sim"
                Else
                    Bdados.Executa "Update tab_item_pagamento set TPI_PAGAMENTO_EXISTE = 2 where  TPI_TPR_ARQUIVO  = '" & TPI_TPR_ARQUIVO & "' and TPI_TPR_NUMERO_LOTE = " & NUMERO_LOTE & " and  TPI_COD_PAGAMENTO = " & grdDocumentos.ListItems(Arquivo).SubItems(1)
                    grdDocumentos.ListItems(Arquivo).SubItems(8) = "Não"
                End If
             End If
            Next
        'Fecho o lote
        Campos = "TLP_DATA_FECHAMENTO,TLP_USUARIO_FECHAMENTO,TLP_SITUACAO_LOTE,TLP_VALOR_LOTE,TLP_DATA_ARRECADACAO,TLP_VALOR_ARRECADADO"
        Valores = Bdados.PreparaValor(Date, Aplica.Usuario, sL_Lote_Fechado, LblTotalGeral, TPI_DATA_PAGAMENTO, LblTotalGeral)
        If Bdados.GravaDados("TAB_LOTE_PAGAMENTO", Valores, Campos, "TLP_COD_LOTE = " & NUMERO_LOTE) = False Then
            Bdados.CancelaTrans
            Erro "Erro ao fechar lote"
            Exit Sub
        End If
        'Atualizo o número do lote de pagamento.
        Campos = "TPR_NUMERO_LOTE"
        Valores = Bdados.PreparaValor(NUMERO_LOTE)
        'Se esse processo for executato, então considero que tudo deu certo
        If Bdados.GravaDados("TAB_PAGAMENTO_RECEBIDO", Valores, Campos, "TRP_ARQUIVO = '" & TRP_ARQUIVO & "'") = True Then
            Bdados.GravaTrans
            Avisa "Retorno recebido com sucesso"
        Else
            Bdados.CancelaTrans
            Avisa "Erro no final do processo, lote não pode ser fechado"
        End If
        Exit Sub
        
    Else
        'Gravo os documentos
        Bdados.GravaTrans
        Avisa "Retorno recebido com sucesso"
        Avisa "Arquivo não possue pagamentos"
        Exit Sub
    End If
    Avisa "Retorno não recebido"
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabCabecalho.Exibir Bdados, Me.Name, App.Path
    rodRodape.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Set Arquivo = New Arquivo
 CboArquivos.Preencher Bdados, "SELECT distinct(trp_arquivo)FROM TAB_PAGAMENTO_RECEBIDO"
End Sub


Private Sub grdArquivos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sql As String
    
    If grdArquivos.ListItems.Count >= 1 Then
        sql = " SELECT Lote,"
        sql = sql & " Código, "
        sql = sql & " [Data Pagamento], "
        sql = sql & " [Data Vencimento], "
        sql = sql & " [Valor Titulo],"
        sql = sql & " Desconto,"
        sql = sql & " Juros,"
        sql = sql & " [Valor Pago],"
        sql = sql & " [Pagamento Existe],"
        sql = sql & " Ocorrência ,"
        sql = sql & " [Motivo 01], "
        sql = sql & " [Motivo 02] , "
        sql = sql & " [Motivo 03] , "
        sql = sql & " [Motivo 04],"
        sql = sql & " [Motivo 05],[Despesa Cobrança]"
        sql = sql & " From VIS_ITEM_PAGAMENTO"
        sql = sql & " where Arquivo = '" & grdArquivos.SelectedItem & "'"
        GrdItemPagamentos.Preencher Bdados, sql, 0, 1000, 2300, 2300, 2300, 2000, 1500, 2300, 2500, 2300, 3000, 3000, 3000, 3000, 3000
    End If
End Sub

Private Sub Atualiza_Debitos(Conta As String, DataBaixa As String)
    
    Dim rs As VSRecordset
    Dim Desconto As Double
    Dim Tipo As String
    Dim txtDataBaixa As String
    Dim txtVencimento As String
    Dim RsBase As VSRecordset
    Dim txtSaldoDevedor As String
    Dim txtNotaBaixa As String
    Dim txtValorNotaBaixa As String
    Dim txtValorApagar  As String
    Dim txtDescontoOriginal  As String
    Dim txtValorPago As String
    Dim txtJuros  As String
    Dim txtMulta As String
    Dim sql As String
    
    sql = "select * from vis_conta_receber where Código = " & Conta
    
    If Bdados.AbreTabela(sql, RsBase) Then
        Tipo = 0
        Do Until RsBase.EOF
        'If RsBase.Fields("Código") = 12067726 Then
        '    'MsgBox "oi"
        'End If
        'Debug.Print RsBase.Fields("código")
            txtDataBaixa = DataBaixa
            txtVencimento = RsBase.Fields("Vencimento")
            txtSaldoDevedor = RsBase.Fields("Saldo Devedor")
            txtNotaBaixa = "" & RsBase.Fields("Matricula")
            txtValorNotaBaixa = "" & RsBase.Fields("Valor")
            txtValorPago = "" & RsBase.Fields("VAlor Pago")
            
                If txtNotaBaixa <> "" Then
                    sql = "Select * "
                    sql = sql & " from tab_matricula_desconto "
                    sql = sql & " where TMD_TMA_MATRICULA  =  " & txtNotaBaixa
                    'sql = sql & " and TMD_TIPO IN (6)"
                    If Bdados.AbreTabela(sql, rs) Then
                        Tipo = rs.Fields("TMD_TIPO")
                        Desconto = rs.Fields("tmd_valor")
                       If rs.Fields("TMD_TIPO") = 2 And CDate(txtDataBaixa) > CDate(txtVencimento) Then
                        Exit Sub
                       End If
                    End If
                End If
                If CDate(txtDataBaixa) <= CDate(txtVencimento) Then
                    If Tipo = "0" Then
                        txtValorApagar = (CCur(txtValorNotaBaixa) - ((PegaConfiguracaoEscola(TEC_VALOR_PRIMEIRO_DESCONTO) * CCur(txtValorNotaBaixa)) / 100))
                        txtDescontoOriginal = ((PegaConfiguracaoEscola(TEC_VALOR_PRIMEIRO_DESCONTO) * CCur(txtValorNotaBaixa)) / 100)
                        txtSaldoDevedor = Format(CCur(txtValorApagar) - CCur(txtValorPago), Const_Monetario)
                        txtJuros = 0
                        txtMulta = 0
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtDescontoOriginal, txtValorApagar, CCur(txtSaldoDevedor) + CCur(txtJuros) + CCur(txtMulta)), "TCR_DESCONTO,TCR_VALOR_APAGAR,TCR_SALDO_DEVEDOR", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtJuros, txtMulta), "TCR_JUROS,TCR_MULTA", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                    Else
                        txtValorApagar = (CCur(txtValorNotaBaixa) - ((Desconto * CCur(txtValorNotaBaixa)) / 100))
                        txtDescontoOriginal = ((Desconto * CCur(txtValorNotaBaixa)) / 100)
                        txtSaldoDevedor = Format(CCur(txtValorApagar) - CCur(txtValorPago), Const_Monetario)
                        txtJuros = 0
                        txtMulta = 0
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtDescontoOriginal, txtValorApagar, CCur(txtSaldoDevedor) + CCur(txtJuros) + CCur(txtMulta)), "TCR_DESCONTO,TCR_VALOR_APAGAR,TCR_SALDO_DEVEDOR", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtJuros, txtMulta), "TCR_JUROS,TCR_MULTA", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                    End If
                ElseIf Day(Nvl(txtDataBaixa, 0)) > Val(Nvl(PegaConfiguracaoEscola(TEC_DIA_PRIMEIRO_DESCONTO), 0)) And Day(txtDataBaixa) <= Val(Nvl(PegaConfiguracaoEscola(TEC_DIA_SEGUNDO_DESCONTO), 0)) Then
                    If CDate(txtDataBaixa) > CDate(txtVencimento) And Month(txtDataBaixa) = Month(txtVencimento) Then
                        txtValorApagar = (CCur(txtValorNotaBaixa) - ((PegaConfiguracaoEscola(TEC_VALOR_SEGUNDO_DESCONTO) * CCur(txtValorNotaBaixa)) / 100))
                        txtDescontoOriginal = ((PegaConfiguracaoEscola(TEC_VALOR_SEGUNDO_DESCONTO) * CCur(txtValorNotaBaixa)) / 100)
                        txtSaldoDevedor = Format(CCur(txtValorApagar) - CCur(txtValorPago), Const_Monetario)
                        txtJuros = 0
                        txtMulta = 0
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtDescontoOriginal, txtValorApagar, CCur(txtSaldoDevedor) + CCur(txtJuros) + CCur(txtMulta)), "TCR_DESCONTO,TCR_VALOR_APAGAR,TCR_SALDO_DEVEDOR", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtJuros, txtMulta), "TCR_JUROS,TCR_MULTA", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                    Else
                        txtValorApagar = txtValorNotaBaixa
                        txtDescontoOriginal = "0,00"
                        txtSaldoDevedor = Format(CCur(txtValorApagar) - CCur(Nvl(txtValorPago, 0)), Const_Monetario)
                        txtJuros = CalculaValoresJurosAvulsos(txtDataBaixa, txtVencimento, CDbl(txtSaldoDevedor), PegaConfiguracaoEscola(Juros))
                        txtMulta = CalculaValoresMultaAvulsos(txtDataBaixa, txtVencimento, CDbl(txtSaldoDevedor), PegaConfiguracaoEscola(MIN_MULTA), PegaConfiguracaoEscola(max_MULTA), Normal)
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtDescontoOriginal, txtValorApagar, CCur(txtSaldoDevedor) + CCur(txtJuros) + CCur(txtMulta)), "TCR_DESCONTO,TCR_VALOR_APAGAR,TCR_SALDO_DEVEDOR", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                        Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtJuros, txtMulta), "TCR_JUROS,TCR_MULTA", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                    End If
                Else
'Vai:
                    If Tipo = 3 Then
                        txtDescontoOriginal = "0,00"
                    End If
                    
                    
                    txtDescontoOriginal = 0
                    txtValorApagar = txtValorNotaBaixa
                    txtSaldoDevedor = Format(CCur(txtValorApagar) - CCur(Nvl(txtValorPago, 0)), Const_Monetario)
                    txtJuros = CalculaValoresJurosAvulsos(txtDataBaixa, txtVencimento, CDbl(txtSaldoDevedor), PegaConfiguracaoEscola(Juros))
                    txtMulta = CalculaValoresMultaAvulsos(txtDataBaixa, txtVencimento, CDbl(txtSaldoDevedor), PegaConfiguracaoEscola(MIN_MULTA), PegaConfiguracaoEscola(max_MULTA), Normal)
                    Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(Nvl(txtDescontoOriginal, 0), txtValorApagar, CCur(txtSaldoDevedor) + CCur(txtJuros) + CCur(txtMulta)), "TCR_DESCONTO,TCR_VALOR_APAGAR,TCR_SALDO_DEVEDOR", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                    Bdados.GravaDados "TAB_CONTA_RECEBER", Bdados.PreparaValor(txtJuros, txtMulta), "TCR_JUROS,TCR_MULTA", "TCR_COD_CONTA = " & RsBase.Fields("Código")
                    Bdados.Executa "update tab_baixa_recebimento set tbr_operacao = 3 where tbr_tcr_codigo = " & RsBase.Fields("Código")
                    
                    
                    
                    
                End If
            
            RsBase.MoveNext
        Loop
   End If
End Sub


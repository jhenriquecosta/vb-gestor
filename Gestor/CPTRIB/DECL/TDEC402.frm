VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDEC402"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Tag             =   "TDEC402"
   Begin VTOcx.fraVISUAL fraNotaFiscal 
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1980
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2778
      Altura          =   1905
      Caption         =   " Nota Fiscal"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Begin VTOcx.cboVISUAL cboMostrar 
         Height          =   510
         Left            =   90
         TabIndex        =   13
         Top             =   930
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   900
         Caption         =   "Mostrar"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboTipoNota 
         Height          =   510
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   900
         Caption         =   "Tipo Operação"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtTMAliq 
         Height          =   525
         Left            =   8730
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   926
         Caption         =   "Aliquota"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   3
         AlinhamentoRotulo=   1
         AlinhamentoTexto=   1
         CorFundo        =   14737632
      End
      Begin VTOcx.txtVISUAL txtDataEmissao 
         Height          =   495
         Index           =   0
         Left            =   5700
         TabIndex        =   11
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         Caption         =   "Data Emissão"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtNumNota 
         Height          =   495
         Left            =   4290
         TabIndex        =   10
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Caption         =   "Nº Nota Fiscal"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtInscricaoNota 
         Height          =   495
         Left            =   2220
         TabIndex        =   9
         Top             =   360
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   873
         Caption         =   "CPF/CNPJ"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDataEmissao 
         Height          =   495
         Index           =   1
         Left            =   6870
         TabIndex        =   12
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         Caption         =   ""
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         AlinhamentoTexto=   1
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Height          =   525
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4020
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmLimpar 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   6210
         TabIndex        =   17
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   688
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         CausesValidation=   0   'False
         Height          =   390
         Left            =   7260
         TabIndex        =   18
         Top             =   105
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   688
         Caption         =   "Sair"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   390
         Left            =   5160
         TabIndex        =   16
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   688
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1230
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   675
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2170
      Altura          =   1905
      Caption         =   " Declaração"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   5700
         TabIndex        =   6
         Top             =   780
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   780
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Index           =   1
         Left            =   1950
         TabIndex        =   4
         Top             =   780
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   780
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         Caption         =   "Período"
         Text            =   ""
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3090
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   390
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   503
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIM 
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   390
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   503
         Caption         =   "Insc. Municipal"
         Text            =   ""
         Restricao       =   2
         RetirarMascara  =   0   'False
      End
   End
   Begin VTOcx.cboVISUAL cboRelatorio 
      Height          =   315
      Left            =   210
      TabIndex        =   14
      Top             =   3630
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      Caption         =   "Relatório"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   1138
      Icone           =   "TDEC402.frx":0000
   End
End
Attribute VB_Name = "TDEC402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboRelatorio_Click()
    If cboRelatorio = "DECLARAÇÕES GERADAS" Then
        fraNotaFiscal.Visible = False
        cboRelatorio.Top = fraNotaFiscal.Top
    Else
        fraNotaFiscal.Visible = True
        cboRelatorio.Top = 3630
    End If
    rodVISUAL1.Top = cboRelatorio.Top + 390
    Me.Height = rodVISUAL1.Top + 945
End Sub

Private Sub cmdBuscar_Click()
    Dim Relatorio As Object
    Dim Sql As String
    Dim criterio As String
    
    If cboRelatorio = "" Then
        Util.Avisa "Por favor, selecione um relatório"
        cboRelatorio.SetFocus
        Exit Sub
    End If
    
    'DECLARAÇÃO
    If txtIM <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TDC_TCI_IM='" & txtIM & "'"
    End If
    If txtPeriodo(0) <> "" And txtPeriodo(1) <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "MES_PERIODO BETWEEN " & Left(txtPeriodo(0), 2) & " AND " & Left(txtPeriodo(1), 2)
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "ANO_PERIODO BETWEEN " & Right(txtPeriodo(0), 4) & " AND " & Right(txtPeriodo(1), 4)
    End If
    If cboTipo <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TDC_TIPO_DEC=" & cboTipo.Coluna(1).Valor
    End If
    If cboStatus <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TDC_STATUS=" & cboStatus.Coluna(1).Valor
    End If
    
    'NOTA FISCAL
    'LUCAS
'    If fraNotaFiscal.Visible = False Then GoTo PEGA_RELATORIO
    If cboTipoNota <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TNF_COD_OPERACAO=" & cboTipoNota.ListIndex + 1
    End If
    If txtInscricaoNota <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TNF_INSCRICAO_OPERACAO='" & txtInscricaoNota & "'"
    End If
    If txtNumNota <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TNF_NUM_NOTA='" & txtNumNota & "'"
    End If
    If txtDataEmissao(0) <> "" And txtDataEmissao(1) <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TNS_DATA_NOTA BETWEEN " & Bdados.Converte(txtDataEmissao(0), TCDataHora) & " AND " & Bdados.Converte(txtDataEmissao(1), TCDataHora)
    End If
    If cboMostrar <> "" Then
        If criterio <> "" Then criterio = criterio & " and "
        criterio = criterio & "TNS_NOTA_CANCELADA=" & cboMostrar.ListIndex
    End If
'LUCAS
'PEGA_RELATORIO:
'    Select Case UCase(cboRelatorio.Text)
'        Case UCase("Declarações Criadas")
'            Set Relatorio = New AR_DeclaracoesGeradas
'            Sql = "SELECT * FROM VIS_DECLARACAO"
'        Case UCase("Declarações Notas Fiscais")
'            Set Relatorio = New AR_DeclaracoesNotasFiscais
'            Sql = "SELECT * FROM VIS_DECLARACAO_NOTA_FISCAL"
'        Case UCase("Entrega de Declaração")
'            Set Relatorio = New AR_Declaracao
'            Sql = "SELECT * FROM VIS_DECLARACAO_NOTA_FISCAL"
'        Case Else
'            Util.Avisa "Relatório Inválido"
'            Exit Sub
'    End Select
    
    If criterio <> "" Then Sql = Sql & " WHERE " & criterio
    
    VisualizarActiveReport Relatorio, Bdados, Sql
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmLimpar_Click()
    Dim RelatorioAtual As Integer
    RelatorioAtual = cboRelatorio.ListIndex
    LimpaCampos Me
    cboRelatorio.ListIndex = RelatorioAtual
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If cboRelatorio = "" Then cboRelatorio.ListIndex = 0
End Sub

Private Sub Form_Load()
    
    cabVISUAL1.Exibir Bdados, Me.Tag, App.Path
    rodVISUAL1.Exibir Bdados, Me.Tag
    
    cboMostrar.AddItem "Notas Não Canceladas"
    cboMostrar.AddItem "Notas Canceladas"
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
    cboStatus.PreencherGeral Bdados, "STATUS DECLARACAO"
    cboRelatorio.PreencherGeral Bdados, "RELATORIO DECLARACAO"
    cboTipoNota.PreencherGeral Bdados, "OPERACAO NOTA FISCAL"
End Sub

Private Sub txtIM_Change()
    txtRazao = ""
End Sub

Private Sub txtIM_Validate(Cancel As Boolean)
    If Trim$(txtIM) <> "" Then
        If Not BuscarContribuinte(txtIM, txtRazao) Then
            Avisa "Contribuinte não encontrado."
            txtIM = ""
            Cancel = True
        End If
    End If
End Sub

Private Sub txtInscricaoNota_Validate(Cancel As Boolean)
    'BuscaContribuinte txtInscricaoNota
End Sub

Private Sub txtPeriodo_LostFocus(Index As Integer)
    If Len(Trim(txtPeriodo(Index))) <> 6 Then Exit Sub
    txtPeriodo(Index) = Left(txtPeriodo(Index), 2) & "/" & Right(txtPeriodo(Index), 4)
End Sub

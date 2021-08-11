VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Begin VB.Form formObrigConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTodos 
      Caption         =   "Check1"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   6600
      Width           =   255
   End
   Begin VB.Frame FraMenagem 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1950
      TabIndex        =   25
      Top             =   6405
      Width           =   4110
      Begin VB.Label LblMensagem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRIBUINTE NOTIFICADO EM 01/01/2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   75
         TabIndex        =   26
         Top             =   30
         Width           =   4005
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   7035
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdRelatorio 
         Height          =   375
         Left            =   6330
         TabIndex        =   10
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   661
         Caption         =   "&Relatorio"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   7545
         TabIndex        =   9
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         Caption         =   "&Imprimir DAM"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10440
         TabIndex        =   12
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
         Left            =   9270
         TabIndex        =   11
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
      Height          =   2400
      Index           =   3
      Left            =   30
      TabIndex        =   16
      Top             =   630
      Width           =   11580
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   390
         TabIndex        =   17
         Top             =   930
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboRestricao 
         Height          =   315
         Left            =   690
         TabIndex        =   2
         Tag             =   "Tributo"
         Top             =   1605
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   556
         Caption         =   "Restrição"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6540
         TabIndex        =   3
         Tag             =   "Tributo"
         ToolTipText     =   "746 - STATUS OBRIGACAO"
         Top             =   1590
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   10440
         TabIndex        =   8
         Top             =   1965
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   582
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   690
         TabIndex        =   20
         Top             =   1260
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3390
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   540
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   270
         TabIndex        =   4
         Tag             =   "Periodo Inicial"
         Top             =   1980
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
         Left            =   2730
         TabIndex        =   5
         Tag             =   "Periodo Final"
         Top             =   1980
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
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   8340
         TabIndex        =   7
         Top             =   1980
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   5190
         TabIndex        =   6
         Tag             =   "Periodo Inicial"
         Top             =   1980
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         Caption         =   "Periodo(dd/mm/aaaa)"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   3870
         TabIndex        =   0
         Top             =   540
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
         Left            =   7350
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   540
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtDAM 
         Height          =   300
         Left            =   8040
         TabIndex        =   1
         Top             =   540
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
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
         TabIndex        =   27
         Tag             =   "Tributo"
         Top             =   135
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   705
         TabIndex        =   28
         Top             =   540
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   529
         Caption         =   "Inscricão"
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
         TabIndex        =   18
         Top             =   1590
         Width           =   45
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   1138
      Icone           =   "TOBR401.frx":0000
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   13
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
      TabIndex        =   21
      Top             =   5490
      Visible         =   0   'False
      Width           =   795
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1620
      Left            =   45
      TabIndex        =   23
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
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   3615
      Left            =   30
      TabIndex        =   14
      Top             =   3045
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6376
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtEnderecoContrib 
      Height          =   300
      Left            =   0
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "Periodo Final"
      Top             =   0
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
   Begin VTOcx.txtVISUAL txtVISUAL1 
      Height          =   300
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   529
      Caption         =   "Inscricão"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VB.Label LblOBS 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1995
      TabIndex        =   31
      Top             =   6750
      Visible         =   0   'False
      Width           =   9600
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
         Caption         =   "Consultar Processo de DAT"
         Index           =   5
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Dados do Pagamento"
         Index           =   6
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Notas Fiscais"
         Index           =   7
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Cancelar"
         Index           =   9
      End
   End
End
Attribute VB_Name = "TOBR401"
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
Dim InscProprietario As String
Public String_Taxas    As String
Public Total_Taxas     As Double
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

Private Sub chkTodos_Click()
    Dim b As Boolean
    Dim i As Integer
    b = chkTodos.Value
    For i = 1 To lstObrig.ListItems.Count
        With lstObrig.ListItems
            .Item(i).Checked = b
        End With
    Next i
End Sub

Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Dim Inscri As String
    Set Obrig = New Obrigacao
    
    Inscri = txtIM
     
    'If Trim(txtIm) <> "" Then Conta.ExecutaAtualizacao txtIm
    If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), Inscri, _
            CInt(cboRestricao.Coluna(1).Valor), CInt(cboStatus.Coluna(1).Valor), txtPeriodoInicial, txtPeriodoFinal, _
            txtExercicioInicial, txtExercicioFinal, , txtImovel, , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), txtDAM) Then
        Avisa "Nenhum registro encontrado."
 cboImposto.SetFocus
    End If
    If txtIM <> "" Then
        Inscri = txtIM
    Else
        Inscri = txtImovel
    End If
    If Bdados.AbreTabela("SELECT MAX(TNT_ENTREGA) as data,COUNT(*) as Total FROM TAB_NOTIFICACAO WHERE TNT_INSCRICAO = '" & Inscri & "'AND NOT TNT_ENTREGA IS NULL") Then
        If Not IsNull(Bdados.Tabela(0)) Then
            FraMenagem.Visible = True
            LblMensagem.Visible = True
            LblMensagem.Caption = "CONTRIBUINTE NOTIFICADO EM " & Bdados.Tabela(0) & " - TOTAL DE NOTIFICAÇÕES: " & Bdados.Tabela(1)
            FraMenagem.Width = LblMensagem.Width + 100
        Else
            FraMenagem.Visible = False
            LblMensagem.Visible = False
        End If
    Else
        FraMenagem.Visible = False
        LblMensagem.Visible = False
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


Private Sub cmdImprimir_Click()
On Error GoTo trata
    
    If lstObrig.ListItems.Count = 0 Then Exit Sub
    
    If lstObrig.ListItems.Count > 1 Then
        If Not Util.Confirma("Confirma impressão de " & lstObrig.ListItems.Count & " obrigações") Then Exit Sub
    Else
        If Not Util.Confirma("Confirma impressão da obrigação") Then Exit Sub
    End If

    Screen.MousePointer = 11
    Dim Inscricao As String, Razao As String, Endereco As String, sqlContrib As String
    Dim rs As VSRecordset
    Dim i As Double
    
    For i = 1 To lstObrig.ListItems.Count
        With lstObrig.ListItems
            .Item(i).Selected = True
            Inscricao = .Item(i).SubItems(1)
            sqlContrib = "select tci_im,tci_cgc_cpf,tci_nome,tci_logradouro,tci_nome_logradouro,tci_numero,tci_bairro,tci_cep FROM  Tab_Contribuinte where tci_im='" & Inscricao & "'"
            If Bdados.AbreTabela(sqlContrib, rs) Then
                Razao = rs("tci_nome")
                Endereco = rs("tci_logradouro") & " " & rs("tci_nome_logradouro") & "," & rs("tci_numero") & "-" & rs("tci_bairro")
                txtEndereco = Endereco
            End If
            Call Pega_taxas
            If .Item(i).Checked Then
                'NOVO METODO PARA EXIBIR O NOME DO CONTRIBUINTE NO DAM - GLEYSON - 13/05/2011
                ExibeContribuinte lstObrig.ListItems(i).SubItems(1)
                If Trim(txtImovel) = "" Then
                    ImprimeSelecionado lstObrig, Razao, Endereco, False, , tdiImpressora, String_Taxas, Total_Taxas, Inscricao, txtEndereco
                Else
                    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, False, , tdiImpressora, String_Taxas, Total_Taxas, InscProprietario, txtEnderecoContrib
                End If
            End If
            txtEndereco = ""
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

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIM
End Sub

Private Sub cmdRelatorio_Click()
    Dim CondRelatorio As String
    Dim FORMULA As String
    Dim Inscricao As String
    
    
    On Error GoTo trata
    
    If txtIM <> "" Then
        Inscricao = txtIM
        CondRelatorio = " ({TAB_OBRIGACAO_CONTRIBUINTE.TOC_TIPO_INSCRICAO} = 2 "
        Conta.ExecutaAtualizacao Trim(Inscricao), etiContribuinte, True, , , , , , , , "" & cboImposto.Coluna(0).Valor, Nvl("" & cboStatus.Coluna(1).Valor, 1), txtExercicioInicial, txtExercicioFinal
    ElseIf txtImovel <> "" Then
        Inscricao = txtImovel
        CondRelatorio = " ({TAB_OBRIGACAO_CONTRIBUINTE.TOC_TIPO_INSCRICAO} = 1"
        Conta.ExecutaAtualizacao Trim(Inscricao), etiImovel, True, , , , , , , , "" & cboImposto.Coluna(0).Valor, Nvl("" & cboStatus.Coluna(1).Valor, 1), txtExercicioInicial, txtExercicioFinal
    Else
        CondRelatorio = "1 = 1"
        Conta.ExecutaAtualizacao "", , True, , , , , , , , "" & cboImposto.Coluna(0).Valor, Nvl("" & cboStatus.Coluna(1).Valor, 1), txtExercicioInicial, txtExercicioFinal
    End If
    Screen.MousePointer = 11
    With Rpt
        
        If Not .DefinirArquivo(Bdados, App.Path & "\TObrigacao.rpt") Then Exit Sub
        
        If Trim(Inscricao) <> "" Then CondRelatorio = CondRelatorio & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_INSCRICAO} = '" & Trim(Inscricao) & "')"
        
        If cboImposto.ListIndex <> -1 Then
            CondRelatorio = CondRelatorio & " and {Tab_Imposto.tip_cod_imposto} = '" & cboImposto.Coluna(0).Valor & "'"
        End If
        FORMULA = ""
        If Trim(txtExercicioInicial) <> "" And Trim(txtExercicioFinal) <> "" Then
        
            txtExercicioInicial = IIf(Len(txtExercicioInicial) = 4, txtExercicioInicial, Right(txtExercicioInicial, 4) & Left(txtExercicioInicial, 2))
            txtExercicioFinal = IIf(Len(txtExercicioFinal) = 4, txtExercicioFinal, Right(txtExercicioFinal, 4) & Left(txtExercicioFinal, 2))
  
            CondRelatorio = CondRelatorio & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtExercicioInicial & " and {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <= " & txtExercicioFinal & ""
            FORMULA = txtExercicioInicial & " - " & txtExercicioFinal
        End If
        If cboRestricao.ListIndex <> -1 Then
            If cboRestricao.Coluna(1).Valor = 1 Then
                CondRelatorio = CondRelatorio & " and  ({VIS_STATUS_OBRIGACAO.TGE_CODIGO}  = 2 or {VIS_STATUS_OBRIGACAO.TGE_CODIGO}  = 4 or {VIS_STATUS_OBRIGACAO.TGE_CODIGO}  = 5) "
            ElseIf cboRestricao.Coluna(1).Valor = 2 Then
                CondRelatorio = CondRelatorio & " and {VIS_STATUS_OBRIGACAO.TGE_CODIGO} = 3"
            End If
        End If
        If cboStatus.ListIndex <> -1 Then
            CondRelatorio = CondRelatorio & " and {TAB_OBRIGACAO_CONTRIBUINTE.TOC_STATUS_OBRIGACAO} =" & cboStatus.Coluna(1).Valor
        End If
        If txtIM <> "" Then
        .Formulas "CONTRIBUINTE", txtIM & " - " & txtRazao
        Else
        .Formulas "CONTRIBUINTE", txtImovel & " - " & txtRazao
        End If
        .Formulas "ENDERECO", txtEndereco
        .Selecao = CondRelatorio 'Right(CondRelatorio, Len(CondRelatorio) - 4)
        '.Formulas "FILTRO", FORMULA
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            .Formulas "VT_NOME_DAM", "Relatório de DARM Emitido"
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "GAF")
        Else
            .Formulas "VT_NOME_DAM", "Relatório de DAM Emitido"
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        End If
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        '.Arvore = False
        .Visualizar
    
    End With
    Set Rpt = Nothing
    Avisa "Impressão concluída."
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    Screen.MousePointer = 0
    Erro Err.Description
    Exit Sub
    Resume
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Activate()
    If Left(Trim(Me.Tag), 1) = "C" Then
        txtIM = Mid(Me.Tag, 2)
        txtIm_LostFocus
    ElseIf Left(Trim(Me.Tag), 1) = "I" Then
        txtImovel = Mid(Me.Tag, 2)
        txtImovel_LostFocus
    ElseIf Len(Trim(Me.Tag)) > 0 Then
        txtIM = Me.Tag
        txtIm_LostFocus
    End If
    GrdTaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
    FraMenagem.Visible = False
        LblMensagem.Visible = False
        txtIM.SetFocus
End Sub
Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    cboRestricao.PreencherGeral Bdados, "RESTRICAO DAM"
    
End Sub

Private Sub lstObrig_DblClick()
  Dim Sql As String
    If (lstObrig.SelectedItem Is Nothing) Then Exit Sub
    'Checo se a obrigação foi cancelda...
    Sql = "Select toc_justificativa_alteracao "
    Sql = Sql & " from tab_obrigacao_contribuinte "
    Sql = Sql & " where toc_cod_obrigacao  = " & lstObrig.SelectedItem
    If Bdados.AbreTabela(Sql) Then
        If "" & Bdados.Tabela(0) <> "" Then Avisa "" & Bdados.Tabela(0)
    End If
    If lstObrig.SelectedItem.SubItems(20) <> "" Then
        LblOBS.AutoSize = True
        LblOBS.Visible = True
        LblOBS = "OBSERVAÇÃO :" & lstObrig.SelectedItem.SubItems(20)
    Else
        LblOBS.Visible = False
    End If
End Sub
Private Sub lstObrig_Click()
    If Not lstObrig.SelectedItem Is Nothing Then
        'NOVO METODO PARA EXIBIR O NOME DO CONTRIBUINTE NO DAM - GLEYSON - 13/05/2011
        ExibeContribuinte lstObrig.SelectedItem.SubItems(1)
    End If
End Sub
Private Sub lstObrig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not lstObrig.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuReimprime(0).Caption = "Imprimir DAM da obrigação nº " & lstObrig.SelectedItem.Text
            Me.PopupMenu mnuGeral
        End If
    End If
End Sub
Private Sub ExibeContribuinte(Ic As String)
    If Len(Ic) < 15 Then
        txtImovel = ""
        txtIM = Ic
        txtIM = BuscaContribuinte(Ic, txtRazao, txtEndereco)
    Else
        txtIM = ""
        txtImovel = Ic
        txtImovel = BuscaContribuinte(Ic, txtRazao, txtEndereco, InscProprietario, etiImovel)
    End If
End Sub


Private Sub mnuReimprime_Click(Index As Integer)
    Select Case Index
        Case 0
                If lstObrig.SelectedItem Is Nothing Then Exit Sub
                If Not Cobranca.LiberaImpressaoDam(Nvl(lstObrig.SelectedItem.SubItems(15), 0)) Then Exit Sub
                With lstObrig.SelectedItem
                    NovaData = Imposto.DataVencimentoNova(.SubItems(5))
                    If Trim(NovaData) = "" Then Exit Sub
                End With
                'Pego as taxas
                Call Pega_taxas
'                ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela, String_Taxas, Total_Taxas
                If Trim(txtImovel) = "" Then
                    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, False, NovaData, tdiTela, String_Taxas, Total_Taxas, txtIM, txtEndereco
                Else
                    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, False, NovaData, tdiTela, String_Taxas, Total_Taxas, InscProprietario, txtEnderecoContrib
                End If
          Case 2
                Load TOBR405
                TOBR405.Caption = "TOBR405 - Cotas de Parcelamento"
                TOBR405.cabVISUAL1.Formulario = "Formulário para impressão das cotas de parcelamento."
                TOBR405.txtIM = txtIM
                TOBR405.txtEndereco = txtEndereco
                TOBR405.txtRazao = txtRazao
                TOBR405.txtIM.Enabled = False
                TOBR405.Tag = lstObrig.SelectedItem.SubItems(16)
                TOBR405.Show
                
            Case 3
                Load TOBR405
                TOBR405.cabVISUAL1.Formulario = "Formulário para impressão dos parcelamentos de ofício."
                TOBR405.Caption = "TOBR405 - Parcelamento de Ofício"
                TOBR405.txtIM = txtIM
                TOBR405.txtRazao = txtRazao
                TOBR405.txtIM.Enabled = False
                TOBR405.txtEndereco = txtEndereco
                TOBR405.Tag = lstObrig.SelectedItem
                TOBR405.Show
            Case 5
                Load TOBR406
                TOBR406.Tag = lstObrig.SelectedItem
                TOBR406.Show 1
            Case 6
                Load TOBR407
                TOBR407.Tag = lstObrig.SelectedItem
                TOBR407.Show 1
            Case 7
                Load TOBR410
                TOBR410.Tag = Trim(lstObrig.SelectedItem.SubItems(1)) & "/" & Trim(lstObrig.SelectedItem.SubItems(2)) & "/" & Trim(lstObrig.SelectedItem.SubItems(11))
                TOBR410.Show 1
    End Select
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If Len(txtIM) = 10 Or Len(txtIM) = 11 Then
            Ic = Imposto.FormataInscricao(txtIM, InscContrib)
        Else
            Ic = txtIM
        End If
    Else
        Ic = txtIM
    End If
    If Trim(txtIM) <> "" Then
        txtIM = BuscaContribuinte(Ic, txtRazao, txtEndereco)
        If Trim(txtIM) = "" Then
            Avisa "Inscricão não encontrada"
            txtRazao = ""
            txtEndereco = ""
            
            txtIM.SetFocus
        End If
    End If
    
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
            txtRazao = ""
            txtEndereco = ""
            txtImovel.SetFocus
        End If
    End If
End Sub

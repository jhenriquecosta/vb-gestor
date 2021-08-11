VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   18
      Top             =   7125
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdRelatorio 
         Height          =   375
         Left            =   4995
         TabIndex        =   22
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
         Left            =   3285
         TabIndex        =   19
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
         Left            =   7380
         TabIndex        =   11
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
         Left            =   6210
         TabIndex        =   10
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
      Height          =   2640
      Index           =   3
      Left            =   30
      TabIndex        =   15
      Top             =   690
      Width           =   8460
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   150
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   705
         TabIndex        =   1
         Top             =   510
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Inscrição"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   840
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   529
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Tag             =   "Periodo Inicial"
         Top             =   1500
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         Caption         =   "Exercício"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboRestricao 
         Height          =   315
         Left            =   690
         TabIndex        =   7
         Tag             =   "Tributo"
         Top             =   1875
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   556
         Caption         =   "Restrição"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtExercicioFinal 
         Height          =   300
         Left            =   2580
         TabIndex        =   4
         Tag             =   "Periodo Inicial"
         Top             =   1500
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   3810
         TabIndex        =   5
         Tag             =   "Periodo Inicial"
         Top             =   1515
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         Caption         =   "Período(ddmmaaaa)"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   6945
         TabIndex        =   6
         Tag             =   "Periodo Inicial"
         Top             =   1515
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   930
         TabIndex        =   8
         Tag             =   "Tributo"
         Top             =   2235
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
         Left            =   7365
         TabIndex        =   9
         Top             =   2235
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
         Top             =   1170
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtGrupo 
         Height          =   300
         Left            =   3735
         TabIndex        =   2
         Top             =   510
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
         Caption         =   "Grupo(ddssqqqq)"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MaxLen          =   8
         MinLen          =   2
         AutoTAB         =   -1  'True
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   17
         Top             =   1590
         Width           =   45
      End
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   3660
      Left            =   30
      TabIndex        =   13
      Top             =   3390
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   6456
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TOBR401.frx":0000
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
   Begin VB.Menu mnuGeral 
      Caption         =   "Geral"
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "Reimprime"
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

Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Dim Inscri As String
    Set Obrig = New Obrigacao
    
    Inscri = txtIm
    If Trim(txtGrupo) <> "" Then
        Inscri = txtGrupo
    End If
    If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), Inscri, _
            CInt(cboRestricao.Coluna(1).Valor), CInt(cboStatus.Coluna(1).Valor), txtPeriodoInicial, txtPeriodoFinal, txtExercicioInicial, txtExercicioFinal) Then
        Avisa "Nenhum registro encontrado."
        cboImposto.SetFocus
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
    Screen.MousePointer = 11
    Dim i As Double
    For i = 1 To lstObrig.ListItems.Count
        With lstObrig.ListItems
            .Item(i).Selected = True
            ImprimeSelecionado lstObrig, txtRazao, txtEndereco, False, tdiImpressora
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

Private Sub cmdRelatorio_Click()
    Dim CondRelatorio As String
    Dim FORMULA As String
    If Trim(txtIm) = "" And Trim(txtGrupo) = "" Then Exit Sub
    On Error GoTo trata
    Screen.MousePointer = 11
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TObrigLancado.rpt") Then Exit Sub
        If Trim(txtIm) <> "" Then
            CondRelatorio = CondRelatorio & " and {Tab_Conta_Contribuinte.TCC_INSCRICAO} = '" & txtIm & "'"
        Else
            CondRelatorio = CondRelatorio & " and {Tab_Conta_Contribuinte.TCC_INSCRICAO} like '" & txtGrupo & "*'"
        End If
        If cboImposto.ListIndex <> -1 Then
            CondRelatorio = CondRelatorio & " and {Tab_Conta_Contribuinte.tcc_tip_cod_imposto} = '" & cboImposto.Coluna(0).Valor & "'"
        End If
        FORMULA = ""
        If Trim(txtExercicioInicial) <> "" And Trim(txtExercicioFinal) <> "" Then
            CondRelatorio = CondRelatorio & " AND {Tab_Conta_Contribuinte.tcc_periodo} >= " & txtExercicioInicial & " and {Tab_Conta_Contribuinte.tcc_periodo} <= " & txtExercicioFinal & ""
            FORMULA = txtExercicioInicial & " - " & txtExercicioFinal
        End If
        If cboRestricao.ListIndex <> -1 Then
            If cboRestricao.Coluna(1).Valor = 1 Then
                CondRelatorio = CondRelatorio & " and {Tab_Conta_Contribuinte.tcc_status_conta} <> 3"
            ElseIf cboRestricao.Coluna(1).Valor = 2 Then
                CondRelatorio = CondRelatorio & " and {Tab_Conta_Contribuinte.tcc_status_conta} = 3"
            End If
        End If
        .Selecao = Right(CondRelatorio, Len(CondRelatorio) - 4)
        .Formulas "FILTRO", FORMULA
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Arvore = False
        .Visualizar
    
    End With
    Avisa "Impressão concluída."
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    Screen.MousePointer = 0
    Erro Err.Description
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
    cboRestricao.PreencherGeral Bdados, "RESTRICAO DAM"
End Sub

Private Sub lstObrig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lstObrig.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuReimprime.Caption = "Imprimir DAM da obrigação nº " & lstObrig.SelectedItem
            Me.PopupMenu mnuGeral
        End If
    End If
End Sub

Private Sub mnuReimprime_Click()
    If lstObrig.SelectedItem Is Nothing Then Exit Sub
    
    With lstObrig.SelectedItem
        NovaData = Imposto.DataVencimentoNova(.SubItems(5))
        If Trim(NovaData) = "" Then Exit Sub
    End With
    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela
End Sub

Private Sub txtGrupo_LostFocus()
    If Trim(txtGrupo) = "" Then Exit Sub
    txtIm = ""
    txtRazao = ""
    txtEndereco = ""
    txtExercicioInicial.SetFocus
End Sub

Private Sub txtIm_LostFocus()
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco)
End Sub

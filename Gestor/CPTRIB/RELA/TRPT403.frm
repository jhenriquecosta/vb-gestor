VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRPT403 
   BackColor       =   &H80000016&
   Caption         =   "TRPT403"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   6
      Top             =   5985
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   105
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7980
         TabIndex        =   9
         Top             =   105
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   390
         Left            =   5820
         TabIndex        =   7
         Top             =   105
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   688
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9015
         TabIndex        =   10
         Top             =   105
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1138
      Formulario      =   "TREL402"
      Descricao       =   "Relatórios Gerenciais"
      Icone           =   "TRPT403.frx":0000
   End
   Begin VTOcx.grdVISUAL grdRelatorios 
      Height          =   3825
      Left            =   -15
      TabIndex        =   5
      Top             =   2400
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   6747
      CorFundo        =   -2147483633
      Caption         =   "Relação de Descontos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      OcultarRodape   =   -1  'True
      MarcaUnico      =   -1  'True
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
      Height          =   1785
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   585
      Width           =   9885
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   255
         TabIndex        =   1
         Tag             =   "Periodo Inicial"
         Top             =   825
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
         TabIndex        =   2
         Tag             =   "Periodo Final"
         Top             =   825
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
         Left            =   7995
         TabIndex        =   4
         Top             =   825
         Width           =   1500
         _ExtentX        =   2646
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
         Left            =   5685
         TabIndex        =   3
         Tag             =   "Periodo Inicial"
         Top             =   825
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   529
         Caption         =   "Geração"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   450
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboUsuario 
         Height          =   315
         Left            =   825
         TabIndex        =   14
         Tag             =   "Tributo"
         Top             =   1185
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   556
         Caption         =   "Usuário"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   13
         Top             =   1590
         Width           =   45
      End
   End
End
Attribute VB_Name = "TRPT403"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SubSelecao As String
Dim Rpt As New VSRelatorio
Dim Base As Boolean

Private Sub cmdBuscar_Click()

    Dim sql As String
    sql = " SELECT TOC_COD_OBRIGACAO AS Documento,"
    sql = sql & " TOC_INSCRICAO  as Inscrição  ,"
    sql = sql & " TIP_SIGLA_IMPOSTO as Imposto ,"
    sql = sql & " TOC_PERIODO as Período,TOC_DATA_VENCIMENTO as Vencimento,TOC_DATA_GERACAO as Geração,"
    sql = sql & " TOC_VALOR_OBRIGACAO as Valor,TOC_VALOR_MULTA as Multa ,TOC_VALOR_JUROS as Juros  ,"
    sql = sql & " TOC_CORRECAO_MONETARIA as Correção,"
    sql = sql & " TOC_TOTAL_TAXA_INCLUSA as Taxa,"
    sql = sql & " TOC_DESCONTO as [Desconto %],"
    sql = sql & " (TOC_DESCONTO * TOC_VALOR_OBRIGACAO /100) as [Valor Desconto R$],"
    sql = sql & " TOC_USUARIO_DESCONTO as Usuário,"
    sql = sql & " TOC_DATA_DESCONTO  As [Data Desconto]"
    sql = sql & " From TAB_OBRIGACAO_CONTRIBUINTE,TAB_IMPOSTO"
    sql = sql & " Where TOC_DESCONTO > 0 AND TOC_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
    
    SubSelecao = "{TAB_OBRIGACAO_CONTRIBUINTE.TOC_TIP_COD_IMPOSTO} > '1'"
    
    If cboImposto.ListIndex > -1 Then
        sql = sql & " and TOC_TIP_COD_IMPOSTO = " & cboImposto.Coluna(0).Valor
        SubSelecao = SubSelecao & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_TIP_COD_IMPOSTO} = '" & cboImposto.Coluna(0).Valor & "'"
    End If
    If txtExercicioInicial <> "" And txtExercicioFinal <> "" Then
        sql = sql & " and TOC_PERIODO >= '" & txtExercicioInicial & "' AND TOC_PERIODO <= '" & txtExercicioFinal & "'"
        SubSelecao = SubSelecao & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtExercicioInicial & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <= " & txtExercicioFinal
    ElseIf txtExercicioInicial <> "" And txtExercicioFinal = "" Then
        sql = sql & " and TOC_PERIODO >= '" & txtExercicioInicial & "' AND TOC_PERIODO <= '" & txtExercicioInicial & "'"
        SubSelecao = SubSelecao & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} >= " & txtExercicioInicial & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_PERIODO} <= " & txtExercicioInicial
    End If
    
    If txtPeriodoInicial <> "" And txtPeriodoFinal <> "" Then
        SubSelecao = SubSelecao & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_DATA_GERACAO} >= #" & txtPeriodoInicial & "# AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_DATA_GERACAO} <= #" & txtPeriodoFinal & "#"
    ElseIf txtPeriodoInicial <> "" And txtPeriodoFinal = "" Then
        SubSelecao = SubSelecao & " AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_DATA_GERACAO} >= #" & txtPeriodoInicial & "# AND {TAB_OBRIGACAO_CONTRIBUINTE.TOC_DATA_GERACAO} <= #" & txtPeriodoInicial & "#"
    End If
        
    If cboUsuario.ListIndex <> -1 Then
        sql = sql & " AND TOC_USUARIO_DESCONTO LIEK '%" & cboUsuario.Coluna(0).Valor & "%'"
        SubSelecao = SubSelecao & " AND {TAB_TOC_CONTRIBUINTE.TOC_USUARIO_DESCONTO} LIKE '%" & cboUsuario.Coluna(0).Valor & "%'"
    End If
    If Base = True Then Exit Sub
    grdRelatorios.Preencher Bdados, sql
End Sub

Private Sub cmdImprimir_Click()
    
    Base = True
    cmdBuscar_Click
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TListaDescontos.RPT") Then Exit Sub
         .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
         .Selecao = SubSelecao
         .Visualizar
    End With
    Base = False
    
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    cboImposto.SetFocus
End Sub

Private Sub cmdSair_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim Obrig As New OBRIGACAO
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
    
    cboUsuario.Preencher Bdados, "SELECT TUS_COD_USUARIO,TUS_NOME FROM TAB_USUARIO", 1

End Sub


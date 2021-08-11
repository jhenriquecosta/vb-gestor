VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRPT405 
   BackColor       =   &H80000016&
   Caption         =   "TRPT405"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   6195
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6960
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   6
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
   Begin VTOcx.grdVISUAL grdRelatorios 
      Height          =   3735
      Left            =   -15
      TabIndex        =   1
      Top             =   2730
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   6588
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
      Height          =   2130
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   585
      Width           =   9885
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboDoc 
         Height          =   315
         Left            =   270
         TabIndex        =   9
         Tag             =   "Tributo"
         Top             =   345
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   556
         Caption         =   "Tipo Relatório"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1095
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   5670
         TabIndex        =   11
         Top             =   1095
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
         Left            =   9150
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1095
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   675
         TabIndex        =   13
         Top             =   1095
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
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   360
         TabIndex        =   14
         Top             =   1425
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   660
         TabIndex        =   15
         Top             =   1755
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1138
      Icone           =   "TRPT405.frx":0000
   End
End
Attribute VB_Name = "TRPT405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SubSelecao As String
Dim Rpt As New VSRelatorio
Dim Base As Boolean
Private Relatorio As String

Private Sub cmdBuscar_Click()
    Dim Sql As String
    
    If txtIm <> "" And txtImovel <> "" Then
        Util.Avisa "Informe " & txtIm.Caption & " ou " & txtImovel.Caption
        txtIm.SetFocus
        Exit Sub
    End If
    'LEGENDA
    '1 = LANÇAMENTO DE CONTRIBUINTE ADIMPLENTE
    '2 = LANÇAMENTO DE CONTRIBUINTE INADIMPLENTE
    '3 = LANÇAMENTO DE IMÓVEL ADIMPLENTE
    '4 = LANÇAMENTO DE IMÓVEL INADIMPLENTE
    If cboDoc.Coluna(1).Valor = 1 Then
        Sql = "SELECT * FROM VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO WHERE 1 = 1 "
        Relatorio = "TMaioresContribuintesPagos.rpt"
    ElseIf cboDoc.Coluna(1).Valor = 2 Then
        Sql = "SELECT * FROM VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO WHERE 1 = 1 "
        Relatorio = "TMaioresContribuintesDebitos.rpt"
    ElseIf cboDoc.Coluna(1).Valor = 3 Then
        Sql = "SELECT * FROM VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO WHERE 1 = 1 "
        Relatorio = "TMaioresImoveisPagos.rpt"
    ElseIf cboDoc.Coluna(1).Valor = 4 Then
        Sql = "SELECT * FROM VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO WHERE 1 = 1 "
        Relatorio = "TMaioresImoveisAberto.rpt"
    End If
    
    If cboDoc.Coluna(1).Valor = 1 Then
        SubSelecao = " {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO.TRIBUTO} <> '0' "
    ElseIf cboDoc.Coluna(1).Valor = 2 Then
        SubSelecao = " {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO.TRIBUTO} <> '0'"
    ElseIf cboDoc.Coluna(1).Valor = 3 Then
        SubSelecao = " {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO.TRIBUTO} <> '0'"
    ElseIf cboDoc.Coluna(1).Valor = 4 Then
        SubSelecao = " {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO.TRIBUTO} <> '0'"
    End If
        
    
    If cboImposto.ListIndex <> -1 Then
        Sql = Sql & " AND Tributo = '" & cboImposto.Coluna(0).Valor & "'"
        '--------------------------------------------------------------------------'
        If cboDoc.Coluna(1).Valor = 1 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO.TRIBUTO} = '" & cboImposto.Coluna(0).Valor & "'"
        ElseIf cboDoc.Coluna(1).Valor = 2 Then
            SubSelecao = SubSelecao & " AND  {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO.TRIBUTO} = '" & cboImposto.Coluna(0).Valor & "'"
        ElseIf cboDoc.Coluna(1).Valor = 3 Then
            SubSelecao = SubSelecao & " AND  {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO.TRIBUTO} = '" & cboImposto.Coluna(0).Valor & "'"
        ElseIf cboDoc.Coluna(1).Valor = 4 Then
            SubSelecao = SubSelecao & " AND  {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO.TRIBUTO} = '" & cboImposto.Coluna(0).Valor & "'"
        End If
    End If
    If txtIm <> "" Then
        Sql = Sql & " and Contribuinte = '" & txtIm & "'"
        '-----------------------------------------------------------------------------------'
        If cboDoc.Coluna(1).Valor = 1 Then
            SubSelecao = SubSelecao & " AND  {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO.CONTRIBUINTE} = '" & txtIm & "'"
        ElseIf cboDoc.Coluna(1).Valor = 2 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO.CONTRIBUINTE} = '" & txtIm & "'"
        ElseIf cboDoc.Coluna(1).Valor = 3 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO.CONTRIBUINTE} = '" & txtIm & "'"
        ElseIf cboDoc.Coluna(1).Valor = 4 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO.CONTRIBUINTE} = '" & txtIm & "'"
        End If
    ElseIf txtImovel <> "" Then
        Sql = Sql & " and Contribuinte = '" & txtImovel & "'"
        '-------------------------------------------------------------------------------------'
        If cboDoc.Coluna(1).Valor = 1 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_PAGO.CONTRIBUINTE} = '" & txtImovel & "'"
        ElseIf cboDoc.Coluna(1).Valor = 2 Then
            SubSelecao = SubSelecao & " AND  {VIS_LANC_CONTRIB_CONTRIBUINTE_STATUS_ABERTO.CONTRIBUINTE} = '" & txtImovel & "'"
        ElseIf cboDoc.Coluna(1).Valor = 3 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_PAGO.CONTRIBUINTE} = '" & txtImovel & "'"
        ElseIf cboDoc.Coluna(1).Valor = 4 Then
            SubSelecao = SubSelecao & "  AND {VIS_LANC_IMOVEL_CONTRIBUINTE_STATUS_ABERTO.CONTRIBUINTE} = '" & txtImovel & "'"
        End If
    End If
    
    If Base = True Then Exit Sub
    grdRelatorios.Preencher Bdados, Sql
End Sub

Private Sub cmdImprimir_Click()

    Base = True
    cmdBuscar_Click
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\" & Relatorio) Then Exit Sub
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

Private Sub cmdPesquisaInscricao_Click()
AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
 Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
 Dim Obrig As New OBRIGACAO
    cboDoc.PreencherGeral Bdados, "RELATORIO FINANCEIRO"
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
        Ic = txtIm
    End If
    If Trim(txtIm) <> "" Then
        txtIm = BuscaContribuinte(Ic, txtRazao, txtEndereco)
        If Trim(txtIm) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
End Sub

Private Sub txtImovel_LostFocus()
    Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
End Sub


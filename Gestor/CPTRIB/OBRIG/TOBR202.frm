VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TOBR202 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   765
      Left            =   30
      TabIndex        =   30
      Top             =   4740
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1349
      Altura          =   1905
      Caption         =   " Dados da Isencão"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Begin VTOcx.txtVISUAL txtVISUAL1 
         Height          =   300
         Left            =   6210
         TabIndex        =   7
         Top             =   360
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   529
         Caption         =   "Data Requerimento"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   2940
         TabIndex        =   6
         Top             =   360
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   529
         Caption         =   "No. Requerimento"
         Text            =   ""
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   529
         Caption         =   "Protocolo"
         Text            =   ""
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Critérios"
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
      Height          =   1935
      Index           =   3
      Left            =   30
      TabIndex        =   23
      Top             =   690
      Width           =   11565
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   150
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   360
         TabIndex        =   24
         Top             =   810
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   10200
         TabIndex        =   4
         Top             =   1500
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   660
         TabIndex        =   25
         Top             =   1140
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   660
         TabIndex        =   1
         Top             =   495
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3615
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   4125
         TabIndex        =   2
         Top             =   480
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
         Left            =   7635
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtDAM 
         Height          =   300
         Left            =   8070
         TabIndex        =   3
         Top             =   480
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
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   28
         Top             =   1590
         Width           =   45
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1125
      Left            =   45
      TabIndex        =   18
      Top             =   5535
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados da Obrigração"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtTaxas 
         Height          =   300
         Left            =   8790
         TabIndex        =   29
         Top             =   750
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   529
         Caption         =   "Taxas já inclusas"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtVence 
         Height          =   300
         Left            =   9240
         TabIndex        =   11
         Top             =   360
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   529
         Caption         =   "Vencimento"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtTributo 
         Height          =   300
         Left            =   5490
         TabIndex        =   21
         Top             =   330
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   529
         Caption         =   "Tributo"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtJuros 
         Height          =   300
         Left            =   5640
         TabIndex        =   14
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         Caption         =   "Juros"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   300
         Left            =   45
         TabIndex        =   12
         Top             =   750
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   529
         Caption         =   "Valor Obrigação"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtMulta 
         Height          =   300
         Left            =   3600
         TabIndex        =   13
         Top             =   750
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         Caption         =   "Multa"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         Restricao       =   2
         Requerido       =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   300
         Left            =   3420
         TabIndex        =   20
         Top             =   330
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         Caption         =   "Periodo"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtInscricao 
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   330
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   529
         Caption         =   "Contribuinte"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   6720
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1085
      Begin VTOcx.cmdVISUAL cmdObrig 
         Height          =   375
         Left            =   7950
         TabIndex        =   8
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10590
         TabIndex        =   10
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   9420
         TabIndex        =   9
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1138
      Icone           =   "TOBR202.frx":0000
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   15
      Top             =   90
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   2355
      Left            =   30
      TabIndex        =   22
      Top             =   2670
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   4154
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TOBR202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodObrigacao As String

Private Function CriticaCampos() As Boolean
    CriticaCampos = True
    If Not Edita.CriticaCampos(Me) Then
        CriticaCampos = False
        Exit Function
    End If
'    If Len(txtPeriodoInicial) <> Len(txtPeriodoFinal) Then
'        Avisa "Período inconsistente."
'        txtPeriodoInicial.SetFocus
'        CriticaCampos = False
'        Exit Function
'    End If
'    If Len(txtPeriodoInicial) > 4 Then
 '       If Right(Trim(txtPeriodoInicial), 4) <> Right(Trim(txtPeriodoFinal), 4) Then
  '          Avisa "Período deve ser dentro do mesmo ano."
           Rem txtPeriodoInicial.SetFocus
   '         CriticaCampos = False
    '    End If
   ' End If
End Function

Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Set Obrig = New Obrigacao
    If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.coluna(0).Valor), txtIm, , , , , txtPeriodoInicial, _
            txtPeriodoFinal, , txtImovel, etlNaoPagos, , txtDAM) Then
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

Private Sub cmdObrig_Click()
    Dim Obrig As New Obrigacao
    Dim Resultado As Boolean
    Dim Valores As String
    Dim Campos As String
    If Not CriticaCampos Then Exit Sub
    Screen.MousePointer = 11
    If Confirma("Deseja realmente isentar a obrigação?") Then
        If Obrig.Grava_Log_Obrigacao(lstObrig.SelectedItem, Alteracao, AplicacoesVTFuncoes.Usuario, "") = True Then
            If Obrig.AlteraObrigacao(CodObrigacao, txtVence, txtValor, txtMulta, txtJuros, txtTaxas, etsCreditoIsento) Then
                Informa "Registro Gravado com Sucesso."
                cmdBuscar_Click
            Else
                Avisa "Problemas ao Gravar Registro."
            End If
        Else
            Avisa "Erro ao gravar log de obrigação."
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscGrupo, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
End Sub


Private Sub lstObrig_DblClick()
    Dim Obrig As New Obrigacao
    
    If lstObrig.ListItems.Count = 0 Then Exit Sub
    CodObrigacao = lstObrig.SelectedItem
    txtInscricao = lstObrig.SelectedItem.SubItems(1)
    txtTributo = lstObrig.SelectedItem.SubItems(3)
    txtPeriodo = lstObrig.SelectedItem.SubItems(4)
    txtVence = lstObrig.SelectedItem.SubItems(5)
    txtValor = Format(lstObrig.SelectedItem.SubItems(6), Const_Monetario)
    txtMulta = Format(lstObrig.SelectedItem.SubItems(7), Const_Monetario)
    txtJuros = Format(lstObrig.SelectedItem.SubItems(8), Const_Monetario)
    txtTaxas = Format(Nvl(lstObrig.SelectedItem.SubItems(10), 0), Const_Monetario)
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
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
    txtIm = BuscaContribuinte(Ic, txtRazao, txtEndereco)
    
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


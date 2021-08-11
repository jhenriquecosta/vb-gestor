VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   7140
      Width           =   11400
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
      TabIndex        =   26
      Top             =   690
      Width           =   11565
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
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
         Left            =   390
         TabIndex        =   27
         Top             =   810
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   7470
         TabIndex        =   5
         Top             =   1500
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   529
         Caption         =   "Periodo Final"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   270
         TabIndex        =   4
         Top             =   1500
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
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   10200
         TabIndex        =   6
         Top             =   1500
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   690
         TabIndex        =   28
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
         Left            =   690
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
         Left            =   3645
         TabIndex        =   29
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
         Left            =   4155
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
         TabIndex        =   30
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
         TabIndex        =   31
         Top             =   1590
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TOBR301.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1125
      Left            =   45
      TabIndex        =   19
      Top             =   5085
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados da Obrigração"
      CorTexto        =   0
      CorFaixa        =   8421504
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtVence 
         Height          =   300
         Left            =   9240
         TabIndex        =   10
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
         TabIndex        =   22
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
         TabIndex        =   13
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
         Left            =   75
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      TabIndex        =   18
      Top             =   7800
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1085
      Begin VTOcx.cmdVISUAL cmdObrig 
         Height          =   375
         Left            =   7290
         TabIndex        =   7
         Top             =   150
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "&Eliminar Obrigação"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10590
         TabIndex        =   9
         Top             =   150
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
         Left            =   9420
         TabIndex        =   8
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1138
      Icone           =   "TOBR301.frx":2123
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   16
      Top             =   90
      Width           =   375
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   705
      Left            =   60
      TabIndex        =   23
      Top             =   6225
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1244
      Altura          =   1905
      Caption         =   " Dados complementares"
      CorTexto        =   0
      CorFaixa        =   8421504
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6750
         TabIndex        =   15
         Top             =   300
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtTaxas 
         Height          =   300
         Left            =   75
         TabIndex        =   14
         Top             =   330
         Width           =   2370
         _ExtentX        =   4180
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
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   2415
      Left            =   30
      TabIndex        =   24
      Top             =   2670
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   4260
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.Label Label1 
      Caption         =   "Motivo"
      Height          =   255
      Left            =   90
      TabIndex        =   32
      Top             =   6945
      Width           =   840
   End
End
Attribute VB_Name = "TOBR301"
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
    If Len(txtPeriodoInicial) <> Len(txtPeriodoFinal) Then
        Avisa "Período inconsistente."
        txtPeriodoInicial.SetFocus
        CriticaCampos = False
        Exit Function
    End If
'    If Len(txtPeriodoInicial) > 4 Then
 '       If Right(Trim(txtPeriodoInicial), 4) <> Right(Trim(txtPeriodoFinal), 4) Then
  '          Avisa "Período deve ser dentro do mesmo ano."
            txtPeriodoInicial.SetFocus
   '         CriticaCampos = False
    '    End If
   ' End If
End Function

Private Sub cmdBuscar_Click()
    Dim Obrig As Obrigacao
    Set Obrig = New Obrigacao
    If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), txtIm, , , , , txtPeriodoInicial, _
            txtPeriodoFinal, , txtImovel, , , txtDAM) Then
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
    txtMotivo.Tag = "Motivo"
    If Not CriticaCampos Then Exit Sub
    Screen.MousePointer = 11
    If Confirma("Deseja realmente eliminar a obrigação?") Then
        If Obrig.Grava_Log_Obrigacao(lstObrig.SelectedItem, Exclusao, AplicacoesVTFuncoes.Usuario, txtMotivo) = True Then
            If Obrig.EliminaObrigacao(CodObrigacao) Then
                Informa "Registro eliminado."
                cmdBuscar_Click
            Else
                Avisa "Problemas ao gravar registro."
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
    cboStatus.PreencherGeral Bdados, "STATUS OBRIGACAO"
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
    cboStatus.SetarLinha Format(Nvl(lstObrig.SelectedItem.SubItems(13), -1)), 1
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Ic As String
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
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

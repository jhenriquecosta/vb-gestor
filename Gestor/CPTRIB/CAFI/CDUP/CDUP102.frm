VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDUP102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDUP102"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   1138
      Icone           =   "CDUP102.frx":0000
   End
   Begin VTOcx.grdVISUAL grdUtilizacao 
      Height          =   3120
      Left            =   30
      TabIndex        =   7
      Top             =   2910
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5503
      CorBorda        =   32768
      Caption         =   "Cadastros"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1125
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados do Propriet�rio"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   1
         Top             =   375
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   503
         Caption         =   "Nome/Raz�o Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Formato         =   8
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   9
         Top             =   750
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   529
         Caption         =   "Endere�o"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.fraVISUAL fraOcupacao 
      CausesValidation=   0   'False
      Height          =   915
      Left            =   15
      TabIndex        =   11
      Top             =   1890
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1614
      Altura          =   1905
      Caption         =   " Dados da Utiliza��o"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cboVISUAL cboEquipamento 
         Height          =   510
         Left            =   4140
         TabIndex        =   14
         Top             =   330
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   900
         Caption         =   "Equipamento Utilizado"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         Editavel        =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboDestinacao 
         Height          =   510
         Left            =   210
         TabIndex        =   13
         Top             =   330
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   900
         Caption         =   "Destina��o "
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         Editavel        =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtArea 
         Height          =   480
         Left            =   8130
         TabIndex        =   2
         Tag             =   "�rea Ocupada  (M2)"
         Top             =   360
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   847
         Caption         =   "�rea Ocupada  (M2)"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   12
      Top             =   6075
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6495
         TabIndex        =   16
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7500
         TabIndex        =   3
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   9480
         TabIndex        =   5
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8490
         TabIndex        =   4
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.txtVISUAL txtcod 
      Height          =   285
      Left            =   45
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   503
      Caption         =   "Ano"
      Text            =   ""
      Restricao       =   2
      AlinhamentoRotulo=   1
      CorRotulo       =   4210752
      CorTexto        =   4194304
      MaxLen          =   20
      MinLen          =   20
   End
End
Attribute VB_Name = "CDUP102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UtilizacaoSolo  As New eUtilizacaoSolo
Private Sub LimpaCampo()
        cboDestinacao.Text = ""
      cboEquipamento.Text = ""
      txtArea = ""
      cboEquipamento.Preencher Bdados, "select distinct tus_equipamento from tab_cad_utilizacao_subsolo"
     cboDestinacao.Preencher Bdados, "select distinct tus_destinacao from tab_cad_utilizacao_subsolo"
End Sub

Private Sub cmdBuscar_Click()
    UtilizacaoSolo.PreencherGrd grdUtilizacao, txtIm
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdUtilizacao.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
      
    With UtilizacaoSolo
        .Icod = txtcod
        .Im = txtIm
        .Destinacao = cboDestinacao
        .Equipamento = cboEquipamento
        .AreaOcupada = txtArea
        If .Salvar = True Then
            Avisa "Dados Salvos com Sucesso."
            LimpaCampo
            UtilizacaoSolo.PreencherGrd grdUtilizacao, txtIm
        End If
    End With
    Screen.MousePointer = 0
End Sub
    

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
     cboEquipamento.Preencher Bdados, "select distinct tus_equipamento from tab_cad_utilizacao_subsolo"
     cboDestinacao.Preencher Bdados, "select distinct tus_destinacao from tab_cad_utilizacao_subsolo"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub grdUtilizacao_DblClick()
    If grdUtilizacao.ListItems.Count >= 1 Then
        fraOcupacao.Enabled = True
        txtcod = grdUtilizacao.SelectedItem
        cboDestinacao = grdUtilizacao.SelectedItem.SubItems(1)
        cboEquipamento = grdUtilizacao.SelectedItem.SubItems(2)
        txtArea = grdUtilizacao.SelectedItem.SubItems(3)
        txtIm = grdUtilizacao.SelectedItem.SubItems(4)
        txtIm_LostFocus
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
    
End Sub

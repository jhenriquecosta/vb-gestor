VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDOB102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDOB102"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "CDOB102.frx":0000
   End
   Begin VTOcx.grdVISUAL grdObra 
      Height          =   3705
      Left            =   30
      TabIndex        =   11
      Top             =   3240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6535
      CorBorda        =   32768
      Caption         =   "Obras"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1050
      Left            =   30
      TabIndex        =   12
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1852
      Altura          =   1905
      Caption         =   " Dados do Propriet�rio"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   14
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
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   90
         TabIndex        =   0
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
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
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2775
         TabIndex        =   13
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
   End
   Begin VTOcx.fraVISUAL Form1 
      Height          =   1395
      Left            =   30
      TabIndex        =   15
      Top             =   1785
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   2461
      Altura          =   1905
      Caption         =   " Dados da Obra"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtArea 
         Height          =   480
         Left            =   7500
         TabIndex        =   5
         Tag             =   "�rea Atingida  (M2)"
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   847
         Caption         =   "�rea Atingida  (M2)"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.cboVISUAL cboTipoIntervencao 
         Height          =   510
         Left            =   840
         TabIndex        =   2
         Tag             =   "Tipo de Interven��o"
         Top             =   345
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   900
         Caption         =   "Tipo de Interven��o"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtMotivo 
         Height          =   480
         Left            =   825
         TabIndex        =   6
         Tag             =   "Motivo"
         Top             =   900
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   847
         Caption         =   "Motivo"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   100
      End
      Begin VTOcx.txtVISUAL txtPrevisao 
         Height          =   480
         Left            =   5370
         TabIndex        =   4
         Tag             =   "Previs�o"
         Top             =   360
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   847
         Caption         =   "Previs�o (meses)"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   20
      End
      Begin VTOcx.txtVISUAL txtInicioAtividade 
         Height          =   480
         Left            =   3915
         TabIndex        =   3
         Tag             =   "Inicio da Obra"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   847
         Caption         =   "Inicio da Obra"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   10
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   16
      Top             =   7005
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6450
         TabIndex        =   18
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
         Left            =   7440
         TabIndex        =   7
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
         Left            =   9420
         TabIndex        =   9
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
         Left            =   8430
         TabIndex        =   8
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
      Left            =   450
      TabIndex        =   17
      Top             =   825
      Visible         =   0   'False
      Width           =   60
      _ExtentX        =   106
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
Attribute VB_Name = "CDOB102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObraParticular  As New eObraParticular

Private Sub LimpaCampo()
      cboTipoIntervencao.ListIndex = -1
      txtInicioAtividade = ""
      txtPrevisao = ""
      txtArea = ""
      txtMotivo = ""
End Sub



Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdObra.ListItems.Clear
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
      
    With ObraParticular
        .Icod = txtcod
        .Im = txtIm
        .TipoIntervencao = cboTipoIntervencao.Coluna(1).VALOR
        .DataInicio = txtInicioAtividade
        .PREVISAO = txtPrevisao
        .AreaAtingida = txtArea
        .Motivo = txtMotivo
        If .Salvar = True Then
            Avisa "Dados Salvos com Sucesso."
            LimpaCampo
            ObraParticular.PreencherGrd grdObra, txtIm
        End If
    End With
    Screen.MousePointer = 0
End Sub
    

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboTipoIntervencao.PreencherGeral Bdados, "TIPO INTEVENCAO OBRA"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub grdObra_DblClick()
    If grdObra.ListItems.Count >= 1 Then
        txtcod = grdObra.SelectedItem
        cboTipoIntervencao.SetarLinha grdObra.SelectedItem.SubItems(6), 1
        txtInicioAtividade = grdObra.SelectedItem.SubItems(2)
        txtPrevisao = grdObra.SelectedItem.SubItems(3)
        txtArea = grdObra.SelectedItem.SubItems(4)
        txtMotivo = grdObra.SelectedItem.SubItems(5)
        txtIm = grdObra.SelectedItem.SubItems(7)
        txtIm_LostFocus
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)

End Sub
Private Sub cmdBuscar_Click()
    ObraParticular.PreencherGrd grdObra, txtIm
End Sub


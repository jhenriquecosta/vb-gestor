VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDMA102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDMA102"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "CDMA102.frx":0000
   End
   Begin VTOcx.fraVISUAL fraOcupacao 
      Height          =   1200
      Left            =   15
      TabIndex        =   10
      Top             =   1875
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2117
      Altura          =   1905
      Caption         =   " Dados do Equipamento"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtSerie 
         Height          =   285
         Left            =   345
         TabIndex        =   2
         Tag             =   "Série (Nº)"
         Top             =   375
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   503
         Caption         =   "Série (Nº )"
         Text            =   ""
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   50
      End
      Begin VTOcx.txtVISUAL txtLocalizacao 
         Height          =   285
         Left            =   300
         TabIndex        =   5
         Tag             =   "Localização"
         Top             =   810
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   503
         Caption         =   "Localização"
         Text            =   ""
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   100
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   3225
         TabIndex        =   3
         Tag             =   "Tipo"
         Top             =   360
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   6150
         TabIndex        =   4
         Tag             =   "Status"
         Top             =   360
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1125
      Left            =   15
      TabIndex        =   11
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   705
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   13
         Top             =   750
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Tag             =   "Ins. Municipal"
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
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   1
         Top             =   375
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   503
         Caption         =   "Nome/Razão Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
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
   Begin VTOcx.grdVISUAL grdMaquina 
      Height          =   3705
      Left            =   30
      TabIndex        =   14
      Top             =   3150
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6535
      CorBorda        =   32768
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   15
      Top             =   6915
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6450
         TabIndex        =   17
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      TabIndex        =   16
      Top             =   930
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
Attribute VB_Name = "CDMA102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaquinaEquipamento  As New eMaquinaEquipamentoEletromec
Private Sub LimpaCampo()
      txtSerie = ""
      cboStatus.ListIndex = -1
      cboTipo.ListIndex = -1
      txtLocalizacao = ""
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdMaquina.ListItems.Clear
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
      
    With MaquinaEquipamento
        .Icod = txtcod
        .Im = txtIm
        .Serie = txtSerie
        .Status = cboStatus.Coluna(1).VALOR
        .Tipo = cboTipo.Coluna(1).VALOR
        .Localizacao = txtLocalizacao
        If .Salvar = True Then
            Avisa "Dados Salvos com sucesso."
            LimpaCampo
            MaquinaEquipamento.PreencherGrd grdMaquina, txtIm
        End If
    End With
    Screen.MousePointer = 0
End Sub
    

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
      cboStatus.PreencherGeral Bdados, "STATUS CADASTRO MAQUINA"
     cboTipo.PreencherGeral Bdados, "TIPO CADASTRO MAQUINA"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub grdMaquina_DblClick()
    If grdMaquina.ListItems.Count >= 1 Then
      fraOcupacao.Enabled = True
      txtcod = grdMaquina.SelectedItem
      txtSerie = grdMaquina.SelectedItem.SubItems(1)
      cboStatus.SetarLinha grdMaquina.SelectedItem.SubItems(6), 1
      cboTipo.SetarLinha grdMaquina.SelectedItem.SubItems(5), 1
      txtLocalizacao = grdMaquina.SelectedItem.SubItems(4)
      txtIm = grdMaquina.SelectedItem.SubItems(7)
      txtIm_LostFocus
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)

End Sub
Private Sub cmdBuscar_Click()
    MaquinaEquipamento.PreencherGrd grdMaquina, txtIm
End Sub


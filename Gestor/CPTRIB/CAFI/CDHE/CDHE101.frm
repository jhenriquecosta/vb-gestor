VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDHE101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDHE101"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1138
      Icone           =   "CDHE101.frx":0000
   End
   Begin VTOcx.fraVISUAL fraHorario 
      Height          =   930
      Left            =   30
      TabIndex        =   10
      Top             =   1830
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   1640
      Altura          =   1905
      Caption         =   " Horário Especial"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cboVISUAL cboHorarioM 
         Height          =   510
         Left            =   165
         TabIndex        =   2
         Tag             =   "Primeiro Horário"
         ToolTipText     =   "HORARIO ESPECIAL"
         Top             =   345
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   900
         Caption         =   "Primeiro Horário"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cboVISUAL cboHorarioV 
         Height          =   510
         Left            =   1785
         TabIndex        =   3
         Tag             =   "Segundo Horário"
         ToolTipText     =   "HORARIO ESPECIAL"
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   900
         Caption         =   "Segundo Horário"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cboVISUAL cboHorarioN 
         Height          =   510
         Left            =   3510
         TabIndex        =   4
         Tag             =   "Terceiro Horário"
         ToolTipText     =   "HORARIO ESPECIAL"
         Top             =   360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   900
         Caption         =   "Terceiro Horário"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtMotivo 
         Height          =   480
         Left            =   5250
         TabIndex        =   5
         Tag             =   "Motivo"
         Top             =   390
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   847
         Caption         =   "Motivo"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   100
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1050
      Left            =   45
      TabIndex        =   11
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1852
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
         Left            =   90
         TabIndex        =   0
         Tag             =   "Ins. Municipal"
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
         Caption         =   "Nome/Razão Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2775
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
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   14
      Top             =   2850
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   7515
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
         Left            =   9495
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
         Left            =   8505
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
End
Attribute VB_Name = "CDHE101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HorarioEspecial  As New eHorarioEspecial
Private Sub LimpaCampo()
      cboHorarioM.ListIndex = -1
      cboHorarioV.ListIndex = -1
      cboHorarioN.ListIndex = -1
      txtMotivo = ""
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
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
      
    With HorarioEspecial
        .Im = txtIm
        .Horario1 = cboHorarioM.Coluna(1).VALOR
        .Horario2 = cboHorarioV.Coluna(1).VALOR
        .Horario3 = cboHorarioN.Coluna(1).VALOR
        .Motivo = txtMotivo
        If .Salvar = True Then
            Avisa "Dados Salvos com sucesso."
            LimpaCampo
        End If
    End With
    Screen.MousePointer = 0
End Sub
    
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
    
Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboHorarioM.PreencherGeral Bdados, "HORARIO ESPECIAL"
     cboHorarioV.PreencherGeral Bdados, "HORARIO ESPECIAL"
     cboHorarioN.PreencherGeral Bdados, "HORARIO ESPECIAL"
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
   
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub


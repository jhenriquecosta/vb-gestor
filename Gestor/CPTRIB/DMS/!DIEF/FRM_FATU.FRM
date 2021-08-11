VERSION 4.00
Begin VB.Form FRM_FATU 
   BorderStyle     =   0  'None
   Caption         =   "SIAT - Declaração Eletrônica: Informe"
   ClientHeight    =   3195
   ClientLeft      =   630
   ClientTop       =   1635
   ClientWidth     =   8670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   3600
   Left            =   570
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Top             =   1290
   Width           =   8790
   Begin Threed.SSPanel SSPanel1 
      Height          =   3045
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   8505
      _version        =   65536
      _extentx        =   15002
      _extenty        =   5371
      _stockprops     =   15
      borderwidth     =   2
      bevelouter      =   0
      Begin VB.Frame frm_aliquota 
         Enabled         =   0   'False
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1065
         Left            =   30
         TabIndex        =   6
         Top             =   1890
         Width           =   7185
         Begin VB.OptionButton opt_menor_17 
            Caption         =   "Não"
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   3720
            TabIndex        =   8
            Top             =   600
            Width           =   2235
         End
         Begin VB.OptionButton opt_maior_17 
            Caption         =   "Sim"
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2160
            TabIndex        =   7
            Top             =   600
            Width           =   1065
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "FRM_FATU.frx":0000
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   $"FRM_FATU.frx":0442
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   990
            TabIndex        =   9
            Top             =   180
            Width           =   5955
         End
      End
      Begin VB.Frame frame_epp 
         Height          =   1935
         Left            =   30
         TabIndex        =   1
         Top             =   -60
         Width           =   8295
         Begin VB.OptionButton opt_fat2 
            Caption         =   $"FRM_FATU.frx":04B6
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   90
            TabIndex        =   4
            Top             =   960
            Width           =   6765
         End
         Begin VB.OptionButton opt_fat3 
            Caption         =   "Faturamento  superior a R$240.000,00"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   90
            TabIndex        =   3
            Top             =   1380
            Width           =   3825
         End
         Begin VB.OptionButton opt_fat1 
            Caption         =   "Faturamento igual ou inferior a R$160.000,00"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   90
            TabIndex        =   2
            Top             =   600
            Width           =   4245
         End
         Begin VB.Label Label12 
            Caption         =   "Informe o Faturamento Acumulado dentro do ano-calendário."
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   645
            Left            =   150
            TabIndex        =   5
            Top             =   210
            Width           =   7095
         End
      End
      Begin Threed.SSCommand Cancelar 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   7440
         TabIndex        =   11
         Top             =   2520
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   714
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   16711680
         bevelwidth      =   1
         outline         =   0   'False
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   405
         Left            =   7440
         TabIndex        =   10
         Top             =   2040
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   714
         _stockprops     =   78
         caption         =   "&Aceitar"
         forecolor       =   16711680
         bevelwidth      =   1
         outline         =   0   'False
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   3135
      Left            =   60
      Top             =   30
      Width           =   8595
   End
End
Attribute VB_Name = "FRM_FATU"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancelar_Click()
dec002.cmd_iniciar_Click
dec002.Refresh
Unload Me
DoEvents
End Sub


Private Sub Form_Load()
CentrarFormu Me
opt_fat1 = False
opt_fat2 = False
opt_fat3 = False


End Sub

Private Sub opt_fat1_Click()

If opt_fat1.Value = True Then
   'opt_fat2.Value = False
   'opt_fat3.Value = False
   'opt_fat2.Enabled = False
   'opt_fat3.Enabled = False
   frm_aliquota.Enabled = True
   opt_maior_17.Enabled = True
   opt_menor_17.Enabled = True
   Label1.Enabled = True
Else
   'opt_fat2.Value = False
   'opt_fat3.Value = False
   'opt_fat2.Enabled = True
   'opt_fat3.Enabled = True
   frm_aliquota.Enabled = False
   opt_maior_17.Enabled = False
   opt_menor_17.Enabled = False
   Label1.Enabled = False
End If

End Sub

Private Sub opt_fat2_Click()

If opt_fat2.Value = True Then
'   opt_fat1.Value = False
'   opt_fat3.Value = False
'   opt_fat1.Enabled = False
'   opt_fat3.Enabled = False
   frm_aliquota.Enabled = True
   opt_maior_17.Enabled = True
   opt_menor_17.Enabled = True
   Label1.Enabled = True
Else
'   opt_fat1.Value = False
'   opt_fat3.Value = False
'   opt_fat1.Enabled = True
'   opt_fat3.Enabled = True
   frm_aliquota.Enabled = False
   opt_maior_17.Enabled = False
   opt_menor_17.Enabled = False
   Label1.Enabled = False
End If

End Sub


Private Sub opt_fat3_Click()


If opt_fat3.Value = True Then
'   opt_fat1.Value = False
'   opt_fat2.Value = False
'   opt_fat1.Enabled = False
'   opt_fat2.Enabled = False
   frm_aliquota.Enabled = False
   opt_maior_17.Enabled = False
   opt_menor_17.Enabled = False
   Label1.Enabled = False
Else
'   opt_fat1.Value = False
'   opt_fat2.Value = False
'   opt_fat1.Enabled = True
'   opt_fat2.Enabled = True
   frm_aliquota.Enabled = False
   opt_maior_17.Enabled = False
   opt_menor_17.Enabled = False
   Label1.Enabled = False
End If

End Sub


Private Sub SSCommand1_Click()
On Error GoTo erro_fatu
G_opt_fat1 = False
G_opt_fat2 = False
G_opt_fat3 = False

Dim mensa1 As String
Dim mensa2 As String
Dim mensa3 As String
mensa1 = "O Contribuinte excedeu o valor do faturamento anual para usufruir dos benefícios"
mensa2 = "do regime jurídico tributário aplicável à Empresa de Pequeno Porte. A declaração"
mensa3 = "e apuração devem ser efetuadas observando as regras do Regime Normal."

If opt_fat1 = False Then
   If opt_fat2 = False Then
     If opt_fat3 = False Then
        resposta = MsgBox("Não foi marcada opção. Deseja abandonar a declaração?", vbYesNo + vbDefaultButton1, NOME_SISTEMA_DEFAULT)
        If resposta = 6 Then
           'GoTo sair
           Cancelar_Click
        Else
           Exit Sub
        End If
     Else
        G_opt_fat3 = True
        MsgBox mensa1 & Chr(13) & mensa2 & Chr(13) & mensa3, 48, NOME_SISTEMA_DEFAULT
        Cancelar_Click
     End If
  Else
      G_opt_fat2 = True
       If opt_maior_17.Value = True Then
           g_aliquota_Maior17 = 1   'sim
           g_condicao = 4
        ElseIf opt_menor_17.Value = True Then
           g_aliquota_Maior17 = 2   'nao
           g_condicao = 6
        Else
            resposta = MsgBox("Não foi marcada opção de Alíquota. Deseja abandonar a declaração?", vbYesNo + vbDefaultButton1, NOME_SISTEMA_DEFAULT)
            If resposta = 6 Then
               Cancelar_Click
               'GoTo sair
            Else
               Exit Sub
            End If
        End If
     
   End If
Else
  G_opt_fat1 = True
  If opt_maior_17.Value = True Then
     g_aliquota_Maior17 = 1   'sim
     g_condicao = 3
  ElseIf opt_menor_17.Value = True Then
     g_aliquota_Maior17 = 2   'nao
     g_condicao = 5
  Else
      resposta = MsgBox("Não foi marcada opção de Alíquota. Deseja abandonar a declaração?", vbYesNo + vbDefaultButton1, NOME_SISTEMA_DEFAULT)
      If resposta = 6 Then
         Cancelar_Click
         'GoTo sair
      Else
         Exit Sub
      End If
  End If
End If


sair:
Unload Me
erro_fatu:
 If Err <> 0 Then
    MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
    Resume Next
 End If
End Sub



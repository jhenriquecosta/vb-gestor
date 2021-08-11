VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TVAL206 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCOB202"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TVAL206.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1170
      Left            =   90
      TabIndex        =   10
      Top             =   735
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   2064
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483637
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   330
         Left            =   285
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   750
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   582
         Caption         =   "Contribuinte"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtControle 
         Height          =   330
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   582
         Caption         =   "Controle"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtAno 
         Height          =   330
         Left            =   2865
         TabIndex        =   1
         Top             =   375
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         Caption         =   "Ano"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   330
         Left            =   90
         TabIndex        =   0
         Top             =   375
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   582
         Caption         =   "Insc. Municipal"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         AutoTAB         =   -1  'True
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   9
      Top             =   1980
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdValidar 
         Height          =   375
         Left            =   5070
         TabIndex        =   3
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Validar"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6225
         TabIndex        =   4
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   1138
      Icone           =   "TVAL206.frx":2123
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   2
      Left            =   2340
      TabIndex        =   6
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmd 
      Height          =   375
      Index           =   0
      Left            =   2340
      TabIndex        =   7
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdVISUAL1 
      Height          =   375
      Left            =   1170
      TabIndex        =   8
      Top             =   0
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TVAL206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdValidar_Click()
    Dim Controle As Double
    On Error Resume Next
    Controle = 2 * CDbl(Right(txtIm, 2) + Left(txtIm, 8))
    Controle = Controle + 3 * (CDbl(Left(txtIm, 8) + Right(txtIm, 2)) + CDbl(Nvl(txtAno, 0)))
    
    If Controle = Trim(txtControle) Then
        Informa "Número de Controle OK!"
    Else
        Erro "Número de Controle inválido!"
    End If
    txtControle.SetFocus
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub txtIm_LostFocus()
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
    txtIm = BuscaContribuinte(Ic, txtRazao)
    
End Sub

VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAID201A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAID201A.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   5865
      Left            =   45
      TabIndex        =   8
      Top             =   690
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   10345
      Caption         =   "Consulta"
      Descricao       =   "Consulta contribuintes, trazendo informações gerais"
      corFaixa        =   32768
      Icone           =   "TAID201A.frx":2123
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   6855
         TabIndex        =   1
         Top             =   720
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         Caption         =   "&Buscar"
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtNome 
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   750
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   503
         Caption         =   "Nome"
         Text            =   ""
      End
      Begin VTOcx.grdVISUAL grdPesquisa 
         Height          =   4695
         Left            =   75
         TabIndex        =   2
         Top             =   1110
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   8281
         CorBorda        =   32768
         CorTitulo       =   32768
         CorCaption      =   16777215
         CorDica         =   32768
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   7
      Top             =   6585
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   1032
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdOK 
         Height          =   375
         Left            =   5520
         TabIndex        =   3
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&OK"
         Acao            =   8
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6735
         TabIndex        =   4
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   1138
      Icone           =   "TAID201A.frx":243D
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   135
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "TAID201A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PessoaFisica As Boolean
'Public FuncaoAidf As AIDF
Dim Grafica As cGraficaAidf
Dim Contribuinte As cContribuinte
Dim Formulario As Form

Public Sub Inicia(FormChamador As Form)
    Set Formulario = FormChamador
End Sub

Private Sub cmdBuscar_Click()
    If Trim(txtNome.Text) = "" Then
        Util.Avisa "Informe um nome para a pesquisa"
        txtNome.SetFocus
        Exit Sub
    End If
    If Me.Tag = "TAID201GR" Then
        If Grafica.PreencherGrid(grdPesquisa, , txtNome) = False Then
            Util.Avisa "Nenhuma gráfica encontrada."
            Exit Sub
        End If
        Exit Sub
    End If
    If Me.Tag = "TAID201IM" Then
        If Contribuinte.PreencherGrid(grdPesquisa, txtNome) = False Then
            Util.Avisa "Nenhum contribuinte encontrado."
            Exit Sub
        End If
    End If
    grdPesquisa.SetFocus
End Sub


Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdOK_Click()
    grdPesquisa_DblClick
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision

    Screen.MousePointer = 0
'    Set FuncaoAidf = New AIDF
    Set Grafica = New cGraficaAidf
    Set Contribuinte = New cContribuinte
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set FuncaoAidf = Nothing
    Set Grafica = Nothing
    Set Contribuinte = Nothing
End Sub

Private Sub grdPesquisa_DblClick()
    Dim Controle As Control
    If grdPesquisa.SelectedItem Is Nothing Then Exit Sub
    Select Case Me.Tag
        Case "TAID201IM"
            Formulario.txtIm = grdPesquisa.SelectedItem
            'Unload Me
            Formulario.txtIm.Enabled = True
            Set Controle = Formulario.txtIm
        Case "TAID101IM"
            Formulario.txtIm = grdPesquisa.SelectedItem.SubItems(1)
            'Unload Me
            Formulario.txtIm.Enabled = True
            Set Controle = Formulario.txtIm
        Case "TAID201GR"
            Formulario.txtImGrafica.Text = Edita.TiraTudo(grdPesquisa.SelectedItem.SubItems(1))
            Set Controle = Formulario.txtImGrafica
            DoEvents
            'Unload Me
    End Select
    Unload Me
    If Not Controle Is Nothing Then
        Controle.SetFocus
        SendKeys "{TAB}"
    End If
End Sub

'Private Sub Timer_Timer()
'    On Error Resume Next
'End Sub

'Private Sub txtNome_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'
'End Sub

Private Sub txtNome_Change()

End Sub

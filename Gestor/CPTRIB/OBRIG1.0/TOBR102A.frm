VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TOBR102.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   5430
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1085
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6960
         TabIndex        =   15
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdObrig 
         Height          =   375
         Left            =   8160
         TabIndex        =   5
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10500
         TabIndex        =   7
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
         Left            =   9345
         TabIndex        =   6
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
      Height          =   1245
      Index           =   3
      Left            =   30
      TabIndex        =   11
      Top             =   690
      Width           =   11535
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   8910
         TabIndex        =   3
         Top             =   150
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         Caption         =   ""
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   5760
         TabIndex        =   2
         Tag             =   "Periodo Inicial"
         Top             =   150
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   529
         Caption         =   "Periodo(dd/mm/aaaa)"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   510
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   90
         TabIndex        =   14
         Top             =   870
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   10140
         TabIndex        =   4
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   90
         TabIndex        =   0
         Tag             =   "Inscrição"
         Top             =   150
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   529
         Caption         =   "Inscrição"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   150
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   12
         Top             =   1710
         Width           =   45
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1138
      Icone           =   "TOBR102.frx":2123
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   90
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   3420
      Left            =   0
      TabIndex        =   16
      Top             =   1980
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6033
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
End
Attribute VB_Name = "TOBR102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obrig As New Obrigacao
Private Function CriticaCampos() As Boolean
    CriticaCampos = True
    If Not Edita.CriticaCampos(Me) Then
        CriticaCampos = False
        Exit Function
    End If
    If Len(txtPeriodoInicial) > Len(txtPeriodoFinal) Then
        Avisa "Período inconsistente."
        txtPeriodoInicial.SetFocus
        CriticaCampos = False
        Exit Function
    End If
    If Len(txtPeriodoInicial) > 4 Then
        If Right(Trim(txtPeriodoInicial), 4) <> Right(Trim(txtPeriodoFinal), 4) Then
            Avisa "Período deve ser dentro do mesmo ano."
            txtPeriodoInicial.SetFocus
            CriticaCampos = False
        End If
    End If
End Function

Private Sub cmdBuscar_Click()
    If Not Obrig.CarregaPeriodosObrigacao(lstObrig, txtIm, _
        txtPeriodoInicial, txtPeriodoFinal) Then
        Avisa "Nenhum registro encontrado."
    End If
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstObrig.ListItems.Clear
    txtIm.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdExcluir_Click()
    If Confirma("Confirma exclusão do período " & txtPeriodoInicial & " a " & _
        txtPeriodoFinal & " para a inscrição nº " & txtIm & "?") Then
        If Obrig.EliminaPeriodoObrigacao(txtIm, txtPeriodoInicial) Then
            Obrig.CarregaPeriodosObrigacao lstObrig, txtIm
            Avisa "Registro eliminado."
        Else
            Avisa "Não foi possível eliminar registro."
        End If
    End If
End Sub

Private Sub cmdObrig_Click()
    
    Dim Resultado As Boolean
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    lstObrig.Preencher Bdados, ""
    
    If Obrig.GravaPeriodoObrigacao(txtIm, txtPeriodoInicial, txtPeriodoFinal) Then
        Obrig.CarregaPeriodosObrigacao lstObrig, txtIm
        Informa "Dados gravados com sucesso."
    Else
        Informa "Problemas ao gravar dados."
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscGrupo, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub Frame1_DblClick(Index As Integer)
'    Dim Obrig As New Obrigacao
'    Obrig.GBDam
'    Obrig.GBDARM
End Sub

Private Sub lstObrig_Click()
    If lstObrig.ListItems.Count = 0 Then Exit Sub
    txtIm = lstObrig.SelectedItem
    txtIm_LostFocus
    txtPeriodoInicial = lstObrig.SelectedItem.SubItems(1)
    txtPeriodoFinal = lstObrig.SelectedItem.SubItems(2)
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
    BuscaContribuinte Ic, txtRazao, txtEndereco
End Sub

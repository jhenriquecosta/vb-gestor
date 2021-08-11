VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCOB402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB402.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1155
      Left            =   60
      TabIndex        =   13
      Top             =   750
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   2037
      Altura          =   1905
      Caption         =   " Opções de Busca"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   720
         TabIndex        =   0
         Top             =   360
         Width           =   2610
         _ExtentX        =   4604
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
         Left            =   3360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   375
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   7455
         TabIndex        =   15
         Top             =   420
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Período"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   14
         Top             =   795
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   397
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Contribuinte"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         FloodColor      =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtContrib 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1500
         MaxLength       =   80
         TabIndex        =   3
         Tag             =   "DESCRICAOATIVIDADE"
         Top             =   750
         Width           =   7845
      End
      Begin VB.TextBox txtPeriodo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8145
         MaxLength       =   4
         TabIndex        =   2
         Top             =   375
         Width           =   1200
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   12
      Top             =   5100
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdImprimeAtividade 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   90
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   661
         Caption         =   "Imprimir Por atividade"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimeRelacao 
         Height          =   375
         Left            =   4770
         TabIndex        =   6
         Top             =   90
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Imprimir Relação"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7950
         TabIndex        =   8
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBusca 
         Height          =   375
         Left            =   6780
         TabIndex        =   5
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9135
         TabIndex        =   9
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   5145
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VTOcx.grdVISUAL lstAtv 
      Height          =   3105
      Left            =   60
      TabIndex        =   4
      Top             =   1980
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5477
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   1138
      Icone           =   "TCOB402.frx":2123
   End
   Begin VB.Menu mnuGeral 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnureimprime 
         Caption         =   "&Reimprimir ALVARA"
      End
   End
End
Attribute VB_Name = "TCOB402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Criterio As String

Private Sub cmdBusca_Click()
    Dim Sql As String
    Criterio = ""
    Screen.MousePointer = 11
    Sql = "Select tci_im as Im,tci_nome as Razão_Social, tai_cod_sequencial as Protocolo, "
    Sql = Sql & " tai_periodo as Periodo, tai_data_impressao as Data_Impresso,"
    Sql = Sql & " TAI_TUS_COD_USUARIO AS Usuario "
    Sql = Sql & " from TAB_ALVARA_IMPRESSO, "
    Sql = Sql & " tab_contribuinte  where tai_tci_im = tci_im"
    
    If Trim(txtIm) <> "" Then
        Sql = Sql & " and tci_im ='" & txtIm & "'"
        Criterio = "{Tab_Contribuinte.tci_im} = '" & txtIm & "'"
    End If
    If Trim(txtPeriodo) <> "" Then
        Sql = Sql & " and tai_periodo =" & txtPeriodo
        Criterio = Criterio & IIf(Criterio <> "", " AND ", "") & "{TAB_ALVARA_IMPRESSO.tai_periodo}=" & txtPeriodo
    End If
    lstAtv.Preencher Bdados, Sql, 1200, 4800, 850, 1300, 1500
    If lstAtv.ListItems.Count = 0 Then
        Util.Informa "Nenhum registro encontrado."
    End If
    
'    lstAtv.Mensagem = "Total lançado" & Format(lstAtv.Colunas(5).Soma, Const_Monetario)
    Screen.MousePointer = 0

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprimeAtividade_Click()
    Screen.MousePointer = 11
    With RPT
        If Not .DefinirArquivo(Bdados, App.Path + "\ALVARA_IMPRESSO_ATIVIDADE.rpt") Then Exit Sub
        If Trim(txtPeriodo) <> "" Then
            .Formulas "VT_PERIODO", "RELAÇÃO DE ALVARAS IMPRESSOS EM " & txtPeriodo
            .Selecao = "{VIS_ALVARA_IMPRESSO.TAI_PERIODO} = " & txtPeriodo
        Else
            .Formulas "VT_PERIODO", "RELAÇÃO DE ALVARAS IMPRESSOS"
        End If
        .Visualizar
    End With
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdImprimeRelacao_Click()
    Screen.MousePointer = 11
    With RPT
        If Not .DefinirArquivo(Bdados, App.Path + "\ALVARA_IMPRESSO_RELACAO.rpt") Then Exit Sub
        If Trim(txtPeriodo) <> "" Then
            .Formulas "VT_PERIODO", "RELAÇÃO DE ALVARAS IMPRESSOS EM " & txtPeriodo
            .Selecao = "{VIS_ALVARA_IMPRESSO.TAI_PERIODO} = " & txtPeriodo
        Else
            .Formulas "VT_PERIODO", "RELAÇÃO DE ALVARAS IMPRESSOS"
        End If
        .Visualizar
    End With
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    lstAtv.Preencher Bdados, ""
    txtIm.SetFocus
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
'    Dim Sql As String
'
'    Sql = "Select tai_tci_im as Im, tci_nome as Razao,tai_periodo as Periodo, tai_data_impressao as [Data Impresso] " & _
'        ",TAI_TUS_COD_USUARIO AS Usuario from TAB_ALVARA_IMPRESSO, tab_contribuinte  where tai_tci_im = tci_im"
'    lstAtv.Preencher Bdados, Sql, 1400
'    AtualizaCabecalho lstAtv
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
    
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtDescAtiv_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtMult_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub lstAtv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 2 And lstAtv.ListItems.Count > 0 Then
'        mnuReimprime.Caption = "Reimprimir ALVARA " & lstAtv.SelectedItem.SubItems(2) & " de " & lstAtv.SelectedItem.SubItems(1)
'        Me.PopupMenu mnuGeral
'    End If
End Sub

Private Sub txtIm_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    If Trim(txtIm) = "" Then
        txtContrib = ""
    Else
        If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
            txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
        End If
        Sql = "Select tci_nome from tab_contribuinte where tci_im ='" & txtIm & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtContrib = rs(0)
        Else
            Avisa "Contribuinte não encontrado."
            txtIm.SetFocus
        End If
    End If
    Bdados.FechaTabela rs
End Sub

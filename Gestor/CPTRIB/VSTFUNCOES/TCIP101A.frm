VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIP101A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BCP-Consultoria e Tecnologia em Administração Pública"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ControlBox      =   0   'False
   Icon            =   "TCIP101A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   6660
      TabIndex        =   5
      Top             =   6330
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSFrame fra 
      Height          =   885
      Left            =   45
      TabIndex        =   7
      Top             =   690
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   1561
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtNome 
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
         Left            =   975
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "Nome"
         Top             =   105
         Width           =   6480
      End
      Begin VB.ComboBox cboPesquisa 
         DataField       =   "ttl_nome"
         DataSource      =   "dtTipLogr"
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
         Height          =   330
         ItemData        =   "TCIP101A.frx":030A
         Left            =   960
         List            =   "TCIP101A.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "TIPOCONTRIBUINTE"
         Top             =   480
         Width           =   2025
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   180
         Width           =   675
         _ExtentX        =   1191
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
         Caption         =   "Nome"
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   0
         FloodColor      =   12632256
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   540
         Width           =   780
         _ExtentX        =   1376
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
         Caption         =   "Pesquisa"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6510
         TabIndex        =   2
         Top             =   480
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   7140
      TabIndex        =   6
      Top             =   1020
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   1138
      Icone           =   "TCIP101A.frx":0338
   End
   Begin VTOcx.grdVISUAL lstPesq 
      Height          =   4635
      Left            =   30
      TabIndex        =   3
      Top             =   1635
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   8176
   End
   Begin VTOcx.cmdVISUAL cmdOK 
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   6330
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      Caption         =   "&OK"
      Acao            =   8
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIP101A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Public PessoaFisica As Boolean
Dim Controle As Control
Dim Contribuinte As Control

Public Sub Inicia(TipoPessoa As TipoContrib, Optional ControleDestino As Control, Optional Nome As Control)
    cboPesquisa.ListIndex = 0
    Set Controle = ControleDestino
    Set Contribuinte = Nome
End Sub

Private Sub cmdBuscar_Click()
    Dim Rs As VSRecordset
    Dim Sql As String
    
    If Trim(txtNome.Text) = "" Then
        Util.Avisa "Informe um nome para a pesquisa"
        txtNome.SetFocus
        Exit Sub
    End If
    
    'If cboPesquisa.ListIndex > -1 Then
        Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where "
        Select Case cboPesquisa.ListIndex
            Case 0 'INICIO
                Sql = Sql & "(tci_nome like '" & txtNome & "%')"
            Case 1 'MEIO
                Sql = Sql & "(tci_nome like '%" & txtNome & "%')"
            Case 2 'FIM
                Sql = Sql & "(tci_nome like '%" & txtNome & "')"
            Case 3 'EXATO
                Sql = Sql & "(tci_nome = '" & txtNome & "')"
        End Select
        Sql = Sql & " and (tci_tsc_cod_sit_cad =1) and (not tci_tae_cae is null) and (tci_tae_cae > 0)"
        Sql = Sql & " order by tci_nome"
        
        lstPesq.Preencher Bdados, Sql, 1200, 3600, 1900
        
        If lstPesq.ListItems.Count = 0 Then
            Call Util.Avisa("Nenhum contribuinte encontrado.")
            txtNome.SetFocus
        Else
            lstPesq.SetFocus
        End If
    'Else
    '    Call Util.Avisa("Informe o tipo de contribuinte do imposto.")
    '    cboTipoContrib.SetFocus
    'End If
End Sub

Private Sub cmdEnter_Click()
    If Me.ActiveControl.Name = "txtNome" Then
        cmdBuscar_Click
    ElseIf Me.ActiveControl.Name = "lstPesq" Then
        lstPesq_DblClick
    Else
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdOK_Click()
    lstPesq_DblClick
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Screen.MousePointer = 0
    
    
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    If lstPesq.ListItems.Count = 0 Then Exit Sub
    
    Me.Tag = lstPesq.SelectedItem.Text
    If Not Controle Is Nothing Then
        Controle = lstPesq.SelectedItem.Text
        If Not Contribuinte Is Nothing Then
            Contribuinte = lstPesq.SelectedItem.SubItems(1)
        End If
    End If
    Unload Me
    If Controle.Enabled = False Then
        Controle.Enabled = True
    End If
    Controle.SetFocus
    SendKeys "{TAB}"
    DoEvents
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
    
End Sub

Private Sub txtNome_GotFocus()
    lstPesq.ListItems.Clear
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


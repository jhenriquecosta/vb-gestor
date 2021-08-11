VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{EA761AE1-8FDE-4340-8E6D-420E99B0C363}#1.0#0"; "VTControles.ocx"
Begin VB.Form TCIP101A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   6660
      TabIndex        =   3
      Top             =   6330
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   5670
      TabIndex        =   2
      Top             =   6330
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin Threed.SSFrame fra 
      Height          =   885
      Left            =   45
      TabIndex        =   5
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
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   465
         Width           =   5610
      End
      Begin VB.ComboBox cboTipoContrib 
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
         ItemData        =   "TCIP101A.frx":0000
         Left            =   1815
         List            =   "TCIP101A.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "TIPOCONTRIBUINTE"
         Top             =   90
         Width           =   3585
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   1260
         TabIndex        =   6
         Top             =   510
         Width           =   495
         _ExtentX        =   873
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
         AutoSize        =   1
         Alignment       =   4
         FloodColor      =   12632256
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   7
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
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
         Caption         =   "Tipo Contribuinte"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   7140
      TabIndex        =   4
      Top             =   1020
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   1138
      Icone           =   "TCIP101A.frx":002E
   End
   Begin VTOcx.grdVISUAL lstPesq 
      Height          =   4635
      Left            =   60
      TabIndex        =   8
      Top             =   1635
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4339
   End
End
Attribute VB_Name = "TCIP101A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PessoaFisica As Boolean

Private Sub cboTipoContrib_Click()
    If cboTipoContrib.ListIndex > -1 Then
        txtNome.Enabled = True
        If cboTipoContrib.ListIndex = 0 Then
            PessoaFisica = True
            lbl(0).Caption = "Nome:"
        Else
            PessoaFisica = False
            lbl(0).Caption = "Razão Social:"
        End If
    Else
        txtNome.Enabled = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim Rs As VSRecordset
    Dim Sql As String
    
    If cboTipoContrib.ListIndex > -1 Then
        Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where (tci_nome like '" & txtNome & "%' or tci_nome like '% " & txtNome & "%'"
        Sql = Sql & ") and (tci_tsc_cod_sit_cad =1)"
'        If Not PessoaFisica Then
'            Sql = Sql & " and tci_tipo_contribuinte = 2)"
'        Else
'            Sql = Sql & " and tci_tipo_contribuinte = 1)"
'        End If
        If Not Bdados.AbreTabela(Sql, Rs) Then
            Call Util.Avisa("Nenhum contribuinte encontrado.")
        End If
        Bdados.FechaTabela Rs
        lstPesq.Preencher Bdados, Sql, 1400
    Else
        Call Util.Avisa("Informe o tipo de contribuinte do imposto.")
        cboTipoContrib.SetFocus
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
    If Me.ActiveControl.Name = "txtNome" Then
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
    AtualizaCabecalho lstPesq
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Screen.MousePointer = 0
    
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    Select Case TCIP101A.Tag
        Case "TCIP101"
            TCIP101.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIP101.txtIM.Enabled = True
            TCIP101.txtIM.SetFocus
            TCIP101.txtIM.Enabled = True
        Case "TCIP102"
            TCIP102.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIP102.txtIM.Enabled = True
            TCIP102.txtIM.SetFocus
            TCIP102.txtIM.Enabled = True
        Case "TCIP201"
            TCIP201.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIP201.txtIM.Enabled = True
            TCIP201.txtIM.SetFocus
        Case "TCIS102"
            TCIS102.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIS102.txtIM.Enabled = True
            TCIS102.txtIM.SetFocus
        Case "TCIS201"
            TCIS201.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIS201.txtIM.Enabled = True
            TCIS201.txtIM.SetFocus
        Case "TAID201"
            TAID201.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TAID201.txtIM.SetFocus
        Case "TAID201G"
            TAID201.txtImGrafica = lstPesq.SelectedItem.Text
            Unload Me
            TAID201.txtImGrafica.SetFocus
        Case "TAID101"
            TAID101.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TAID101.txtIM.SetFocus
        Case "TCIU101"
            TCIU101.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIU101.txtIM.Enabled = True
            TCIU101.txtIM.SetFocus
        Case "TCIU101B"
            TCIU101.txtIMBc = lstPesq.SelectedItem.Text
            Unload Me
            TCIU101.txtIMBc.Enabled = True
            TCIU101.txtIMBc.SetFocus
        Case "TCIU201"
            TCIU201.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCIU201.txtIM.Enabled = True
            TCIU201.txtIM.SetFocus
        Case "TCOB102"
            TCOB102.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCOB102.txtIM.SetFocus
        Case "TCOB202"
            TCOB202.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCOB202.txtIM.SetFocus
        Case "TRET101"
            TRET101.txtimRem = lstPesq.SelectedItem.Text
            Unload Me
            TRET101.txtimRem.SetFocus
        Case "TSUB101"
            TSUB101.txtImToma = lstPesq.SelectedItem.Text
            Unload Me
            TSUB101.txtImToma.SetFocus
        Case "TSUB101B"
            TSUB101.txtimRem = lstPesq.SelectedItem.Text
            Unload Me
            TSUB101.txtimRem.SetFocus
          Case "TCIS101"
            TCIS101.txtImRepresentante = lstPesq.SelectedItem.Text
            Unload Me
            TCIS101.txtImRepresentante.Enabled = True
            TCIS101.txtImRepresentante.SetFocus
        Case "TCTA101"
            TCTA101.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCTA101.txtIM.SetFocus
        Case "TCOB401"
            TCOB401.txtIM = lstPesq.SelectedItem.Text
            Unload Me
            TCOB401.txtIM.SetFocus
    End Select
    SendKeys "{TAB}"
    DoEvents
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


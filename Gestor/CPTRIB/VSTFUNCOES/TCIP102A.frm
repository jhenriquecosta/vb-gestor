VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIP102A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BCP-Consultoria e Tecnologia em Administração Pública"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   885
      Left            =   45
      TabIndex        =   7
      Top             =   720
      Width           =   9330
      _ExtentX        =   16457
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
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   465
         Width           =   5010
      End
      Begin VB.ComboBox cboTipoInscricao 
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
         ItemData        =   "TCIP102A.frx":0000
         Left            =   1305
         List            =   "TCIP102A.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "TIPOINSCRICAO"
         Top             =   90
         Width           =   2565
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   510
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "Nome/Razão"
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   4
         FloodColor      =   12632256
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   9
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
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
         Caption         =   "Tipo Inscricao"
         BorderWidth     =   1
         BevelOuter      =   0
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   420
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8280
         TabIndex        =   5
         Top             =   420
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdOK 
         Height          =   375
         Left            =   7320
         TabIndex        =   4
         Top             =   420
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&OK"
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
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
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   1138
      Formulario      =   "Pesquisa Inscrição Municipal/Cadastral"
      Icone           =   "TCIP102A.frx":0031
   End
   Begin VTOcx.grdVISUAL lstPesq 
      Height          =   4995
      Left            =   45
      TabIndex        =   3
      Top             =   1680
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   8811
      CorBorda        =   16711680
      CorTitulo       =   16711680
      CorCaption      =   16777215
   End
End
Attribute VB_Name = "TCIP102A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnInscricaoMunicipal As Boolean
Dim Controle As Control
Dim Contribuinte As Control
Dim tpInscricao As TipoInsc


Public Sub Inicia(tipoInscricao As TipoInsc, Optional ControleDestino As Control, Optional Nome As Control)
    'cboTipoInscricao.ListIndex = TipoInscricao
    tpInscricao = tipoInscricao
    
    If tipoInscricao = InscContrib Then
        cboTipoInscricao.ListIndex = 0
        cboTipoInscricao.Enabled = False
    ElseIf tipoInscricao = InscImovel Then
        cboTipoInscricao.ListIndex = 1
        cboTipoInscricao.Enabled = False
        
    ElseIf tipoInscricao = InscCpfCnpj Then
        cboTipoInscricao.ListIndex = 2
        cboTipoInscricao.Enabled = True
    
    Else
        cboTipoInscricao.ListIndex = 0
    End If
    Set Controle = ControleDestino
    Set Contribuinte = Nome
    
End Sub

Private Sub cboTipoContrib_Click()
    
End Sub

Private Sub cboTipoInscricao_Click()
    If cboTipoInscricao.ListIndex > -1 Then
        txtNome.Enabled = True
        If cboTipoInscricao.ListIndex = 0 Then
            blnInscricaoMunicipal = True
            lbl(0).Caption = "Nome/Razão:"
        ElseIf cboTipoInscricao.ListIndex = 1 Then
            blnInscricaoMunicipal = False
            lbl(0).Caption = "Nome/Razão:"
        Else
            blnInscricaoMunicipal = True
            lbl(0).Caption = "Cpf/Cnpj:"
        End If
    Else
        txtNome.Enabled = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo TrataErro
    Dim Rs As VSRecordset
    Dim Sql As String
    
    If Not cboTipoInscricao.ListIndex > -1 Then Exit Sub
    
    If Trim(txtNome.Text) = "" Then
        Util.Avisa "Informe um nome para a pesquisa"
        txtNome.SetFocus
        Exit Sub
    End If
    
    Dim strCaption As String
    strCaption = lstPesq.Caption
    
    Screen.MousePointer = 11
    
    lstPesq.ListItems.Clear
    
    lstPesq.Caption = "Carregando..."
    DoEvents
    
    If blnInscricaoMunicipal Then
        Sql = "Select tci_im as IM, tci_nome as Razao,"
        Sql = Sql & " tci_cgc_cpf as CPF_CGC "
        Sql = Sql & " from Tab_Contribuinte where (tci_nome like '%" & txtNome & "%'"
        Sql = Sql & ") and (tci_tsc_cod_sit_cad =1)"
        Sql = Sql & " order by tci_nome"
        If tpInscricao = InscCpfCnpj And cboTipoInscricao.ListIndex = 2 Then
        
            Sql = "Select tci_im as IM, tci_nome as Razao,"
            Sql = Sql & " tci_cgc_cpf as CPF_CGC "
            Sql = Sql & " from Tab_Contribuinte where (tci_cgc_cpf like '%" & txtNome & "%'"
            Sql = Sql & ") and (tci_tsc_cod_sit_cad =1)"
            Sql = Sql & " order by tci_nome"
        
        End If
    Else
        Sql = "SELECT tim_ic AS Inscricao, TCI_NOME AS PROPRIETARIO, " & Bdados.Concatena & " ' ' " & Bdados.Concatena & " tci_logradouro " & Bdados.Concatena & " ' ' " & Bdados.Concatena & " TCI_NOME_LOGRADOURO AS Endereco " _
        & " " _
        & " From VIS_IMOVEL "
        
        Sql = Sql & " where tci_nome like '%" & txtNome & "%'"
        Sql = Sql & " order by 2"
    End If
    
    
    lstPesq.Preencher Bdados, Sql, 1650, 4000
    lstPesq.Caption = strCaption
    Screen.MousePointer = 0
    
    If lstPesq.ListItems.Count = 0 Then
        Util.Avisa "Nenhum registro encontrado"
        txtNome.SetFocus
    Else
        'lstPesq.SetFocus
    End If
    
    Exit Sub
TrataErro:
    Util.Erro (Err.Description)

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

Private Sub Form_Activate()
    If cboTipoInscricao.ListIndex = -1 Then
        cboTipoInscricao.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Screen.MousePointer = 0
    
End Sub

Private Sub lstPesq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstPesq, ColumnHeader
End Sub

Private Sub lstPesq_DblClick()
    On Error GoTo TrataErro
    Dim strResultado As String
    Dim strNome As String
    If lstPesq.ListItems.Count = 0 Then Exit Sub
    
    Me.Tag = lstPesq.SelectedItem.Text
    strResultado = Trim(lstPesq.SelectedItem.Text)
    
    strNome = Trim(lstPesq.SelectedItem.SubItems(1))
    
    If tpInscricao = TipoInsc.InscCpfCnpj Then strResultado = Trim(lstPesq.SelectedItem.SubItems(2))
    
    Unload Me
    
    If Not Controle Is Nothing Then
         If cboTipoInscricao.ListIndex >= 1 Then
                If Len(strResultado) <= 14 Then
                    If InStr(1, strResultado, ".") > 1 Then
                       Controle.Formato = formCPF
                    Else
                       Controle.Formato = formCGC
                   End If
                Else
                    Controle.Formato = formCGC
                End If
         Else
           Controle.Formato = formNenhum
         End If
        Controle.Text = strResultado
        
        If Not Contribuinte Is Nothing Then
            Contribuinte.Text = strNome
        End If
        Controle.SetFocus
        DoEvents
        SendKeys "{TAB}"
    End If
        
    DoEvents
    
    Exit Sub
    
TrataErro:
    If Err.Number = 5 Then 'CONTROLE NAO PODE RECEBER O FOCO
        'Stop
        Resume Next
    Else
        Util.Erro Err.Description
    End If
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


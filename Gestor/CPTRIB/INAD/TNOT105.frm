VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TNOT105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TNOT105"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   27
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TNOT105.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   705
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin ActiveTabs.SSActiveTabs tabCND 
      Height          =   3900
      Left            =   1350
      TabIndex        =   11
      Top             =   2160
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   6879
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TNOT105.frx":2123
      Images          =   "TNOT105.frx":21D4
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3480
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   6138
         _Version        =   131082
         TabGuid         =   "TNOT105.frx":2E6D
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2715
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   720
            Width           =   8070
         End
         Begin VTOcx.txtVISUAL txtValidade 
            Height          =   300
            Left            =   6030
            TabIndex        =   22
            Tag             =   "Validade"
            Top             =   390
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   529
            Caption         =   "Validade"
            Text            =   ""
            Enabled         =   0   'False
            Formato         =   0
            Restricao       =   2
            MaxLen          =   10
         End
         Begin VTOcx.txtVISUAL txtFinalidade 
            Height          =   300
            Left            =   0
            TabIndex        =   23
            Tag             =   "Finalidade"
            Top             =   60
            Width           =   8010
            _ExtentX        =   14129
            _ExtentY        =   529
            Caption         =   "Finalidade"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtEmissao 
            Height          =   300
            Left            =   150
            TabIndex        =   24
            Tag             =   "Validade"
            Top             =   390
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   529
            Caption         =   "Emissão"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            MaxLen          =   10
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3480
         Index           =   1
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   6138
         _Version        =   131082
         TabGuid         =   "TNOT105.frx":2E95
         Begin VTOcx.grdVISUAL grdDebitos 
            Height          =   3030
            Left            =   30
            TabIndex        =   15
            Top             =   420
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   5345
            Caption         =   "Débitos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
         Begin VTOcx.txtVISUAL txtPerInicio 
            Height          =   300
            Left            =   30
            TabIndex        =   25
            Tag             =   "Periodo Inicial"
            Top             =   60
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
            Caption         =   "Periodo"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   6
         End
         Begin VTOcx.txtVISUAL txtPerFim 
            Height          =   300
            Left            =   2115
            TabIndex        =   26
            Tag             =   "Periodo Final"
            Top             =   60
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            Caption         =   "até"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   6
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3480
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   6138
         _Version        =   131082
         TabGuid         =   "TNOT105.frx":2EBD
         Begin VTOcx.grdVISUAL grdCertidoes 
            Height          =   3375
            Left            =   30
            TabIndex        =   17
            Top             =   45
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   5953
            Caption         =   "Certidões emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   6120
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   5940
         TabIndex        =   7
         Top             =   75
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdNova 
         Height          =   375
         Left            =   6930
         TabIndex        =   9
         Top             =   75
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Nova"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8805
         TabIndex        =   10
         Top             =   75
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7830
         TabIndex        =   8
         Top             =   75
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Formulario      =   "TNOT105"
      Descricao       =   "Emissao de Certidao"
      Icone           =   "TNOT105.frx":2EE5
      Codigo          =   "TNOT105"
   End
   Begin VTOcx.txtVISUAL txtIc 
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   529
      Caption         =   "Insc. Cadastral"
      Text            =   ""
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtIM 
      Height          =   300
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   529
      Caption         =   "Insc. Municipal"
      Text            =   ""
      Restricao       =   2
      Mascara         =   "00000000-00"
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   300
      Left            =   3330
      TabIndex        =   20
      Top             =   1080
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtContribuinte 
      Height          =   300
      Left            =   2790
      TabIndex        =   21
      Top             =   1440
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtEmissaoInicio 
      Height          =   300
      Left            =   600
      TabIndex        =   3
      Tag             =   "Periodo Inicial"
      Top             =   1800
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   529
      Caption         =   "Emissão"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtEmissaoFim 
      Height          =   300
      Left            =   2835
      TabIndex        =   4
      Tag             =   "Periodo Final"
      Top             =   1800
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      Caption         =   "até"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
   End
   Begin VTOcx.txtVISUAL txtNotifInicio 
      Height          =   300
      Left            =   6450
      TabIndex        =   5
      Tag             =   "Periodo Inicial"
      Top             =   1800
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      Caption         =   "Nº"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   8
   End
   Begin VTOcx.txtVISUAL txtNotifFim 
      Height          =   300
      Left            =   8010
      TabIndex        =   6
      Tag             =   "Periodo Final"
      Top             =   1800
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   529
      Caption         =   "até"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   8
   End
   Begin VB.Menu mnuNotifica 
      Caption         =   "."
      Visible         =   0   'False
      Begin VB.Menu mnuEmitir 
         Caption         =   "."
      End
   End
End
Attribute VB_Name = "TNOT105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Private Tipo As enuTipoCertidao
Private Certidao As cCertidao

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    tabCND.Tabs(1).Selected = True
    grdCertidoes.Preencher Bdados, ""
    grdDebitos.Preencher Bdados, ""
End Sub

Private Sub cmdBuscar_Click()
    If grdCertidoes.Preencher(Bdados, Certidao.exibirCertidoes(Tipo, CStr(cboTributo.Coluna(0).Valor), txtIm, txtEmissaoInicio, txtEmissaoFim, txtNotifInicio, txtNotifFim)) Then
        grdCertidoes.Mensagem = "Total : " & Format(grdCertidoes.Colunas(5).Soma, Const_Monetario)
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cabVisual.Exibir Bdados, Me.Tag, App.Path
    
    cboTributo.Preencher Bdados, "SELECT tip_nome_imposto, tip_cod_imposto FROM TAB_TRIBUTO ORDER BY tip_nome_imposto"
    
    Select Case Me.Tag
        Case "TNOT101": Tipo = tipCND
    End Select
    
    Set Certidao = New cCertidao
    txtTexto = Certidao.recuperarTexto(Tipo)
    txtEmissao = Date
    txtEmissao_LostFocus
End Sub

Private Sub grdCertidoes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not grdCertidoes.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuEmitir.Caption = "Emitir certidão nº " & grdCertidoes.SelectedItem
            mnuEmitir.Tag = grdCertidoes.SelectedItem.SubItems(1) & "|" & grdCertidoes.SelectedItem.SubItems(2) & "|" & grdCertidoes.SelectedItem.SubItems(3)
            Me.PopupMenu mnuNotifica
        End If
    End If
End Sub

Private Sub txtEmissao_LostFocus()
    txtValidade = Certidao.DataValidade(Tipo, txtEmissao)
End Sub

Private Sub txtic_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    If Trim$(txtIc) <> "" Then
        txtEndereco = ""
        Sql = "SELECT * FROM VIS_ENDERECO_IMOVEL WHERE tim_ic='" & txtIc & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtEndereco = "" & rs!TTL_NOME & " " & rs!tlg_nome & ", " & rs!tim_numero & " " & rs!tim_complemento & " - " & rs!TBA_NOME
        End If
        Bdados.FechaTabela rs
        Sql = "SELECT tim_tci_im FROM TAB_IMOVEL WHERE tim_ic='" & txtIc & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtIm = "" & rs!tim_tci_im
            txtIM_LostFocus
        End If
        Bdados.FechaTabela rs
    End If
End Sub

Private Sub txtIM_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    If Trim$(txtIm) <> "" Then
        txtContribuinte = ""
        Sql = "SELECT tci_nome FROM TAB_CONTRIBUINTE WHERE tci_im='" & txtIm & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            txtContribuinte = "" & rs!tci_nome
        End If
        Bdados.FechaTabela rs
    End If
End Sub

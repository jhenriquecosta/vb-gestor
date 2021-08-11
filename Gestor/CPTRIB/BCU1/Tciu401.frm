VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCIU401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   53
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "Tciu401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.cboVISUAL cboBairro 
      Height          =   510
      Left            =   6075
      TabIndex        =   7
      Top             =   2205
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   900
      Caption         =   "Bairro"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Alinhamento     =   1
   End
   Begin Threed.SSFrame fra 
      Height          =   5295
      Index           =   0
      Left            =   30
      TabIndex        =   27
      Top             =   720
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9340
      _Version        =   196610
      ForeColor       =   128
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " Opções de consulta "
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   1185
         TabIndex        =   4
         Top             =   1485
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         Caption         =   "Logradouro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         Editavel        =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   315
         Left            =   2445
         TabIndex        =   5
         Top             =   1680
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.ComboBox cboRelatorio 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "Tciu401.frx":2123
         Left            =   1635
         List            =   "Tciu401.frx":2130
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   300
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1695
         Width           =   795
      End
      Begin VB.TextBox txtIM 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   330
         MaxLength       =   11
         TabIndex        =   1
         Top             =   930
         Width           =   1245
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8955
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1695
         Width           =   540
      End
      Begin VB.TextBox txtSetor 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8385
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1695
         Width           =   480
      End
      Begin VB.TextBox txtContrib 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Top             =   930
         Width           =   8145
      End
      Begin VB.TextBox txtic 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   330
         TabIndex        =   10
         Top             =   3150
         Width           =   1425
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   8610
         TabIndex        =   26
         Top             =   4725
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7440
         TabIndex        =   25
         Top             =   4725
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   6300
         TabIndex        =   24
         Top             =   4725
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   3
         Left            =   5130
         TabIndex        =   23
         Top             =   4725
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VB.Frame fraCamposFicha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3030
         Left            =   210
         TabIndex        =   39
         Top             =   2220
         Width           =   9615
         Begin VB.TextBox txtICAnterior 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1620
            TabIndex        =   11
            Top             =   930
            Width           =   1185
         End
         Begin VTOcx.cboVISUAL cboLoteamento 
            Height          =   510
            Left            =   90
            TabIndex        =   51
            Top             =   90
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   900
            Caption         =   "Loteamento"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.txtVISUAL txtLote 
            Height          =   285
            Left            =   4530
            TabIndex        =   16
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            Caption         =   "Lote:"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtQuadraLote 
            Height          =   285
            Left            =   3120
            TabIndex        =   15
            Top             =   300
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            Caption         =   "Qd.:"
            Text            =   ""
         End
         Begin VB.ComboBox cboTipoImovel 
            DataField       =   "ttl_nome"
            DataSource      =   "dtTipLogr"
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":2175
            Left            =   3735
            List            =   "Tciu401.frx":217F
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "Tipo Imovel"
            Top             =   915
            Width           =   2370
         End
         Begin VB.ComboBox cboOcupLote 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":2199
            Left            =   6180
            List            =   "Tciu401.frx":219B
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Tag             =   "1"
            Top             =   915
            Width           =   3375
         End
         Begin VB.ComboBox cboAforado 
            DataField       =   "ttl_nome"
            DataSource      =   "dtTipLogr"
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":219D
            Left            =   2895
            List            =   "Tciu401.frx":21A7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "Logradouro"
            Top             =   915
            Width           =   765
         End
         Begin VB.ComboBox cboUso 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":21B5
            Left            =   90
            List            =   "Tciu401.frx":21B7
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "16"
            Top             =   1545
            Width           =   4725
         End
         Begin VB.ComboBox cboDestinacao 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":21B9
            Left            =   5040
            List            =   "Tciu401.frx":21BB
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "11"
            Top             =   1545
            Width           =   4515
         End
         Begin VB.ComboBox cboPadrao 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":21BD
            Left            =   90
            List            =   "Tciu401.frx":21BF
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Tag             =   "12"
            Top             =   2115
            Width           =   4725
         End
         Begin VB.ComboBox cboTipologia 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "Tciu401.frx":21C1
            Left            =   5040
            List            =   "Tciu401.frx":21C3
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "9"
            Top             =   2115
            Width           =   4515
         End
         Begin VB.TextBox txtValor02 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1605
            TabIndex        =   22
            Top             =   2700
            Width           =   1185
         End
         Begin VB.TextBox txtValor01 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   75
            TabIndex        =   21
            Top             =   2700
            Width           =   1110
         End
         Begin VTOcx.cboVISUAL cboEdificio 
            Height          =   510
            Left            =   6180
            TabIndex        =   52
            Top             =   90
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   900
            Caption         =   "Edificio"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IC"
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   50
            Top             =   690
            Width           =   165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aforado"
            Height          =   195
            Index           =   26
            Left            =   2940
            TabIndex        =   49
            Top             =   690
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   225
            Index           =   27
            Left            =   3765
            TabIndex        =   48
            Top             =   690
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupação"
            Height          =   195
            Index           =   28
            Left            =   6180
            TabIndex        =   47
            Top             =   690
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IC Anterior"
            Height          =   195
            Index           =   25
            Left            =   1605
            TabIndex        =   46
            Top             =   690
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uso"
            Height          =   195
            Index           =   29
            Left            =   90
            TabIndex        =   45
            Top             =   1305
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destinação"
            Height          =   195
            Index           =   30
            Left            =   5040
            TabIndex        =   44
            Top             =   1305
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Padrão"
            Height          =   195
            Index           =   31
            Left            =   90
            TabIndex        =   43
            Top             =   1905
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipologia"
            Height          =   195
            Index           =   32
            Left            =   5040
            TabIndex        =   42
            Top             =   1905
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Venal"
            Height          =   195
            Index           =   35
            Left            =   90
            TabIndex        =   41
            Top             =   2460
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "até"
            Height          =   195
            Index           =   36
            Left            =   1305
            TabIndex        =   40
            Top             =   2730
            Width           =   240
         End
      End
      Begin VTOcx.txtVISUAL txtNumero 
         Height          =   525
         Left            =   5310
         TabIndex        =   6
         Top             =   1470
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   926
         Caption         =   "Numero"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório"
         Height          =   195
         Index           =   23
         Left            =   930
         TabIndex        =   38
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod Logr"
         Height          =   195
         Index           =   20
         Left            =   285
         TabIndex        =   37
         Top             =   1485
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imóvel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   300
         TabIndex        =   34
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra"
         Height          =   195
         Index           =   7
         Left            =   8985
         TabIndex        =   33
         Top             =   1470
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
         Height          =   195
         Index           =   6
         Left            =   8400
         TabIndex        =   32
         Top             =   1470
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social"
         Height          =   195
         Index           =   5
         Left            =   1620
         TabIndex        =   31
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IM"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   30
         Top             =   720
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localização"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   29
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contribuinte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   28
         Top             =   540
         Width           =   1170
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1138
      Icone           =   "Tciu401.frx":21C5
   End
   Begin VTOcx.grdVISUAL grid 
      Height          =   2295
      Left            =   30
      TabIndex        =   35
      Top             =   6060
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   4048
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.Menu mnugeral 
      Caption         =   "geral"
      Visible         =   0   'False
      Begin VB.Menu mnuCad 
         Caption         =   "Consulta Cadastro"
      End
      Begin VB.Menu mnuLanca 
         Caption         =   "Consulta Lancamentos"
      End
   End
End
Attribute VB_Name = "TCIU401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cadastro As VSImposto
Private SelecaoRpt As String

Private Enum Relatorio
    Listagem = 1
    Aforamento = 2
    Ficha = 3
End Enum

Private OpcaoRelatorio As Relatorio

Sub PreencheImovel(RsAux As Object) 'Aki
    cboTipoLogr = RsAux!TTL_NOME
    cboLogr = RsAux!tlg_nome
    cboBairro = RsAux!TBA_NOME
    cboLoteamento = RsAux!tim_loteamento
    txtQuadra = RsAux!tim_QUADRA
End Sub

Sub PreencheContrib(RsAux As Object) 'Aki
    txtIM = RsAux!tim_tci_im
End Sub


Private Sub AtualizaUF(Combo As ComboBox)
   Combo.Clear
   Combo.AddItem "MA"
   Combo.AddItem "AC"
   Combo.AddItem "AM"
   Combo.AddItem "AP"
   Combo.AddItem "AL"
   Combo.AddItem "BA"
   Combo.AddItem "CE"
   Combo.AddItem "DF"
   Combo.AddItem "ES"
   Combo.AddItem "GO"
   Combo.AddItem "MG"
   Combo.AddItem "MS"
   Combo.AddItem "MT"
   Combo.AddItem "PA"
   Combo.AddItem "PB"
   Combo.AddItem "PE"
   Combo.AddItem "PI"
   Combo.AddItem "PR"
   Combo.AddItem "SC"
   Combo.AddItem "SE"
   Combo.AddItem "SP"
   Combo.AddItem "RJ"
   Combo.AddItem "RN"
   Combo.AddItem "RO"
   Combo.AddItem "RR"
   Combo.AddItem "RS"
   Combo.AddItem "TO"
End Sub


Private Sub cboRelatorio_Click()
    If cboRelatorio.ListIndex = 0 Then  ' Listagem dados cadastrais
        Habilita_Campos_Formulario Listagem
        OpcaoRelatorio = Listagem
    ElseIf cboRelatorio.ListIndex = 1 Then  ' Listagem dados Aforamento
        Habilita_Campos_Formulario Aforamento
        OpcaoRelatorio = Aforamento
    ElseIf cboRelatorio.ListIndex = 2 Then  ' Ficha Cadastral
        Habilita_Campos_Formulario Ficha
        OpcaoRelatorio = Ficha
    ElseIf cboRelatorio.ListIndex = -1 Then  ' Tela Limpa
        txtContrib.Enabled = False
        txtIM.Enabled = False
        txtCodLogr.Enabled = False
        cboTipoLogr.Enabled = False
        cboLogr.Enabled = False
        cboBairro.Enabled = False
        cboLoteamento.Enabled = False
        txtQuadra.Enabled = False
        txtIC.Enabled = False
        fraCamposFicha.Enabled = False
        Dim ContLabel As Integer
        For ContLabel = 25 To 32
            Label1(ContLabel).Enabled = True
        Next
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Resume Next
    Dim Sql As String
    Dim Condicao As String
    
    'If cboRelatorio.ListIndex = -1 Then Util.Avisa "Selecione o relatório.": Exit Sub
    OpcaoRelatorio = IIf(OpcaoRelatorio = 0, Listagem, OpcaoRelatorio)
    Select Case cmd(Index).Caption
        Case "&Buscar"
            Select Case OpcaoRelatorio
                Case Listagem, Aforamento
'                    If cboRelatorio.ListIndex = -1 Then Exit Sub
                    SelecaoRpt = ""
                    Condicao = ""
                    Sql = " SELECT tim_ic as IC,tim_ic_anterior as [IC Anterior], tci_nome AS CONTRIBUINTE, tim_tci_im AS IM, " & _
                        " tim_tlg_cod_logradouro AS [COD LOGR],TIM_TBA_COD_BAIRRO AS BAIRRO,tim_numero as No ,tim_complemento "
                    
                    If OpcaoRelatorio = Aforamento Then
                        Sql = Sql & ", TIM_AFORAMENTO_FICHA AS FICHA, TIM_AFORAMENTO_LIVRO AS LIVRO,TIM_AFORAMENTO_FOLHA AS FOLHA," & _
                        " TIM_AFORAMENTO_DATA AS DATA "
                    End If
                    
                    Sql = Sql & " FROM VIS_IMOVEL"
                        
                    If Trim(txtIM) <> "" Then
                        Condicao = Condicao & " and tim_tci_im ='" & txtIM & "'"
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_tci_im} = '" & txtIM & "'"
                    End If
                    
                    If Trim(cboLogr) <> "" Then
                        Condicao = Condicao & " and tim_tlg_cod_logradouro = " & cboLogr.Coluna(0).Valor
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_tlg_cod_logradouro} = '" & txtCodLogr & "'"
                    End If
                    
'                    If Trim(cboTipoLogr) <> "" Then
'                        Condicao = Condicao & " and tim_tlg_cod_logradouro = " & cboTipoLogr.Coluna(0).VALOR
'                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_tlg_cod_logradouro} = '" & txtCodLogr & "'"
'                    End If
                    
                    
                    If Trim(txtNumero) <> "" Then
                        Condicao = Condicao & " and tim_numero = '" & txtNumero & "'"
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_numero} = '" & txtNumero & "'"
                    End If
                    
                    If Trim(cboBairro) <> "" Then
                        Condicao = Condicao & " AND TIM_TBA_COD_BAIRRO = " & cboBairro.Coluna(1).Valor
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.TIM_TBA_COD_BAIRRO} = " & cboBairro.Coluna(1).Valor
                    End If
                    
                    If Trim(txtIC) <> "" Then
                        Condicao = Condicao & " AND tim_ic = '" & txtIC & "'"
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_ic} = '" & txtIC & "'"
                    End If
                    
                    If Trim(txtICAnterior) <> "" Then
                        Condicao = Condicao & " AND tim_ic_anterior = '" & txtICAnterior & "'"
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_ic_anterior} = '" & txtICAnterior & "'"
                    End If
                    
                    If Trim(cboLoteamento) <> "" Then
                        Condicao = Condicao & " and tim_loteamento = " & cboLoteamento.Coluna(0).Valor
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_loteamento} = " & cboLoteamento.Coluna(0).Valor
                    End If
                    
                    If Trim(txtQuadraLote) <> "" Then
                        Condicao = Condicao & " and tim_quadra = '" & txtQuadraLote & "'"
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_quadra} = '" & txtQuadraLote & "'"
                    End If
                    
                    If Trim(txtLote) <> "" Then
                        Condicao = Condicao & " and tim_lote = '" & txtLote & "'"
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_lote} = '" & txtLote & "'"
                    End If
                    
                    If Trim(cboEdificio) <> "" Then
                        Condicao = Condicao & " and tim_ted_cod_edificio = " & cboEdificio.Coluna(0).Valor
                        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_ted_cod_edificio} = " & cboEdificio.Coluna(0).Valor
                    End If
                    
                    If Trim(txtQuadra) <> "" Then
                        Condicao = Condicao & " and tim_ic like '____" & txtQuadra & "%'"
                        SelecaoRpt = SelecaoRpt & " and MID({VIS_IMOVEL.tim_ic},5,4) = '" & Format(txtQuadra, "0000") & "'"
                    End If
                    
                    If SelecaoRpt <> "" Then
                        Condicao = " where " & Right(Condicao, Len(Condicao) - 4)
                        SelecaoRpt = Right(SelecaoRpt, Len(SelecaoRpt) - 4)
                    End If
                    grid.Preencher Bdados, Sql & Condicao, 1500, 1000, 2000
                Case Ficha
                    Dim Query As String
                    Query = "SELECT distinct  tim_ic  as IC,"
                    Query = Query & " tim_tci_im as IM,tci_nome as Contribuinte,tim_tlg_cod_logradouro as CodLogr,"
                    Query = Query & " TTL_NOME as Logr,"
                    Query = Query & " tlg_nome as Nome,"
                    Query = Query & " tim_numero as [Nº],"
                    Query = Query & " TBA_NOME as Bairro,"
                    Query = Query & " tim_valor as [Valor(R$)],TLO_DESCRICAO AS Loteamento,"
                    Query = Query & " TIM_QUADRA AS Quadra,tim_lote as Lote,Ted_descricao as Edificio,"
                    Query = Query & " tim_ic_anterior as [Insc Anterio],tim_aforado as Aforado  "
                    Query = Query & " FROM "
                    Query = Query & " VIS_IMOVEL,tab_loteamento,tab_edificio, "
                    Query = Query & " TAB_DETALHE_IMOVEL "
                    Query = Query & " where "
                    Query = Query & " tim_ic = tdi_tim_ic and tim_loteamento = tlo_cod_loteamento "
                    Query = Query & " and TIM_TED_COD_EDIFICIO = TED_COD_EDIFICIO"
                                    
                    If cboRelatorio.ListIndex = -1 Then Exit Sub
                    Condicao = ""
                    Sql = "SELECT distinct  tim_ic  as IC,"
                    Sql = Sql & " tim_tci_im as IM,tci_nome as Contribuinte,tim_tlg_cod_logradouro as CodLogr,"
                    Sql = Sql & " TTL_NOME as Logr,"
                    Sql = Sql & " tlg_nome as Nome,"
                    Sql = Sql & " tim_numero as [Nº],"
                    Sql = Sql & " TBA_NOME as Bairro,"
                    Sql = Sql & " tim_valor as [Valor(R$)],TIM_QUADRA as Quadra,tim_lote as Lote,"
                    Sql = Sql & " TED_DESCRICAO AS Edificio,tim_ic_anterior as [Insc Anterio] ,tim_aforado as Aforado "
                    Sql = Sql & " FROM "
                    Sql = Sql & " VIS_IMOVEL,TAB_EDIFICIO, "
                    Sql = Sql & " TAB_DETALHE_IMOVEL "
                    Sql = Sql & " where "
                    Sql = Sql & " tim_ic = tdi_tim_ic AND TIM_TED_COD_EDIFICIO = TED_COD_EDIFICIO"
        
                    SelecaoRpt = "{TAB_BAIRRO.TBA_TMU_COD_MUNICIPIO}=" & Aplicacoes.Codigo_Municipio
                    SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_tmu_cod_municipio}=" & Aplicacoes.Codigo_Municipio
                    
                    'IC
                    If Trim(txtIC) <> "" Then
                        Condicao = " and tim_ic ='" & txtIC & "'"
                        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_ic}='" & txtIC & "'"
                    End If
                    'IM
                    If Trim(txtIM) <> "" Then
                        Condicao = Condicao & " and tim_tci_im = '" & txtIM & "'"
                        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_tci_im}='" & txtIM & "'"
                    End If
                    'Tipo Logradouro
                    If Trim(cboTipoLogr) <> "" Then
                        Condicao = Condicao & " and TTL_NOME = '" & cboTipoLogr & "'"
                        SelecaoRpt = SelecaoRpt & " and {TAB_TIPO_LOGR.TTL_NOME}='" & cboTipoLogr & "'"
                    End If
                    'Logradouro
                    If Trim(cboLogr) <> "" Then
                        Condicao = Condicao & " and tlg_nome = '" & cboLogr & "'"
                        SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_nome}= '" & cboLogr & "'"
                    End If
                    'Bairro
                    If Trim(cboBairro) <> "" Then
                        Condicao = Condicao & " and TBA_NOME = '" & cboBairro & "'"
                        SelecaoRpt = SelecaoRpt & " and {TAB_BAIRRO.TBA_NOME}= '" & cboBairro & "'"
                    End If
                    'Razao Social
                    If Trim(txtContrib) <> "" Then
                        Condicao = Condicao & " and (tci_nome like '%" & txtContrib & "%' or tci_nome like '%" & txtContrib & "%')"
                        SelecaoRpt = SelecaoRpt & " and ({Tab_Contribuinte.tci_nome} like '" & txtContrib & "*' or {Tab_Contribuinte.tci_nome} like '*" & txtContrib & "*')"
                    End If
                    'Quadra
                    If Trim(txtSetor) <> "" Then
                        Condicao = Condicao & " and substring(tim_ic,3,2) =  '" & Format(txtSetor, "00") & "'"
                        SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},3,2)='" & txtSetor & "'"
                    End If
                    
                    'Quadra
                    If Trim(txtQuadra) <> "" Then
                        If AplicacoesVTFuncoes.Municipio = "BALSAS" Then
                            Condicao = Condicao & " and substring(tim_ic,5,4)  = '" & Format(txtQuadra, "0000") & "'"
                            SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},5,4)='" & txtQuadra & "'"
                        Else
                            Condicao = Condicao & " and substring(tim_ic,5,3)  = '" & Format(txtQuadra, "000") & "'"
                            SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},5,3)='" & txtQuadra & "'"
                        End If
                    End If
                    'Ano construcao
                    If Trim(txtICAnterior) <> "" Then
                        Condicao = Condicao & " and tim_ic_anterior= " & txtICAnterior
                        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_ic_anterior}=" & txtICAnterior
                    End If
                    'Cod Logr
                    If Trim(txtCodLogr) <> "" Then
                        Condicao = Condicao & " and tim_tlg_cod_logradouro= '" & txtCodLogr & "'"
                        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_tlg_cod_logradouro}='" & txtCodLogr & "'"
                    End If
                    'Aforado
                    If Trim(cboAforado) <> "" Then
                        Condicao = Condicao & " and LTRIM(RTRIM(TIM_AFORAMENTO_NUMERO))" & IIf((cboAforado = "NAO"), "=", "<>") & "''"
                        SelecaoRpt = SelecaoRpt & " and TRIM({TAB_IMOVEL.TIM_AFORAMENTO_NUMERO})" & IIf((cboAforado = "NAO"), "=", "<>") & "''"
                    End If
                    If Trim(cboTipoImovel) <> "" Then
                        Condicao = Condicao & " and tim_tipo_imovel =" & IIf((cboTipoImovel = "PREDIAL"), "1", "2")
                        SelecaoRpt = SelecaoRpt & " and TRIM({TAB_IMOVEL.tim_tipo_imovel})" & IIf((cboTipoImovel = "PREDIAL"), "1", "2")
                    End If
                    If Trim(txtValor01) <> "" And Trim(txtValor02) <> "" Then
                        Condicao = Condicao & " and tim_valor >=" & CDbl(txtValor01) & " and tim_valor <=" & CDbl(txtValor02)
                        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_valor} >=" & CDbl(txtValor01) & " and {TAB_IMOVEL.tim_valor} <=" & CDbl(txtValor02)
                    End If
                    If cboLoteamento.ListIndex <> -1 Or cboLoteamento.Text <> "" Then
                        Condicao = Condicao & " and tim_loteamento =" & Bdados.Converte(cboLoteamento.Coluna(0).Valor, tctexto)
                    End If
                    If txtQuadraLote <> "" Then
                        Condicao = Condicao & " and TIM_QUADRA like '%" & txtQuadraLote & "%'"
                    End If
                    'Tipo
                    If txtLote <> "" Then
                        Condicao = Condicao & " and TIM_LOTE like '%" & txtLote & "%'"
                    End If
                    If cboEdificio.ListIndex <> -1 Or cboEdificio.Text <> "" Then
                        Condicao = Condicao & " and TIM_TED_COD_EDIFICIO = " & Bdados.Converte(cboEdificio.Coluna(0).Valor, tctexto)
                    End If
                    
                    'Ocupacao
                    'Uso
                    'Destinacao
                    'Padrao
                    'Tipologia
                    'Estrutura
                    'Conservacao
                    Dim Controle As Control
                    Dim CodComponente As String, VlrItem As Integer, CodGrupo As Integer
                    For Each Controle In Me.Controls
                        If IsNumeric(Controle.Tag) Then
                            If Controle.Text <> "" Then
                                CodComponente = Cadastro.BuscaCodItemAvancado(Controle.Text, Controle.Tag)
                                VlrItem = Controle.ListIndex + 1
                                CodGrupo = Controle.Tag
                                
                                Condicao = Condicao & " and tdi_tgc_cod_grupo= " & CodGrupo & " AND tdi_tco_cod_componente = " & CodComponente & " AND tdi_valor_item = " & VlrItem
                                SelecaoRpt = SelecaoRpt & " and {TAB_DETALHE_IMOVEL.tdi_tgc_cod_grupo}= " & CodGrupo
                                SelecaoRpt = SelecaoRpt & " and {TAB_DETALHE_IMOVEL.tdi_tco_cod_componente} = " & CodComponente
                                SelecaoRpt = SelecaoRpt & " and {TAB_DETALHE_IMOVEL.tdi_valor_item} = " & VlrItem
                            End If
                        End If
                    Next Controle
                    If cboLoteamento.ListIndex > -1 Or cboLoteamento.Text <> "" Then
                        Sql = Query
                    End If
                    Sql = Sql & Condicao
                    Screen.MousePointer = 11
                    If grid.Preencher(Bdados, Sql, 1400) Then
                        grid.Mensagem = "Total Valor Venal: R$" & Format(grid.Colunas(9).Soma, Const_Monetario)
                    Else
                        grid.Mensagem = ""
                        Avisa "Nenhum registro encontrado."
                    End If
                    Screen.MousePointer = 0
                    DoEvents
                End Select
        Case "Sai&r"
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub


Private Sub cmdImprimir_Click()
    
    Screen.MousePointer = 11
    If grid.ListItems.Count > 0 Then
        With Rpt
            If cboRelatorio.ListIndex = 0 Then  ' Listagem dados cadastrais
                If Not .DefinirArquivo(Bdados, App.Path & "\TListagemImoveis.rpt") Then Exit Sub
                .Selecao = SelecaoRpt
                .Arvore = False
                .Visualizar
                DoEvents
            ElseIf cboRelatorio.ListIndex = 1 Then  ' Listagem dados Aforamento
                If Not .DefinirArquivo(Bdados, App.Path & "\TListagemAforamento.rpt") Then Exit Sub
                .Selecao = SelecaoRpt
                .Arvore = False
                .Visualizar
                DoEvents
            ElseIf cboRelatorio.ListIndex = 2 Then  ' Ficha Cadastral
                If Not .DefinirArquivo(Bdados, App.Path & "\TCIU201.rpt") Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                .Selecao = SelecaoRpt
                .Titulo = "Ficha Cadastral"
                .Arvore = False
                .Visualizar
                DoEvents
            Else
                Avisa "Selecione uma opção de impressão."
                cboRelatorio.SetFocus
            End If
        End With
    End If
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    SelecaoRpt = ""
    
    Edita.LimpaCampos Me
    grid.Preencher Bdados, ""
    cboRelatorio.ListIndex = -1
    cboRelatorio.SetFocus
End Sub

Private Sub Form_Activate()
    
    Habilita_Campos_Formulario Listagem
    Habilita_Campos_Formulario Aforamento
    Habilita_Campos_Formulario Ficha
    Dim ContLabel As Integer
    For ContLabel = 25 To 32
        Label1(ContLabel).Enabled = True
    Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
            
    Dim Controle As Control
    Dim i As Byte
    Set Cadastro = New VSImposto
    
    cboLogr.Preencher Bdados, "Select tlg_cod_logradouro,(tlg_nome) From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio, 1
    cboTipoLogr.Preencher Bdados, "Select TTL_COD_TIP_LOGR ,(ttl_nome) From Tab_Tipo_Logr", 1
    cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
    cboLoteamento.Preencher Bdados, "Select TLO_COD_LOTEAMENTO,TLO_DESCRICAO from TAB_LOTEAMENTO ORDER BY TLO_DESCRICAO", 1
    cboEdificio.Preencher Bdados, "Select TED_COD_EDIFICIO,TED_DESCRICAO from TAB_EDIFICIO ORDER BY TED_DESCRICAO", 1
    For Each Controle In Controls
        If IsNumeric(Controle.Tag) Then
            If Val(Controle.Tag) < 20 Then Call Edita.AtualizaCombo(Bdados, Controle, "Select tco_descricao_componente From Tab_Componente_Avancado Where tco_grupo = " & Controle.Tag & " order by tco_cod_componente asc")
        End If
    Next
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    AtualizaCabecalho grid
    cboRelatorio.ListIndex = -1
    cboRelatorio_Click
End Sub

'Private Sub grid_DblClick()
''If grid.ListItems.Count > 0 Then
''    TCIU201.Tag = grid.SelectedItem
''    TCIU201.Show
''End If
'End Sub

Private Sub txtAnoAq_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 And grid.ListItems.Count > 0 Then
        mnuCad.Caption = "Consultar Cadastro " & grid.SelectedItem
        mnuLanca.Caption = "Consultar Lancamentos " & grid.SelectedItem
        Me.PopupMenu mnugeral
    End If
End Sub

Private Sub mnuCad_Click()
    TCIU201.Tag = grid.SelectedItem
    TCIU201.Show
End Sub


Private Sub mnuLanca_Click()
    Dim ProjObrig As Object
    
    Set ProjObrig = CreateObject("VSTOBRI.Aplicacoes")
    
    Set ProjObrig.Banco = Bdados.Conexao
    ProjObrig.Usuario = AplicacoesVTFuncoes.Usuario
    ProjObrig.Codigo_Municipio = AplicacoesVTFuncoes.Codigo_Municipio
    ProjObrig.Municipio = AplicacoesVTFuncoes.Municipio
    TempContrib = Trim(grid.SelectedItem)
    ProjObrig.Abre_Aplicacao "TOBR401", 0, Cod_sis, Sistema, Desc_Form, Trim(grid.SelectedItem)
    TempContrib = Trim(grid.SelectedItem)
    
End Sub

Private Sub txtCodLogr_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_Validate(Cancel As Boolean)
Dim Sql As String
    If Trim(txtCodLogr) <> "" Then
        If Bdados.AbreTabela("SELECT tlg_ttl_cod_tip_logr,tlg_cod_logradouro FROM TAB_LOGRADOURO WHERE  tlg_cod_logradouro = '" & txtCodLogr & "'") Then
            cboTipoLogr.SetarLinha Bdados.Tabela(0), 0
            cboLogr.SetarLinha Bdados.Tabela(1), 0
            
            Sql = "SELECT TBA_NOME, TBA_COD_BAIRRO From TAB_BAIRRO WHERE TBA_COD_BAIRRO IN " & _
                " (SELECT DISTINCT TTC_TBA_COD_BAIRRO FROM TAB_TRECHO WHERE TTC_TLG_COD_LOGRADOURO = " & _
                txtCodLogr & ") AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio
            cboBairro.Preencher Bdados, Sql, 1
        Else
            cboTipoLogr.ListIndex = -1
            cboLogr.ListIndex = -1
            cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio, 1
        End If
    Else
        cboTipoLogr.ListIndex = -1
        cboLogr.ListIndex = -1
        cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio, 1
    End If

End Sub

Private Sub txtContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtIC_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus()
    If Me.ActiveControl.Name = "cmdLimpar" Then Exit Sub
    txtIC = Cadastro.FormataInscricao(txtIC, InscImovel)
End Sub

Private Sub txtIM_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIM_LostFocus()
    txtIM = Cadastro.FormataInscricao(txtIM, InscContrib)
    If Trim(txtIM) <> "" Then
        If Bdados.AbreTabela("SELECT TCI_NOME FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & txtIM & "'") Then
            txtContrib = Bdados.Tabela(0)
        End If
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then Exit Sub
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValor01_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValor01_LostFocus()
    txtValor01 = Edita.FormataTexto(txtValor01, Monetario)
End Sub

Private Sub txtValor02_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValor02_LostFocus()
    txtValor02 = Edita.FormataTexto(txtValor02, Monetario)
End Sub

Private Sub Habilita_Campos_Formulario(ModoBusca As Relatorio)
    
Dim ContLabel As Integer

    Select Case ModoBusca
    
        Case Listagem, Aforamento
            txtIM.Enabled = True
            txtCodLogr.Enabled = True
            cboTipoLogr.Enabled = True
            cboLogr.Enabled = True
            cboBairro.Enabled = True
            cboLoteamento.Enabled = True
            txtQuadra.Enabled = True
            txtIC.Enabled = True
            fraCamposFicha.Enabled = False
            txtContrib.Enabled = False
            For ContLabel = 25 To 32
                Label1(ContLabel).Enabled = False
            Next
        Case Ficha
            txtContrib.Enabled = True
            txtIM.Enabled = True
            txtCodLogr.Enabled = True
            cboTipoLogr.Enabled = True
            cboLogr.Enabled = True
            cboBairro.Enabled = True
            cboLoteamento.Enabled = True
            txtQuadra.Enabled = True
            txtIC.Enabled = True
            fraCamposFicha.Enabled = True
            
            For ContLabel = 25 To 32
                Label1(ContLabel).Enabled = True
            Next
        Case Else
    End Select
End Sub



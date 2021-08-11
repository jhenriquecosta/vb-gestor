VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#1.1#0"; "VTControles.ocx"
Begin VB.Form TCIM404 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   9105
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
   ScaleHeight     =   9105
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cboVISUAL cboBairro 
      Height          =   510
      Left            =   5655
      TabIndex        =   6
      Top             =   2145
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   900
      Caption         =   "Bairro"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Alinhamento     =   1
   End
   Begin Threed.SSFrame fra 
      Height          =   5730
      Index           =   0
      Left            =   30
      TabIndex        =   28
      Top             =   660
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   10107
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
         Left            =   1155
         TabIndex        =   4
         Top             =   1485
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   900
         Caption         =   "Logradouro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   315
         Left            =   2505
         TabIndex        =   5
         Top             =   1680
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.ComboBox cboRelatorio 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "TCIM404.frx":0000
         Left            =   1635
         List            =   "TCIM404.frx":0010
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
         TabIndex        =   8
         Top             =   1695
         Width           =   810
      End
      Begin VB.TextBox txtLoteamento 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8145
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1695
         Width           =   750
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
         TabIndex        =   9
         Top             =   2460
         Width           =   1425
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   3750
         TabIndex        =   27
         Top             =   5190
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
         Left            =   2580
         TabIndex        =   26
         Top             =   5190
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
         Left            =   1440
         TabIndex        =   25
         Top             =   5190
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
         Left            =   270
         TabIndex        =   24
         Top             =   5190
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
         Height          =   2850
         Left            =   210
         TabIndex        =   40
         Top             =   2220
         Width           =   9615
         Begin VB.ComboBox cboTipoImovel 
            DataField       =   "ttl_nome"
            DataSource      =   "dtTipLogr"
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":0078
            Left            =   3735
            List            =   "TCIM404.frx":0082
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Tag             =   "Tipo Imovel"
            Top             =   225
            Width           =   2370
         End
         Begin VB.ComboBox cboOcupLote 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":009C
            Left            =   6180
            List            =   "TCIM404.frx":009E
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "1"
            Top             =   225
            Width           =   3375
         End
         Begin VB.TextBox txtAnoConst 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1620
            TabIndex        =   10
            Tag             =   "111"
            Top             =   240
            Width           =   1185
         End
         Begin VB.ComboBox cboAforado 
            DataField       =   "ttl_nome"
            DataSource      =   "dtTipLogr"
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00A0
            Left            =   2895
            List            =   "TCIM404.frx":00AA
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "Logradouro"
            Top             =   225
            Width           =   765
         End
         Begin VB.ComboBox cboUso 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00B8
            Left            =   90
            List            =   "TCIM404.frx":00BA
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Tag             =   "16"
            Top             =   810
            Width           =   4725
         End
         Begin VB.ComboBox cboDestinacao 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00BC
            Left            =   5040
            List            =   "TCIM404.frx":00BE
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "11"
            Top             =   810
            Width           =   4515
         End
         Begin VB.ComboBox cboPadrao 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00C0
            Left            =   90
            List            =   "TCIM404.frx":00C2
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Tag             =   "12"
            Top             =   1380
            Width           =   4725
         End
         Begin VB.ComboBox cboEstrutura 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00C4
            Left            =   90
            List            =   "TCIM404.frx":00C6
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "10"
            Top             =   1950
            Width           =   4725
         End
         Begin VB.ComboBox cboConservacao 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00C8
            Left            =   5040
            List            =   "TCIM404.frx":00CA
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Tag             =   "13"
            Top             =   1950
            Width           =   4515
         End
         Begin VB.ComboBox cboTipologia 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "TCIM404.frx":00CC
            Left            =   5040
            List            =   "TCIM404.frx":00CE
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "9"
            Top             =   1380
            Width           =   4515
         End
         Begin VB.TextBox txtValor02 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1605
            TabIndex        =   21
            Top             =   2565
            Width           =   1185
         End
         Begin VB.TextBox txtValor01 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   75
            TabIndex        =   20
            Top             =   2565
            Width           =   1110
         End
         Begin VTOcx.txtVISUAL TxtPeriodo1 
            Height          =   345
            Left            =   4230
            TabIndex        =   22
            Top             =   2490
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   609
            Caption         =   "Periodo Inicial"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
         End
         Begin VTOcx.txtVISUAL TxtPeriodo2 
            Height          =   345
            Left            =   6990
            TabIndex        =   23
            Top             =   2490
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   609
            Caption         =   "Periodo Final"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IC"
            Height          =   195
            Index           =   8
            Left            =   105
            TabIndex        =   53
            Top             =   30
            Width           =   165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aforado"
            Height          =   195
            Index           =   26
            Left            =   2940
            TabIndex        =   52
            Top             =   0
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   225
            Index           =   27
            Left            =   3765
            TabIndex        =   51
            Top             =   0
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupação"
            Height          =   195
            Index           =   28
            Left            =   6180
            TabIndex        =   50
            Top             =   0
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ano construção"
            Height          =   195
            Index           =   25
            Left            =   1620
            TabIndex        =   49
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uso"
            Height          =   195
            Index           =   29
            Left            =   90
            TabIndex        =   48
            Top             =   570
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destinação"
            Height          =   195
            Index           =   30
            Left            =   5040
            TabIndex        =   47
            Top             =   570
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Padrão"
            Height          =   195
            Index           =   31
            Left            =   90
            TabIndex        =   46
            Top             =   1170
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipologia"
            Height          =   195
            Index           =   32
            Left            =   5040
            TabIndex        =   45
            Top             =   1170
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estrutura"
            Height          =   195
            Index           =   33
            Left            =   90
            TabIndex        =   44
            Top             =   1740
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conservação"
            Height          =   195
            Index           =   34
            Left            =   5040
            TabIndex        =   43
            Top             =   1740
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Venal"
            Height          =   195
            Index           =   35
            Left            =   90
            TabIndex        =   42
            Top             =   2325
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "até"
            Height          =   195
            Index           =   36
            Left            =   1305
            TabIndex        =   41
            Top             =   2595
            Width           =   240
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relatório"
         Height          =   195
         Index           =   23
         Left            =   930
         TabIndex        =   39
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
         TabIndex        =   38
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
         Left            =   180
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1470
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
         Height          =   195
         Index           =   6
         Left            =   8160
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   540
         Width           =   1170
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1138
      Icone           =   "TCIM404.frx":00D0
   End
   Begin VTOcx.grdVISUAL grid 
      Height          =   2595
      Left            =   30
      TabIndex        =   36
      Top             =   6450
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   4339
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
End
Attribute VB_Name = "TCIM404"
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
    ImovAtvEcon = 4
End Enum

Private OpcaoRelatorio As Relatorio

Sub PreencheImovel(RsAux As Object) 'Aki
    cboTipoLogr = RsAux!TTL_NOME
    cboLogr = RsAux!tlg_nome
    cboBairro = RsAux!TBA_NOME
    txtLoteamento = RsAux!Tim_loteamento
    txtQuadra = RsAux!tim_quadra
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
    Select Case cboRelatorio.ListIndex
        Case 0 ' Listagem dados cadastrais
            Habilita_Campos_Formulario Listagem
            OpcaoRelatorio = Listagem
            
        Case 1 ' Listagem dados Aforamento
            Habilita_Campos_Formulario Aforamento
            OpcaoRelatorio = Aforamento
        
        Case 2 ' Ficha Cadastral
            Habilita_Campos_Formulario Ficha
            OpcaoRelatorio = Ficha
    
        Case 3 'Imoveis com Ativ. Economica
            Habilita_Campos_Formulario ImovAtvEcon
            OpcaoRelatorio = ImovAtvEcon
            
        Case Else ' Tela Limpa
            txtContrib.Enabled = False
            txtIM.Enabled = False
            txtCodLogr.Enabled = False
            cboTipoLogr.Enabled = False
            cboLogr.Enabled = False
            cboBairro.Enabled = False
            txtLoteamento.Enabled = False
            txtQuadra.Enabled = False
            txtIc.Enabled = False
            fraCamposFicha.Enabled = False
            Dim ContLabel As Integer
            For ContLabel = 25 To 36
                Label1(ContLabel).Enabled = True
            Next
    End Select
    
End Sub

Private Sub BuscarListagem()
    Dim Sql As String
    Dim Condicao As String
    Dim Aux As Byte
    
    If cboRelatorio.ListIndex = -1 Then Exit Sub
    Condicao = ""
    Aux = 0
    Sql = " SELECT tim_ic as IC, tci_nome AS CONTRIBUINTE, tim_tci_im AS IM, " & _
        " tim_tlg_cod_logradouro AS [COD LOGR],TIM_TBA_COD_BAIRRO AS BAIRRO "
    
    If OpcaoRelatorio = Aforamento Then
        Sql = Sql & ", TIM_AFORAMENTO_FICHA AS FICHA, TIM_AFORAMENTO_LIVRO AS LIVRO,TIM_AFORAMENTO_FOLHA AS FOLHA," & _
        " TIM_AFORAMENTO_DATA AS DATA "
    End If
    
    Sql = Sql & " FROM VIS_IMOVEL"
        
    If Trim(txtIM) <> "" Then
        Condicao = Condicao & " where tim_tci_im ='" & txtIM & "'"
        Aux = 1
    End If
    
    If Trim(txtCodLogr) <> "" Then
        Condicao = Condicao & IIf(Aux = 0, " where ", " and ") & " tim_tlg_cod_logradouro = " & txtCodLogr
        Aux = 1
    End If
    
    If Trim(cboBairro) <> "" Then
        Condicao = Condicao & IIf(Aux = 0, " where ", " and ") & " TIM_TBA_COD_BAIRRO = " & cboBairro.Coluna(1).Valor
        Aux = 1
    End If
    
    If Trim(txtIc) <> "" Then
        Condicao = Condicao & IIf(Aux = 0, " where ", " and ") & " tim_ic = '" & txtIc & "'"
        Aux = 1
    End If
    
    If Trim(txtLoteamento) <> "" Then
        Condicao = Condicao & IIf(Aux = 0, " where ", " and ") & " tim_ic like '__" & txtLoteamento & "%'"
        Aux = 1
    End If

    If Trim(txtQuadra) <> "" Then
        Condicao = Condicao & IIf(Aux = 0, " where ", " and ") & Bdados.ParteTexto("tim_ic", MidVs, 5, 3, True) & " = '" & Format(txtQuadra, "000") & "'"
        Aux = 1
    End If
    
    If SelecaoRpt <> "" Then
        Condicao = " where " & Right(Condicao, Len(Condicao) - 4)
    End If
    If Trim(TxtPeriodo1) <> "" And Trim(TxtPeriodo2) <> "" Then
        Condicao = Condicao & IIf(Aux = 0, " where ", " and ") & " TIM_DATA_REGISTRO >= " & Bdados.Converte(TxtPeriodo1, TCDataHora) & " And TIM_DATA_REGISTRO <= " & Bdados.Converte(TxtPeriodo2, TCDataHora)
    End If
    grid.Preencher Bdados, Sql & Condicao

End Sub

Private Function PrepararRelListagem() As Boolean
    If cboRelatorio.ListIndex = -1 Then Exit Function
    SelecaoRpt = ""
    If Trim(txtIM) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_tci_im} = '" & txtIM & "'"
    End If
    
    If Trim(txtCodLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_tlg_cod_logradouro} = '" & txtCodLogr & "'"
    End If
    
    If Trim(cboBairro) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.TIM_TBA_COD_BAIRRO} = " & cboBairro.Coluna(1).Valor
    End If
    
    If Trim(txtIc) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL.tim_ic} = '" & txtIc & "'"
    End If
    
    If Trim(txtLoteamento) <> "" Then
        SelecaoRpt = SelecaoRpt & " and MID({VIS_IMOVEL.tim_ic},3,2) = '" & Format(txtLoteamento, "00") & "'"
    End If

    If Trim(txtQuadra) <> "" Then
        SelecaoRpt = SelecaoRpt & " and MID({VIS_IMOVEL.tim_ic},5,3) = '" & Format(txtQuadra, "000") & "'"
    End If
    
    If SelecaoRpt <> "" Then
        SelecaoRpt = Right(SelecaoRpt, Len(SelecaoRpt) - 4)
    End If
    If Trim(TxtPeriodo1) <> "" And Trim(TxtPeriodo2) <> "" Then
        SelecaoRpt = SelecaoRpt & " {TAB_IMOVEL.TIM_DATA_REGISTRO} >=" & Bdados.Converte(TxtPeriodo1, TCDataHora) & " AND {TAB_IMOVEL.TIM_DATA_REGISTRO} <= " & Bdados.Converte(TxtPeriodo2, TCDataHora)
    End If
    PrepararRelListagem = True
End Function
Private Sub BuscarFicha()
    Dim Sql As String
    Dim Condicao As String
    
    If cboRelatorio.ListIndex = -1 Then Exit Sub
    Condicao = ""
    '                            "TIM_UNIDADE AS [1ª Unidade],"
    Sql = "SELECT distinct  tim_ic  as IC," & _
                     "tim_tci_im as IM,tci_nome as Contribuinte,tim_tlg_cod_logradouro as CodLogr," & _
                    "TTL_NOME as Logr," & _
                    "tlg_nome as Nome," & _
                    "tim_numero as [Nº]," & _
                    "TBA_NOME as Bairro," & _
                    "tim_valor as [Valor(R$)] " & _
            "FROM Tab_Contribuinte," & _
                    "TAB_IMOVEL, " & _
                    " " & _
                    "vis_bvt " & _
                    ", " & _
                    "TAB_DETALHE_IMOVEL " & _
            "where tci_im = tim_tci_im AND " & _
                    "tim_tlg_cod_logradouro = tlg_cod_logradouro AND " & _
                    "  " & _
                    " " & _
                    " " & _
                    " tim_ic = tdi_tim_ic "

    
    'IC
    If Trim(txtIc) <> "" Then
        Condicao = " and tim_ic ='" & txtIc & "'"
    End If
    'IM
    If Trim(txtIM) <> "" Then
        Condicao = Condicao & " and tim_tci_im = '" & txtIM & "'"
    End If
    'Tipo Logradouro
    If Trim(cboTipoLogr) <> "" Then
        Condicao = Condicao & " and TTL_NOME = '" & cboTipoLogr & "'"
    End If
    'Logradouro
    If Trim(cboLogr) <> "" Then
        Condicao = Condicao & " and tlg_nome = '" & cboLogr & "'"
    End If
    'Bairro
    If Trim(cboBairro) <> "" Then
        Condicao = Condicao & " and TBA_NOME = '" & cboBairro & "'"
    End If
    'Razao Social
    If Trim(txtContrib) <> "" Then
        Condicao = Condicao & " and (tci_nome like '" & txtContrib & "%' or tci_nome like '%" & txtContrib & "%')"
    End If
    'Setor
    If Trim(txtLoteamento) <> "" Then
        Condicao = Condicao & " and " & Bdados.ParteTexto("tim_ic", MidVs, 3, 2, True) & " = '" & Format(txtLoteamento, "00") & "'"
    End If
    'Quadra
    If Trim(txtQuadra) <> "" Then
        Condicao = Condicao & " and " & Bdados.ParteTexto("tim_ic", MidVs, 5, 3, True) & " = '" & Format(txtQuadra, "000") & "'"
    End If
    'Ano construcao
    If Trim(txtAnoConst) <> "" Then
        Condicao = Condicao & " and tim_ano_aquis= " & txtAnoConst
    End If
    'Cod Logr
    If Trim(txtCodLogr) <> "" Then
        Condicao = Condicao & " and tim_tlg_cod_logradouro= '" & txtCodLogr & "'"
    End If
    'Aforado
    If Trim(cboAforado) <> "" Then
        Condicao = Condicao & " and LTRIM(RTRIM(TIM_AFORAMENTO_NUMERO))" & IIf((cboAforado = "NAO"), "=", "<>") & "''"
    End If
    If Trim(cboTipoImovel) <> "" Then
        Condicao = Condicao & " and tim_tipo_imovel =" & IIf((cboTipoImovel = "PREDIAL"), "1", "2")
    End If
    If Trim(txtValor01) <> "" And Trim(txtValor02) <> "" Then
        Condicao = Condicao & " and tim_valor >=" & CDbl(txtValor01) & " and tim_valor <=" & CDbl(txtValor02)
    End If
    If Trim(TxtPeriodo1) <> "" And Trim(TxtPeriodo2) <> "" Then
        Condicao = Condicao & " and TIM_DATA_REGISTRO >=" & Bdados.Converte(TxtPeriodo1, TCDataHora) & " and TIM_DATA_REGISTRO <=" & Bdados.Converte(TxtPeriodo2, TCDataHora)
    End If
    'Tipo
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
            End If
        End If
    Next Controle

    Sql = Sql & Condicao
    Screen.MousePointer = 11
    grid.Preencher Bdados, Sql, 1400
    If grid.ListItems.Count > 0 Then
        grid.Mensagem = "Total Valor Venal: R$" & Format(grid.Colunas(9).Soma, Const_Monetario)
    Else
        grid.Mensagem = ""
    End If
    Screen.MousePointer = 0
    DoEvents
End Sub

Private Function PrepararRelFicha() As Boolean
    If cboRelatorio.ListIndex = -1 Then Exit Function
    
    SelecaoRpt = "{TAB_BAIRRO.TBA_TMU_COD_MUNICIPIO}=" & Aplicacoes.Codigo_Municipio
    SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_tmu_cod_municipio}=" & Aplicacoes.Codigo_Municipio
    
    'IC
    If Trim(txtIc) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_ic}='" & txtIc & "'"
    End If
    'IM
    If Trim(txtIM) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_tci_im}='" & txtIM & "'"
    End If
    'Tipo Logradouro
    If Trim(cboTipoLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_TIPO_LOGR.TTL_NOME}='" & cboTipoLogr & "'"
    End If
    'Logradouro
    If Trim(cboLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_nome}= '" & cboLogr & "'"
    End If
    'Bairro
    If Trim(cboBairro) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_BAIRRO.TBA_NOME}= '" & cboBairro & "'"
    End If
    'Razao Social
    If Trim(txtContrib) <> "" Then
        SelecaoRpt = SelecaoRpt & " and ({Tab_Contribuinte.tci_nome} like '" & txtContrib & "*' or {Tab_Contribuinte.tci_nome} like '*" & txtContrib & "*')"
    End If
    'Setor
    If Trim(txtLoteamento) <> "" Then
        SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},3,2)='" & txtLoteamento & "'"
    End If
    'Quadra
    If Trim(txtQuadra) <> "" Then
        SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},5,3)='" & Format(txtQuadra, "000") & "'"
    End If
    'Ano construcao
    If Trim(txtAnoConst) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_ano_aquis}=" & txtAnoConst
    End If
    'Cod Logr
    If Trim(txtCodLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_tlg_cod_logradouro}='" & txtCodLogr & "'"
    End If
    'Aforado
    If Trim(cboAforado) <> "" Then
        SelecaoRpt = SelecaoRpt & " and TRIM({TAB_IMOVEL.TIM_AFORAMENTO_NUMERO})" & IIf((cboAforado = "NAO"), "=", "<>") & "''"
    End If
    If Trim(cboTipoImovel) <> "" Then
        SelecaoRpt = SelecaoRpt & " and TRIM({TAB_IMOVEL.tim_tipo_imovel})" & IIf((cboTipoImovel = "PREDIAL"), "1", "2")
    End If
    If Trim(txtValor01) <> "" And Trim(txtValor02) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_valor} >=" & CDbl(txtValor01) & " and {TAB_IMOVEL.tim_valor} <=" & CDbl(txtValor02)
    End If
    If Trim(TxtPeriodo1) <> "" And Trim(TxtPeriodo2) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.TIM_DATA_REGISTRO} >=" & Bdados.Converte(TxtPeriodo1, TCDataHora) & " and {TAB_IMOVEL.TIM_DATA_REGISTRO} <=" & Bdados.Converte(TxtPeriodo1, TCDataHora)
    End If
    'Tipo
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
                
                SelecaoRpt = SelecaoRpt & " and {TAB_DETALHE_IMOVEL.tdi_tgc_cod_grupo}= " & CodGrupo
                SelecaoRpt = SelecaoRpt & " and {TAB_DETALHE_IMOVEL.tdi_tco_cod_componente} = " & CodComponente
                SelecaoRpt = SelecaoRpt & " and {TAB_DETALHE_IMOVEL.tdi_valor_item} = " & VlrItem
            End If
        End If
    Next Controle
    PrepararRelFicha = True
End Function
Private Sub BuscarImovAtvEcon()
    Dim Sql As String
    Dim Condicao As String
    
    If cboRelatorio.ListIndex = -1 Then Exit Sub
    SelecaoRpt = ""
    Condicao = ""
        
    '1.
    Sql = "SELECT " & _
                " VIS_DETALHE_IMOVEL_BP_TABULAR.DESTINACAO," & _
                " VIS_IMOVEL.tim_ic," & _
                " VIS_IMOVEL.tim_tci_im," & _
                " VIS_IMOVEL.tci_nome," & _
                " VIS_IMOVEL.TTL_NOME " & Bdados.Concatena & "' '" & Bdados.Concatena & " VIS_IMOVEL.tlg_nome," & _
                " VIS_IMOVEL.TBA_NOME," & _
                " VIS_IMOVEL.tim_numero" & _
        " FROM" & _
            " VIS_DETALHE_IMOVEL_BP_TABULAR," & _
            " VIS_IMOVEL," & _
            " VIS_BVT_COMPLETO" & _
        " WHERE" & _
            " VIS_DETALHE_IMOVEL_BP_TABULAR.TDI_TIM_IC = VIS_IMOVEL.tim_ic AND " & _
            " VIS_BVT_COMPLETO.TTC_SETOR <> '' and" & _
            " not (VIS_DETALHE_IMOVEL_BP_TABULAR.DESTINACAO in ('RELIGIOSO', 'RESIDENCIAL', 'SERVICO PUBLICO COMUNITARIO'))"
    
    '2.1 Setor
    If Trim$(txtLoteamento) <> "" Then
        Condicao = Condicao & " and " & Bdados.ParteTexto("VIS_IMOVEL.tim_ic", MidVs, 3, 2, True) & " = '" & Format(txtLoteamento, "00") & "'"
    End If
    '2.2 Quadra
    If Trim(txtQuadra) <> "" Then
        Condicao = Condicao & " and " & Bdados.ParteTexto("tim_ic", MidVs, 5, 3, True) & " = '" & Format(txtQuadra, "000") & "'"
    End If
    '2.3 Destinacao
    If cboDestinacao <> "" Then
        SelecaoRpt = SelecaoRpt & " and VIS_DETALHE_IMOVEL_BP_TABULAR.DESTINACAO='" & cboDestinacao & "'"
    End If
'    grid.Preencher Bdados, Sql & Condicao
End Sub
Private Function PrepararRelImovAtivEcon() As Boolean
    If cboRelatorio.ListIndex = -1 Then Exit Function
    SelecaoRpt = ""
        
    '1.
    SelecaoRpt = " {@VT_Quadra} <> '' and" & _
                    " {VIS_BVT.TTC_SETOR} <> '' and" & _
                    " not ({VIS_DETALHE_IMOVEL_BP_TABULAR.DESTINACAO} in ['RELIGIOSO', 'RESIDENCIAL', 'SERVICO PUBLICO COMUNITARIO'])"


    '2.1 Setor
    If Trim$(txtLoteamento) <> "" Then
        SelecaoRpt = SelecaoRpt & " and MID({VIS_IMOVEL.tim_ic},3,2) = '" & Format(txtLoteamento, "00") & "' "
    End If
    '2.2 Quadra
    If Trim(txtQuadra) <> "" Then
        SelecaoRpt = SelecaoRpt & " and Mid({VIS_IMOVEL.tim_ic},5,3)='" & Format(txtQuadra, "000") & "'"
    End If
    '2.3 Destinacao
    If cboDestinacao <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_DETALHE_IMOVEL_BP_TABULAR.DESTINACAO}='" & cboDestinacao & "'"
    End If
    PrepararRelImovAtivEcon = True
End Function
Private Sub cmd_Click(Index As Integer)
    
    Select Case cmd(Index).Caption
        Case "&Buscar"
            Select Case OpcaoRelatorio
                Case Listagem, Aforamento
                    BuscarListagem
                    
                    
                Case Ficha
                    BuscarFicha
                    
                    
                Case ImovAtvEcon
                    BuscarImovAtvEcon
                    
            End Select
        Case "Sai&r"
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub


Private Sub cmdImprimir_Click()
    On Error GoTo Trata
    
    Screen.MousePointer = 11
        With Rpt
        Select Case cboRelatorio.ListIndex
            Case 0 ' Listagem dados cadastrais
                If PrepararRelListagem Then
                    If .DefinirArquivo(Bdados, App.Path & "\TListagemImoveis.rpt") Then
                        .Selecao = SelecaoRpt
                        .Arvore = False
                        .Visualizar
                    End If
                    DoEvents
                End If
            Case 1 ' Listagem dados Aforamento
                If PrepararRelListagem Then
                    If .DefinirArquivo(Bdados, App.Path & "\TListagemAforamento.rpt") Then
                        .Selecao = SelecaoRpt
                        .Arvore = False
                        .Visualizar
                    End If
                    DoEvents
                End If
            Case 2 ' Ficha Cadastral
                If PrepararRelFicha Then
                    If .DefinirArquivo(Bdados, App.Path & "\TCIU201.rpt") Then
                        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TCIU201", Aplicacoes.Usuario
                        .Selecao = SelecaoRpt
                        .Titulo = "Ficha Cadastral"
                        .Arvore = False
                        .Visualizar
                    End If
                    DoEvents
                End If
            
            
            Case 3 'Imoveis c/ Ativ. Economica
                If PrepararRelImovAtivEcon Then
                    If .DefinirArquivo(Bdados, App.Path & "\TImovAtvEcon.rpt") Then
                        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TImovAtvEcon", Aplicacoes.Usuario
                        .Selecao = SelecaoRpt
                        If Trim$(txtLoteamento) <> "" Then
                            .Formulas "Setor", Format(txtLoteamento, "00")
                        End If
                        .Titulo = "Imoveis com Ativ. Economica"
                        .Arvore = False
                        .Visualizar
                    End If
                    DoEvents
                End If
            Case Else
                Avisa "Selecione uma opção de impressão."
                cboRelatorio.SetFocus
        End Select
        End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
Trata:
    Screen.MousePointer = 0
    Erro Err.Description
End Sub

Private Sub cmdLimpar_Click()
    SelecaoRpt = ""
    
    Edita.LimpaCampos Me
    grid.Preencher Bdados, ""
    cboRelatorio.ListIndex = -1
    cboRelatorio.SetFocus
End Sub

Private Sub Form_Activate()
    cboRelatorio.SetFocus
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
    
    cboLogr.Preencher Bdados, "Select DISTINCT(tlg_nome),tlg_cod_logradouro From Tab_Logradouro where tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio
    cboTipoLogr.Preencher Bdados, "Select DISTINCT(ttl_nome),TTL_COD_TIP_LOGR From Tab_Tipo_Logr"
    cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
    
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

Private Sub grid_DblClick()
If grid.ListItems.Count > 0 Then
    TCIM201.Tag = grid.SelectedItem
    TCIM201.Show
End If
End Sub

Private Sub txtAnoAq_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtCodLogr_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_Validate(Cancel As Boolean)
Dim Sql As String
    If Trim(txtCodLogr) <> "" Then
        If Bdados.AbreTabela("SELECT tlg_ttl_cod_tip_logr,tlg_cod_logradouro FROM TAB_LOGRADOURO WHERE  tlg_cod_logradouro = '" & txtCodLogr & "'") Then
            cboTipoLogr.SetarLinha Bdados.Tabela(0), 1
            cboLogr.SetarLinha Bdados.Tabela(1), 1
            
            Sql = "SELECT TBA_NOME, TBA_COD_BAIRRO From TAB_BAIRRO WHERE TBA_COD_BAIRRO IN " & _
                " (SELECT DISTINCT TTC_TBA_COD_BAIRRO FROM TAB_TRECHO WHERE TTC_TLG_COD_LOGRADOURO = " & _
                txtCodLogr & ") AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio
            cboBairro.Preencher Bdados, Sql
        Else
            cboTipoLogr.ListIndex = -1
            cboLogr.ListIndex = -1
            cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
        End If
    Else
        cboTipoLogr.ListIndex = -1
        cboLogr.ListIndex = -1
        cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro where TBA_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio
    End If

End Sub

Private Sub txtContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus()
    If Me.ActiveControl.Name = "cmdLimpar" Then Exit Sub
    txtIc = Cadastro.FormataInscricao(txtIc, InscImovel)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
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
            txtLoteamento.Enabled = True
            txtQuadra.Enabled = True
            txtIc.Enabled = True
            fraCamposFicha.Enabled = False
            txtContrib.Enabled = False
            For ContLabel = 25 To 36
                Label1(ContLabel).Enabled = False
            Next
            
        Case Ficha
            txtContrib.Enabled = True
            txtIM.Enabled = True
            txtCodLogr.Enabled = True
            cboTipoLogr.Enabled = True
            cboLogr.Enabled = True
            cboBairro.Enabled = True
            txtLoteamento.Enabled = True
            txtQuadra.Enabled = True
            txtIc.Enabled = True
            fraCamposFicha.Enabled = True
            
            For ContLabel = 25 To 36
                Label1(ContLabel).Enabled = True
            Next
            
        Case ImovAtvEcon
            txtLoteamento.Enabled = True
            txtQuadra.Enabled = True
            cboDestinacao.Enabled = True
            'cboDestinacao.Enabled = True
                
        Case Else
    End Select
End Sub



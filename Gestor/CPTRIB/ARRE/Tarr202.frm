VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TARR202 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TARR202"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdGrid 
      Height          =   3120
      Left            =   45
      TabIndex        =   8
      Top             =   2880
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   5503
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   25
      Top             =   6000
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   926
      CorFundo        =   -2147483638
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   5730
         TabIndex        =   9
         Top             =   90
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         Caption         =   "&Imprimir Resumo"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7665
         TabIndex        =   10
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8850
         TabIndex        =   11
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   24
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "Tarr202.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1380
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   720
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   2434
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
      Begin VTOcx.cboVISUAL CboImposto 
         Height          =   315
         Left            =   1980
         TabIndex        =   3
         Top             =   945
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VB.TextBox txtNumLote 
         Alignment       =   1  'Right Justify
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
         Left            =   1980
         TabIndex        =   0
         Tag             =   "TDR_TLP_COD_LOTE"
         Top             =   150
         Width           =   1605
      End
      Begin VB.ComboBox cboAgente 
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
         ItemData        =   "Tarr202.frx":2123
         Left            =   1980
         List            =   "Tarr202.frx":2125
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "TLP_TAR_COD_AGENTE"
         Top             =   540
         Width           =   4185
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   14
         Top             =   600
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   423
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
         Caption         =   "Agente Arrecadador"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   150
         TabIndex        =   15
         Top             =   150
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   476
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
         Caption         =   "Nº do Lote:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   1005
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   423
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
         Caption         =   "Tributo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   1
         Left            =   6195
         TabIndex        =   26
         Top             =   615
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   318
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
         Caption         =   "Ocorrencia"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cboVISUAL CboOcorrencia 
         Height          =   315
         Left            =   7200
         TabIndex        =   2
         Top             =   555
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   705
      Index           =   3
      Left            =   45
      TabIndex        =   17
      Top             =   2115
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   1244
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
      Caption         =   "Período do Pagamento"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtDtPago1 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         TabIndex        =   4
         Tag             =   "TLP_DATA_ARRECADACAO"
         Top             =   300
         Width           =   1185
      End
      Begin VB.TextBox txtDtPago2 
         Alignment       =   1  'Right Justify
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
         Left            =   3810
         TabIndex        =   5
         Tag             =   "TLP_DATA_ARRECADACAO"
         Top             =   300
         Width           =   1095
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   8
         Left            =   150
         TabIndex        =   18
         Top             =   300
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   476
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
         Caption         =   "Data Inicial"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   9
         Left            =   2610
         TabIndex        =   19
         Top             =   300
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   476
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
         Caption         =   "Data Final"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   705
      Index           =   5
      Left            =   5115
      TabIndex        =   20
      Top             =   2115
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   1244
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
      Caption         =   "Período da Recepção"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.TextBox txtDtEntrada2 
         Alignment       =   1  'Right Justify
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
         Left            =   3780
         TabIndex        =   7
         Tag             =   "tdr_data_entrada"
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtDtEntrada1 
         Alignment       =   1  'Right Justify
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
         Left            =   1110
         TabIndex        =   6
         Tag             =   "tdr_data_entrada"
         Top             =   285
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   10
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   476
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
         Caption         =   "Data Inicial"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   11
         Left            =   2610
         TabIndex        =   22
         Top             =   300
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   476
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
         Caption         =   "Data Final"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   720
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   1138
      Icone           =   "Tarr202.frx":2127
   End
End
Attribute VB_Name = "TARR202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim CodImposto As String
Dim NumAgente  As Double
Dim NumLote As Double


Private Sub cboImposto_Click()
    CodImposto = BuscaCodigo("SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_NOME_IMPOSTO = '" & CboImposto.Text & "'")
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    Dim SELECAO As String
    Dim pos As Integer
    Screen.MousePointer = 11
    With Rpt
        SELECAO = " {TAB_LOG_OCORRENCIA.TLA_TLP_COD_LOTE} <> 0 "
        If txtNumLote <> "" Then
            SELECAO = SELECAO & " and {TAB_LOG_OCORRENCIA.TLA_TLP_COD_LOTE} = " & txtNumLote
        End If
        
        If cboAgente.ListIndex > 0 Then
            SELECAO = SELECAO & " and {TAB_AGENTE_ARRECADADOR.tar_nome_agente} = '" & cboAgente & "'"
        End If
        If CboImposto.ListIndex > -1 Then
            pos = InStr(CboImposto, "#")
            SELECAO = SELECAO & " and {Tab_Imposto.tip_sigla_imposto} = '" & Trim(Left(CboImposto, pos - 1)) & "'"
        End If
        If txtDtPago1 <> "" And txtDtPago2 <> "" Then
            SELECAO = SELECAO & " and {TAB_LOG_OCORRENCIA.TLA_DATA_ARRECADACAO} >= #" & txtDtPago1 & "# and {TAB_LOG_OCORRENCIA.TLA_DATA_ARRECADACAO} <= #" & txtDtPago2 & "#"
        ElseIf txtDtPago1 <> "" And txtDtPago2 = "" Then
            SELECAO = SELECAO & " and {TAB_LOG_OCORRENCIA.TLA_DATA_ARRECADACAO} >= #" & txtDtPago1 & "# and {TAB_LOG_OCORRENCIA.TLA_DATA_ARRECADACAO} <= #" & txtDtPago1 & "#"
        End If
        If txtDtEntrada1 <> "" And txtDtEntrada2 <> "" Then
            SELECAO = SELECAO & " and {TAB_LOG_OCORRENCIA.TLA_DATA_RECEPCAO} >= #" & txtDtEntrada1 & "# and {TAB_LOG_OCORRENCIA.TLA_DATA_RECEPCAO} <= #" & txtDtEntrada2 & "#"
        ElseIf txtDtEntrada1 <> "" And txtDtEntrada2 = "" Then
            SELECAO = SELECAO & " and {TAB_LOG_OCORRENCIA.TLA_DATA_RECEPCAO} >= #" & txtDtEntrada1 & "# and {TAB_LOG_OCORRENCIA.TLA_DATA_RECEPCAO} <= #" & txtDtEntrada1 & "#"
        End If
        If cboOcorrencia.ListIndex > -1 Then
            SELECAO = SELECAO & " and {TAB_LOG_OCORRENCIA.TLA_OCORRENCIA} = " & cboOcorrencia.Coluna(1).Valor
        End If
        If Not .DefinirArquivo(Bdados, App.Path & "\TRelLogOcorrencia.rpt") Then Exit Sub
        .SELECAO = SELECAO
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Arvore = False
        .Visualizar
    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Sql As String
    Dim pos As Integer
    
    Screen.MousePointer = vbHourglass
    Sql = "Select * from vis_log_ocorrencia where 1 =1 "
    
    If txtNumLote <> "" Then
        Sql = Sql & " and Lote = '" & txtNumLote & "'"
    End If
    
    If cboAgente.ListIndex > 0 Then
        Sql = Sql & " and Banco like '%" & cboAgente & "%'"
    End If
    If CboImposto.ListIndex > -1 Then
        pos = InStr(CboImposto, "#")
        Sql = Sql & " and Imposto like '%" & Trim(Left(CboImposto, pos - 1)) & "%'"
    End If
    If txtDtPago1 <> "" And txtDtPago2 <> "" Then
        Sql = Sql & " and Arrecadação >= " & Bdados.Converte(txtDtPago1, TCDataHora) & " and [Arrecadação] <= " & Bdados.Converte(txtDtPago2, TCDataHora)
    ElseIf txtDtPago1 <> "" And txtDtPago2 = "" Then
        Sql = Sql & " and Arrecadação >= " & Bdados.Converte(txtDtPago1, TCDataHora) & " and [Arrecadação] <= " & Bdados.Converte(txtDtPago1, TCDataHora)
    End If
    If txtDtEntrada1 <> "" And txtDtEntrada2 <> "" Then
        Sql = Sql & " and Recepção >= " & Bdados.Converte(txtDtEntrada1, TCDataHora) & " and Recepção <= " & Bdados.Converte(txtDtEntrada2, TCDataHora)
    ElseIf txtDtEntrada1 <> "" And txtDtEntrada2 = "" Then
        Sql = Sql & " and Recepção >= " & Bdados.Converte(txtDtEntrada1, TCDataHora) & " and Recepção <= " & Bdados.Converte(txtDtEntrada1, TCDataHora)
    End If
    If cboOcorrencia.ListIndex > -1 Then
        Sql = Sql & " and Ocorrência like '%" & cboOcorrencia.Text & "%'"
    End If
    grdGrid.Preencher Bdados, Sql
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim rs As VSRecordset
    Dim Obrig As New Obrigacao
    Dim pos As Integer
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboAgente.Clear
    
    cboOcorrencia.PreencherGeral Bdados, "TIPO OCORRENCIA"
    AtualizaCombo Bdados, cboAgente, "Select tar_nome_agente from tab_agente_arrecadador where tar_ativo =0"
    Obrig.PreencheComboTributo CboImposto, False
    CboImposto.AddItem " "
    cboAgente.AddItem " "
    DoEvents
End Sub


Private Sub txtDtArrecada_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtDtEntrada1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtEntrada1_LostFocus()
    txtDtEntrada1 = Edita.FormataTexto(txtDtEntrada1, Data)
End Sub

Private Sub txtDtEntrada2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtEntrada2_LostFocus()
    txtDtEntrada2 = Edita.FormataTexto(txtDtEntrada2, Data)
End Sub

Private Sub txtDtPago1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtPago1_LostFocus()
    txtDtPago1 = Edita.FormataTexto(txtDtPago1, Data)
End Sub

Private Sub txtDtPago2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtDtPago2_LostFocus()
    txtDtPago2 = Edita.FormataTexto(txtDtPago2, Data)
End Sub

Private Sub txtDtRecep_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub




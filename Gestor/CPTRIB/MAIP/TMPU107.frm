VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU107 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame fra 
      Height          =   1995
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   720
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   3519
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
      Begin VB.ComboBox cboBairro 
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
         ItemData        =   "TMPU107.frx":0000
         Left            =   300
         List            =   "TMPU107.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Bairro"
         Top             =   425
         Width           =   2865
      End
      Begin VB.TextBox txtTestada 
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
         Left            =   300
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Testada"
         Top             =   1410
         Width           =   1215
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   300
         TabIndex        =   10
         Top             =   1140
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   397
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "Testada Padrão"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   12
         Top             =   150
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   "Bairro"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   1005
         Index           =   1
         Left            =   5040
         TabIndex        =   13
         Top             =   840
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   1773
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
         Caption         =   "Valor do M² (R$)"
         Begin VB.TextBox txtValorMaximo 
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
            Left            =   225
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Valor"
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox txtValorMinimo 
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
            Left            =   1650
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Valor"
            Top             =   570
            Width           =   1215
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   9
            Left            =   225
            TabIndex        =   14
            Top             =   300
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   397
            _Version        =   196610
            CaptionStyle    =   1
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
            Caption         =   "Máximo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   3
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   11
            Left            =   1650
            TabIndex        =   15
            Top             =   300
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   397
            _Version        =   196610
            CaptionStyle    =   1
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
            Caption         =   "Mínimo"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   3
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   1005
         Index           =   2
         Left            =   1890
         TabIndex        =   16
         Top             =   840
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   1773
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
         Caption         =   "Profundidade (m)"
         Begin VB.TextBox txtProfMaximo 
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
            Left            =   240
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Valor"
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox txtProfMinimo 
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
            Left            =   1665
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Profundidade Máxima"
            Top             =   570
            Width           =   1215
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   1
            Left            =   255
            TabIndex        =   17
            Top             =   300
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   397
            _Version        =   196610
            CaptionStyle    =   1
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
            Caption         =   "Máxima"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   3
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   225
            Index           =   2
            Left            =   1680
            TabIndex        =   18
            Top             =   300
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
            _Version        =   196610
            CaptionStyle    =   1
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
            Caption         =   "Mínima"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   3
            RoundedCorners  =   0   'False
         End
      End
   End
   Begin MSComctlLib.ListView lstParadigma 
      Height          =   2985
      Left            =   60
      TabIndex        =   8
      Top             =   2775
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   5265
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   11
      Top             =   -420
      Width           =   375
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   7170
      TabIndex        =   7
      Top             =   5820
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   5985
      TabIndex        =   6
      Top             =   5820
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   1138
      Icone           =   "TMPU107.frx":0004
   End
End
Attribute VB_Name = "TMPU107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBairro_Click()
    If Bdados.AbreTabela("SELECT TBA_COD_BAIRRO FROM TAB_BAIRRO WHERE TBA_NOME='" & cboBairro & "' AND TBA_TMU_COD_MUNICIPIO=" & Aplicacoes.Codigo_Municipio) Then
        'Util.MontaGrid Bdados,lstParadigma, "SELECT TLO_TESTADA Testada, TLO_PROF_MAXIMA [Prof Maxima], TLO_PROF_MINIMA [Prof Minima], TLO_VALOR_MAX [Vl Maximo], TLO_VALOR_MIN [Vl Minimo] From TAB_LOTE_PADRAO, TAB_BAIRRO WHERE TLO_TMU_COD_MUNICIPIO =" & Aplicacoes.Codigo_Municipio & " AND TLO_TBA_COD_BAIRRO = " & Bdados.Tabela(0)
        Util.MontaGrid Bdados, lstParadigma, "SELECT TBA_NOME Bairro, TLO_TESTADA Testada, TLO_PROF_MAXIMA [Prof Maxima], TLO_PROF_MINIMA [Prof Minima], TLO_VALOR_MAX [Vl Maximo], TLO_VALOR_MIN [Vl Minimo] From TAB_LOTE_PADRAO, TAB_BAIRRO WHERE TBA_COD_BAIRRO = TLO_TBA_COD_BAIRRO AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " AND TLO_TBA_COD_BAIRRO = " & Bdados.Tabela(0)
        lstParadigma.AllowColumnReorder = False
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
Dim Valores As String
Dim Campos As String
Dim Condicao As String

    If Bdados.AbreTabela("SELECT TBA_COD_BAIRRO FROM TAB_BAIRRO WHERE TBA_NOME='" & cboBairro & "' AND TBA_TMU_COD_MUNICIPIO=" & Aplicacoes.Codigo_Municipio) Then
        Valores = Bdados.PreparaValor(Aplicacoes.Codigo_Municipio, Bdados.Tabela(0), txtTestada, txtProfMinimo, txtProfMaximo, txtValorMaximo, txtValorMinimo)
        Campos = "TLO_TMU_COD_MUNICIPIO,TLO_TBA_COD_BAIRRO,TLO_TESTADA,TLO_PROF_MINIMA,TLO_PROF_MAXIMA,TLO_VALOR_MAX,TLO_VALOR_MIN"
        Condicao = "TLO_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " AND TLO_TBA_COD_BAIRRO = " & Bdados.Tabela(0)
        If Bdados.GravaDados("TAB_LOTE_PADRAO", Valores, Campos, Condicao) Then
            Util.Informa "Dados guardados na tabela de paradigma."
            txtTestada = ""
            txtProfMaximo = ""
            txtProfMinimo = ""
            txtValorMaximo = ""
            txtValorMinimo = ""
            Util.MontaGrid Bdados, lstParadigma, "SELECT TBA_NOME Bairro, TLO_TESTADA Testada, TLO_PROF_MAXIMA [Prof Maxima], TLO_PROF_MINIMA [Prof Minima], TLO_VALOR_MAX [Vl Maximo], TLO_VALOR_MIN [Vl Minimo] From TAB_LOTE_PADRAO, TAB_BAIRRO WHERE TBA_COD_BAIRRO = TLO_TBA_COD_BAIRRO AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " AND TLO_TBA_COD_BAIRRO = " & Bdados.Tabela(0)
        Else
            Util.Informa "Não foi possível salvar os dados na tabela de paradigma."
        End If
    End If
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    Edita.AtualizaCombo Bdados, cboBairro, "SELECT TBA_NOME FROM TAB_BAIRRO WHERE TBA_TMU_COD_MUNICIPIO=" & Aplicacoes.Codigo_Municipio
    Util.MontaGrid Bdados, lstParadigma, "SELECT TBA_NOME Bairro, TLO_TESTADA Testada, TLO_PROF_MAXIMA [Prof Maxima], TLO_PROF_MINIMA [Prof Minima], TLO_VALOR_MAX [Vl Maximo], TLO_VALOR_MIN [Vl Minimo] From TAB_LOTE_PADRAO, TAB_BAIRRO WHERE TBA_COD_BAIRRO = TLO_TBA_COD_BAIRRO AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio
    lstParadigma.AllowColumnReorder = False
End Sub

Private Sub lstParadigma_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid Me.ActiveControl, ColumnHeader
End Sub

Private Sub lstParadigma_DblClick()
    cboBairro = lstParadigma.SelectedItem
    txtTestada = lstParadigma.SelectedItem.SubItems(1)
    txtProfMaximo = lstParadigma.SelectedItem.SubItems(2)
    txtProfMinimo = lstParadigma.SelectedItem.SubItems(3)
    txtValorMaximo = lstParadigma.SelectedItem.SubItems(4)
    txtValorMinimo = lstParadigma.SelectedItem.SubItems(5)
End Sub

Private Sub txtProfMaximo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtProfMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtTestada_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtTestada_Validate(Cancel As Boolean)
    txtTestada = Edita.FormataTexto(txtTestada, Monetario)
End Sub

Private Sub txtValorMaximo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValorMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

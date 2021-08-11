VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR301 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR301.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   195
      Left            =   1890
      TabIndex        =   7
      Top             =   -300
      Width           =   375
   End
   Begin Threed.SSFrame fra 
      Height          =   3825
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   645
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   6747
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
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   1680
         TabIndex        =   20
         Top             =   60
         Width           =   4395
         Begin VTOcx.cboVISUAL cboNatueza 
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   150
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   556
            Caption         =   "Natureza Jurídica"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
      End
      Begin VB.TextBox txtJuros 
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
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Juros"
         Top             =   2100
         Width           =   855
      End
      Begin VB.TextBox txtMaxCotas 
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
         Left            =   4740
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Máximo Cotas"
         Top             =   2100
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   6
         Left            =   2940
         TabIndex        =   9
         Top             =   2145
         Width           =   1755
         _ExtentX        =   3096
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
         Caption         =   "NO. Maximo de Cotas"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   705
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2430
         Width           =   5925
         _ExtentX        =   10451
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
         Caption         =   "Valor da Multa Mora (%)"
         ShadowStyle     =   1
         Begin VB.TextBox txtValMaxMulta 
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
            Left            =   4560
            MaxLength       =   8
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   270
            Width           =   1185
         End
         Begin VB.TextBox txtValMinMulta 
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
            MaxLength       =   8
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   270
            Width           =   1065
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   11
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   1155
            _ExtentX        =   2037
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
            Caption         =   "Valor Mínimo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   13
            Left            =   3270
            TabIndex        =   12
            Top             =   300
            Width           =   1185
            _ExtentX        =   2090
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
            Caption         =   "Valor Máximo:"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   705
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "Valor Mínimo para parcelamento (R$):"
         ShadowStyle     =   1
         Begin VB.TextBox txtValMinBase 
            Alignment       =   2  'Center
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
            Left            =   120
            TabIndex        =   0
            Tag             =   "Valor Mínimo da Dívida"
            Top             =   270
            Width           =   2655
         End
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   15
         Left            =   180
         TabIndex        =   14
         Top             =   2145
         Width           =   1425
         _ExtentX        =   2514
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
         Caption         =   "Taxa de Juros(%)"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSFrame fra 
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1350
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1138
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
         Caption         =   "Desconto na Dívida Parcelada(%):"
         ShadowStyle     =   1
         Begin VB.TextBox TxtMinEntrada 
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
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   1
            Tag             =   "Mínimo Entrada"
            Top             =   240
            Width           =   1065
         End
         Begin VB.TextBox txtDesconto 
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
            Left            =   4620
            MaxLength       =   8
            TabIndex        =   2
            Tag             =   "Redução"
            Top             =   240
            Width           =   1185
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   262
            Width           =   1290
            _ExtentX        =   2275
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
            Caption         =   "Mínimo Entrada"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
         Begin Threed.SSPanel lbl 
            Height          =   270
            Index           =   1
            Left            =   3615
            TabIndex        =   17
            Top             =   262
            Width           =   975
            _ExtentX        =   1720
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
            Caption         =   "Redução(%)"
            BorderWidth     =   1
            BevelOuter      =   0
            AutoSize        =   1
            Alignment       =   0
            RoundedCorners  =   0   'False
         End
      End
      Begin Threed.SSFrame fra 
         Height          =   705
         Index           =   2
         Left            =   3090
         TabIndex        =   22
         Top             =   630
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "Número Mínimo de Cotas"
         ShadowStyle     =   1
         Begin VB.TextBox txtValMinCota 
            Alignment       =   2  'Center
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
            Left            =   90
            MaxLength       =   8
            TabIndex        =   23
            Tag             =   "Valor Mínimom Cotas"
            Top             =   270
            Width           =   2775
         End
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   4950
         TabIndex        =   24
         Top             =   3210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   3750
         TabIndex        =   25
         Top             =   3210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5430
      Top             =   1320
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TPAR301.frx":2123
   End
End
Attribute VB_Name = "TPAR301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboNatueza_Click()
    Seta_Natureza
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Valores As String
    Dim campos As String
    Dim Pessoa As String
    
    Dim Sql As String
    Dim rs As VSRecordset
    Select Case Index
        Case 0
            If Not Edita.CriticaCampos(Me) Then Exit Sub
            Valores = Bdados.PreparaValor(cboNatueza.Coluna(0).Valor, Bdados.Converte(txtValMinBase, TCDuplo), Bdados.Converte(txtValMinCota, TCDuplo), _
                    Bdados.Converte(txtJuros, TCDuplo), Bdados.Converte(txtMaxCotas, TCDuplo), Bdados.Converte(TxtMinEntrada, TCDuplo), Bdados.Converte(txtDesconto, TCDuplo))
            campos = "tpp_Pessoa,tpp_valor_min_base_calc,tpp_valor_min_cota,"
            campos = campos & "tpp_valor_juros,tpp_max_cotas,"
            campos = campos & "tpp_valor_minimo_entrada,tpp_reducao"
            Call Bdados.GravaDados("Tab_Parametro_Parcelamento", Valores, campos, "tpp_Pessoa = " & cboNatueza.Coluna(0).Valor)
            Call Util.Informa("Transação Completada.")
            Bdados.FechaTabela rs
            txtValMinBase.SetFocus
        Case 1
            Unload Me
    End Select
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Dim Sql As String
    Dim rs As VSRecordset
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboNatueza.Preencher Bdados, "select * from TAB_NATUREZA_JURIDICA ", 1
End Sub

Private Sub Seta_Natureza()
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT * from Tab_Parametro_Parcelamento where tpp_pessoa = " & cboNatueza.Coluna(0).Valor
    If Bdados.AbreTabela(Sql, rs) Then
        txtValMinBase = Format(rs!tpp_valor_min_base_calc, "###,###,###,##0.00")
        txtValMinCota = Format(rs!tpp_valor_min_cota, "###,###,###,##0.00")
        txtJuros = Format(rs!tpp_valor_juros, "###,###,###,##0.00")
        txtMaxCotas = Format(rs!tpp_max_cotas, "###,###,###,##0.00")
        TxtMinEntrada = Format(rs!tpp_valor_minimo_entrada, "###,###,###,##0.00")
        txtDesconto = Format(rs!tpp_reducao, "###,###,###,##0.00")
    Else
        txtValMinBase = "0,00"
        txtValMinCota = "0,00"
        txtJuros = "0,00"
        txtMaxCotas = "0,00"
        TxtMinEntrada = "0,00"
        txtDesconto = "0,00"
    End If
    Bdados.FechaTabela rs

End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtDesconto_LostFocus()
    txtDesconto = Edita.FormataTexto(txtDesconto, Monetario, True)
End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtJuros_LostFocus()
    txtJuros = Edita.FormataTexto(txtJuros, Monetario, True)
End Sub

Private Sub txtMaxCotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtMaxCotas_LostFocus()
    txtMaxCotas = Edita.FormataTexto(txtMaxCotas, Monetario, True)
End Sub

Private Sub TxtMinEntrada_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub TxtMinEntrada_LostFocus()
    TxtMinEntrada = Edita.FormataTexto(TxtMinEntrada, Monetario, True)
End Sub

Private Sub txtValMaxMulta_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValMaxMulta_LostFocus()
    txtValMaxMulta = Edita.FormataTexto(txtValMaxMulta, Monetario, True)
End Sub

Private Sub txtValMinBase_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValMinBase_LostFocus()
    txtValMinBase = Edita.FormataTexto(txtValMinBase, Monetario, True)
End Sub

Private Sub txtValMinCota_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValMinCota_LostFocus()
    txtValMinCota = Edita.FormataTexto(txtValMinCota, Monetario, True)
End Sub

Private Sub txtValMinMulta_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtValMinMulta_LostFocus()
    txtValMinMulta = Edita.FormataTexto(txtValMinMulta, Monetario, True)
End Sub

VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TPAR102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleMode       =   0  'User
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   32
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TPAR102.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   25
      Top             =   6690
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   953
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   3585
         TabIndex        =   7
         Top             =   105
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "&Gerar Documento"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   5580
         TabIndex        =   8
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1560
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   705
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   2752
      Altura          =   1905
      Caption         =   " Parcelamento"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VB.TextBox txtPeriodo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5100
         TabIndex        =   12
         Tag             =   "Exercicio"
         Top             =   1110
         Width           =   1065
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   0
         Left            =   5130
         TabIndex        =   30
         Top             =   900
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
         PictureMaskColor=   -2147483644
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
         Caption         =   "Exercício"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtCodTrib 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5130
         TabIndex        =   10
         Tag             =   "Codigo Tributo"
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox txtInsc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   11
         Tag             =   "IM"
         Top             =   1140
         Width           =   2265
      End
      Begin VB.TextBox txtParc 
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
         Left            =   150
         MaxLength       =   14
         TabIndex        =   0
         Tag             =   "NO. DAM"
         Top             =   570
         Width           =   2235
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   7
         Left            =   150
         TabIndex        =   21
         Top             =   345
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
         PictureMaskColor=   -2147483644
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
         Caption         =   "Nº PARCELAMENTO"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   9
         Left            =   150
         TabIndex        =   20
         Top             =   915
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
         PictureMaskColor=   -2147483644
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
         Caption         =   "Inscricão"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   3
         Left            =   5130
         TabIndex        =   19
         Top             =   345
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
         PictureMaskColor=   -2147483644
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
         Caption         =   "Cód. Tributo"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtEnderecoContrib 
      Enabled         =   0   'False
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
      Left            =   660
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5895
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4560
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1138
      Icone           =   "TPAR102.frx":2123
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1020
      Left            =   75
      TabIndex        =   18
      Top             =   3345
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   1799
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   15
         Left            =   5085
         TabIndex        =   24
         Top             =   330
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
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
         Caption         =   "Total a Recolher"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   14
         Left            =   1665
         TabIndex        =   23
         Top             =   285
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
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
         Caption         =   "Juros"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   13
         Left            =   105
         TabIndex        =   22
         Top             =   300
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
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
         Caption         =   "Valor Principal"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtImposto 
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
         Left            =   120
         TabIndex        =   4
         Tag             =   "Imposto"
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txtJuros 
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
         Left            =   1650
         TabIndex        =   5
         Tag             =   "Juros"
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtTotalImposto 
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
         Left            =   5100
         TabIndex        =   6
         Tag             =   "Total"
         Top             =   615
         Width           =   1185
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   960
      Index           =   1
      Left            =   60
      TabIndex        =   26
      Top             =   2310
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   1693
      Altura          =   1905
      Caption         =   " Parcela"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VB.TextBox txtCodParc 
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
         Left            =   2370
         MaxLength       =   14
         TabIndex        =   2
         Tag             =   "NO. DAM"
         Top             =   555
         Width           =   1605
      End
      Begin VB.TextBox txtDtVenc 
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
         Left            =   5115
         TabIndex        =   3
         Tag             =   "Data Vencimento"
         Top             =   555
         Width           =   1335
      End
      Begin VB.TextBox txtParcela 
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
         Left            =   165
         TabIndex        =   1
         Tag             =   "Parcela"
         Top             =   555
         Width           =   945
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   17
         Left            =   165
         TabIndex        =   27
         Top             =   300
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
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
         Caption         =   "Nº Parcela"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   1
         Left            =   5145
         TabIndex        =   28
         Top             =   315
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
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
         Caption         =   "Data Vencimento"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   270
         Index           =   2
         Left            =   2370
         TabIndex        =   29
         Top             =   300
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   476
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   0
         BackColor       =   -2147483644
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
         Caption         =   "Cod. Parcela"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL grdParcAntigas 
      Height          =   2205
      Left            =   60
      TabIndex        =   31
      Top             =   4410
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   3889
      Caption         =   "Parcelas Originais"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TPAR102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim CodImposto As String
Dim Exercicio As String
Dim Conta As New ContaCorrente
Dim Aliquota As Double
'Variaveis para o Report
Dim InscMuni As String
Dim RazaoSocial As String
Dim Documento As String
Dim Localizacao As String
Dim Data_Vencimento As String
Dim Codigo_Imovel As String
Dim Valor_Imposto As String
Dim CPFCNPJ As String
Dim Endereco As String
Dim Bairro As String
Dim Cod_Atividade As String
Dim Cod_Cidade As String
Dim Cep As String
Dim Uf As String
Dim Cod_Tributo As String
Dim Juro As String
Dim Multa As String
Dim TotalImposto As String
Dim TaxaServico As Double
Dim BaseDeCalculo As String
Dim VetLinhas(0 To 5) As String
Dim Linhas As Byte
Dim ObsAux As String
Dim NomeImposto As String
Dim TributoTaxa As Boolean
Dim TributoTaxaFixa As Double
Dim Tributo As Double
Dim Alvara As Double
Dim PosTraco As Byte
Dim TSU As Double
Dim AreaConstruida As Double
Dim AreaTotal As Double
Dim ValorTerreno As Double
Dim Valoredific As Double
Dim Zona As Integer
Dim ValorMetro As Double
Dim TaxaParcela As Double
Dim Desconto As String
Dim Reducao As String
Dim DtGeracao As String
Dim CodPagamento As Double
Private Sub cmd_Click(Index As Integer)
    Dim a As Integer
    Dim Valores As String
    Dim Campos As String
    Dim ValorImposto As Double
    Dim RsCob As VSRecordset
    Dim rs As VSRecordset
    Dim sql As String
    Dim SqlParc As String
    Dim Cobranca As New VSCobranca
    
    Screen.MousePointer = 11
    
    Select Case cmd(Index).Caption
        Case "&Gerar Documento"
            If Not Edita.CriticaCampos(Me) Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            Valores = Bdados.PreparaValor(txtParc, txtParcela, Bdados.Converte(Format(Date, "dd/mm/yyyy"), TCDataHora), _
                        txtDtVenc, txtImposto, txtJuros, 1, txtCodParc)
            
            Campos = "TCO_TPA_COD_PARCELAMENTO,TCO_NUM_PARCELA,TCO_DATA_GERACAO,TCO_DATA_VENCIMENTO,TCO_VALOR_PARCELA" & _
                    ",TCO_VALOR_JUROS,TCO_STATUS_OBRIGACAO_PARCELA,TCO_TOC_NUM_OBRIGACAO"
            Bdados.GravaDados "TAB_COTAS_OBRIGACAO", Valores, Campos, "TCO_TOC_NUM_OBRIGACAO=" & txtCodParc
            Informa "Parcela gerada."
            txtParcela.SetFocus
        Case "Sai&r"
           Unload Me
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click(Index As Integer)
    Edita.LimpaCampos Me
    txtParc.SetFocus
End Sub

Private Sub Form_Load()
            
    Dim Controle As Control
    Dim i As Byte
    
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub


Private Sub txtContribuinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCodReceita_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    If Trim(txtCodTrib) = "" Then Exit Sub
    sql = "Select tip_nome_imposto from tab_imposto where tip_cod_imposto ='" & txtCodTrib & "'"
    If Not Bdados.AbreTabela(sql, rs) Then
        Avisa "Código de receita inválido."
        txtCodTrib.SetFocus
        Exit Sub
    End If
End Sub

Private Sub grdParcAntigas_DblClick()
    
    txtCodParc = grdParcAntigas.SelectedItem
    txtParcela = grdParcAntigas.SelectedItem.SubItems(1)
    txtDtVenc = grdParcAntigas.SelectedItem.SubItems(2)
    txtImposto = grdParcAntigas.SelectedItem.SubItems(3)
    txtJuros = grdParcAntigas.SelectedItem.SubItems(4)
    txtTotalImposto = CDbl(Nvl(grdParcAntigas.SelectedItem.SubItems(3), 0)) + CDbl(Nvl(grdParcAntigas.SelectedItem.SubItems(4), 0))
End Sub

Private Sub txtCodParc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodParc_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    Dim i As Byte
    If Trim(txtCodParc) = "" Then Exit Sub
    sql = "Select TCO_NUM_PARCELA,TCO_DATA_VENCIMENTO,TCO_NUM_PARCELA,TCO_DATA_VENCIMENTO,TCO_VALOR_PARCELA,TCO_VALOR_JUROS from TAB_COTAS_OBRIGACAO " & _
            " where TCO_TOC_NUM_OBRIGACAO =" & txtCodParc
    If Bdados.AbreTabela(sql, rs) Then
        txtParcela = "" & rs!TCO_NUM_PARCELA
        txtDtVenc = "" & rs!TCO_DATA_VENCIMENTO
        txtDtVenc = Format("" & rs!TCO_DATA_VENCIMENTO, "dd/mm/yyyy")
        txtImposto = Format("" & rs!TCO_VALOR_PARCELA, Const_Monetario)
        txtJuros = Format("" & rs!TCO_VALOR_JUROS, Const_Monetario)
        txtTotalImposto = Format(CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtJuros, 0)), Const_Monetario)
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtDtVenc_LostFocus()
    txtDtVenc = Edita.FormataTexto(txtDtVenc, Data)
End Sub

Private Sub txtInsc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtInsc_LostFocus()
    On Error GoTo trata
    txtInsc = Imposto.FormataInscricao(txtInsc, InscContrib)
trata:
    If Err.Number = 3265 Then
        Resume Next
    End If
End Sub

Private Sub txtImposto_LostFocus()
    txtImposto = Format(txtImposto, Const_Monetario)
End Sub

Private Sub txtJuros_LostFocus()
    txtJuros = Format(txtJuros, Const_Monetario)
End Sub

Private Sub txtParc_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    Dim i As Byte
    If Trim(txtParc) = "" Then Exit Sub
    sql = "Select TPA_INSCRICAO,TPA_TIP_COD_IMPOSTO,TPA_PERIODO,TPA_TCI_IM,TPA_TIM_IC from Tab_Parcelamento where TPA_NUM_PARCELAMENTO =" & txtParc
    If Not Bdados.AbreTabela(sql, rs) Then
        Informa "Parcelamento inexistente."
        txtParc.SetFocus
        Bdados.FechaTabela rs
    Else
        txtInsc = Trim("" & rs!TPA_INSCRICAO)
        txtInsc = IIf(Trim(txtInsc) = "", IIf(Trim("" & rs!TPA_TCI_IM) = "", "" & rs!TPA_TIM_IC, "" & rs!TPA_TCI_IM), txtInsc)
        txtCodTrib = "" & rs!TPA_TIP_COD_IMPOSTO
        txtPeriodo = "" & rs!TPA_PERIODO
        sql = "SELECT tgt_cod_pagamento as [Cod Parcela],tgt_parcela as Cota,tgt_data_vencimento as Vencimento,tgt_valor_tributo as Valor,tgt_valor_juros as Juros FROM TAB_GERACAO_TRIBUTO WHERE tgt_tpa_num_parcelamento =" & txtParc & " order by tgt_parcela asc"
        grdParcAntigas.Preencher Bdados, sql, 1200, 800, 1200, 1000, 800
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtParcela_LostFocus()
    Dim sql As String
    Dim rs As VSRecordset
    Dim i As Byte
    If Trim(txtParc) = "" Or Trim(txtParcela) = "" Then Exit Sub
    sql = "Select TCO_TOC_NUM_OBRIGACAO,TCO_DATA_VENCIMENTO,TCO_NUM_PARCELA,TCO_DATA_VENCIMENTO,TCO_VALOR_PARCELA,TCO_VALOR_JUROS from TAB_COTAS_OBRIGACAO " & _
            " where TCO_NUM_PARCELA =" & Trim(txtParcela) & " and TCO_TPA_COD_PARCELAMENTO =" & Trim(txtParc)
    If Bdados.AbreTabela(sql, rs) Then
        txtCodParc = "" & rs!TCO_TOC_NUM_OBRIGACAO
        txtDtVenc = "" & rs!TCO_DATA_VENCIMENTO
        txtDtVenc = Format("" & rs!TCO_DATA_VENCIMENTO, "dd/mm/yyyy")
        txtImposto = Format("" & rs!TCO_VALOR_PARCELA, Const_Monetario)
        txtJuros = Format("" & rs!TCO_VALOR_JUROS, Const_Monetario)
        txtTotalImposto = Format(CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtJuros, 0)), Const_Monetario)
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtTotalImposto_LostFocus()
        txtTotalImposto = Format(txtTotalImposto, Const_Monetario)
End Sub

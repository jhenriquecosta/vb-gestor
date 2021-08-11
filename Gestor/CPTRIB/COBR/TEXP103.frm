VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TEXP103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuinte"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   33
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TEXP103.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Selecionar Todos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   5910
      Width           =   2655
   End
   Begin VTOcx.fraVISUAL fraVISUAL4 
      Height          =   765
      Left            =   30
      TabIndex        =   26
      Top             =   6180
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1349
      Altura          =   1905
      Caption         =   " Resultados Parciais do Parcelamento"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VB.TextBox txtTotalParc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1275
         TabIndex        =   17
         Tag             =   "Debito Total"
         Top             =   360
         Width           =   1245
      End
      Begin VB.TextBox txtValorPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4215
         TabIndex        =   18
         Tag             =   "Valor Pago"
         Top             =   345
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   7
         Left            =   2535
         TabIndex        =   30
         Top             =   375
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   423
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
         Caption         =   "Valor a Pagar:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   180
         Index           =   4
         Left            =   5595
         TabIndex        =   29
         Top             =   405
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   318
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
         Caption         =   "Débito Pendente:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   28
         Top             =   375
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   423
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
         Caption         =   "Débito Total:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtdebitoRestante 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   7665
         TabIndex        =   19
         Tag             =   "Debito Pendente"
         Top             =   345
         Width           =   1215
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1830
      Left            =   30
      TabIndex        =   24
      Top             =   660
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   3228
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   7470
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   4020
         TabIndex        =   2
         Top             =   735
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         CorFundo        =   -2147483644
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3540
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   705
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtInscricao 
         Height          =   300
         Left            =   750
         TabIndex        =   1
         Tag             =   "Inscrição Cadastral"
         Top             =   720
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   529
         Caption         =   "Inscricao"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   315
         Left            =   750
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1440
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   556
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
         Enabled         =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   9090
         TabIndex        =   10
         Top             =   1320
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   30
         TabIndex        =   27
         Top             =   420
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   423
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
         Caption         =   "Data Vencimento:"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   3
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtDtVence 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1590
         TabIndex        =   0
         Tag             =   "Data Vencimento"
         Top             =   360
         Width           =   1215
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   23
      Top             =   8145
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   7590
         TabIndex        =   12
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   6420
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   8610
         TabIndex        =   14
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9660
         TabIndex        =   15
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   21
      Top             =   -525
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   20
      Top             =   -360
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1138
      Icone           =   "TEXP103.frx":2123
   End
   Begin VTOcx.grdVISUAL lstParcelas 
      Height          =   2865
      Left            =   30
      TabIndex        =   16
      Top             =   3315
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   5054
      CorFundo        =   -2147483638
      CorTitulo       =   32768
      CorCaption      =   -2147483629
      CorDica         =   -2147483627
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      Ordenavel       =   0   'False
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1110
      Left            =   30
      TabIndex        =   25
      Top             =   6975
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1958
      Altura          =   1905
      Caption         =   " Detalhes"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin Threed.SSPanel lbl 
         Height          =   210
         Index           =   19
         Left            =   60
         TabIndex        =   31
         Top             =   375
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   370
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
         Caption         =   "Observações :"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   1290
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   375
         Width           =   9210
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL3 
      Height          =   765
      Left            =   30
      TabIndex        =   35
      Top             =   2520
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   1349
      Altura          =   1905
      Caption         =   " Descontos(%)"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtMulta 
         Height          =   300
         Left            =   5340
         TabIndex        =   5
         Top             =   330
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         Caption         =   "Multa"
         Text            =   ""
         Restricao       =   3
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtAtualizacao 
         Height          =   300
         Left            =   2880
         TabIndex        =   4
         Top             =   330
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   529
         Caption         =   "Atualização"
         Text            =   ""
         Restricao       =   3
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtJuros 
         Height          =   300
         Left            =   7650
         TabIndex        =   6
         Top             =   330
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         Caption         =   "Juros"
         Text            =   ""
         Restricao       =   3
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtOriginal 
         Height          =   300
         Left            =   360
         TabIndex        =   3
         Top             =   330
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         Caption         =   "Valor Original"
         Text            =   ""
         Restricao       =   3
         CorFundo        =   -2147483644
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
   End
End
Attribute VB_Name = "TEXP103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim MaxCotas As Byte
Dim CodImp As String
Dim Cgc As String
Dim EnderecoContrib As String
Dim CodPagamento As String
Dim ListaDocs As String
Dim Marcou As Boolean

Dim Valor As Double
Dim Multa As Double
Dim Juros As Double
Dim Saldo As Double
Dim Atualizacao As Double
Dim Desconto As Double

Dim Periodo As String
Dim Tributo As String

Private Sub Pega_Dados()
    Dim i As Integer
    Dim Base As String
    Dim Base2 As String
    Tributo = ""
    Periodo = ""
    
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked = True Then
            
            
            If Tributo = "" Then
                Tributo = "[ " & lstParcelas.ListItems(i).SubItems(2) & " ]"
            Else
                If lstParcelas.ListItems.Count - i = 1 Then
                    'Checo se é o fim da tabela, se for coloco a virgula...
                    If Base <> lstParcelas.ListItems(i).SubItems(2) Then
                        Tributo = Tributo & "e [ " & lstParcelas.ListItems(i).SubItems(2) & " ]"
                    End If
                Else
                   'If para não repetir a sigla outra vez...
                    If Base <> lstParcelas.ListItems(i).SubItems(2) Then
                        Tributo = Tributo & ", [ " & lstParcelas.ListItems(i).SubItems(2) & " ]"
                    End If
                End If
            End If
            
             If Periodo = "" Then
                    Periodo = lstParcelas.ListItems(i).SubItems(3)
                Else
                    If lstParcelas.ListItems.Count - i = 1 Then
                        'Checo se é o fim da tabela, se for coloco a virgula...
                        If Base2 <> lstParcelas.ListItems(i).SubItems(3) Then
                            'Periodo = Periodo & "e [ " & lstParcelas.ListItems(I).SubItems(3) & " ]"
                            Periodo = Periodo & "e " & IIf(Len(lstParcelas.ListItems(i).SubItems(3)) = 4, lstParcelas.ListItems(i).SubItems(3), Right(lstParcelas.ListItems(i).SubItems(3), 2) & Left(lstParcelas.ListItems(i).SubItems(3), 4))
                        End If
                    Else
                        If Base2 <> lstParcelas.ListItems(i).SubItems(3) Then
                            'Periodo = Periodo & ", [ " & lstParcelas.ListItems(I).SubItems(3) & " ]"
                            Periodo = Periodo & ", " & IIf(Len(lstParcelas.ListItems(i).SubItems(3)) = 4, lstParcelas.ListItems(i).SubItems(3), Right(lstParcelas.ListItems(i).SubItems(3), 2) & Left(lstParcelas.ListItems(i).SubItems(3), 4))
                        End If
                    End If
             End If
            Base = lstParcelas.ListItems(i).SubItems(2)
            Base2 = lstParcelas.ListItems(i).SubItems(3)
        End If
    Next
    If Len(Periodo) > 240 Then
        Periodo = Mid(Periodo, 1, 240)
    End If
End Sub


Private Sub AtualizaValores()
    Dim i As Integer
    Valor = 0
    Juros = 0
    Multa = 0
    Saldo = 0
    Atualizacao = 0
    Desconto = 0
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked = True Then
            Valor = Valor + lstParcelas.ListItems(i).SubItems(5)
            
            Atualizacao = Atualizacao + lstParcelas.ListItems(i).SubItems(6)
            Juros = Juros + lstParcelas.ListItems(i).SubItems(7)
            Multa = Multa + lstParcelas.ListItems(i).SubItems(8)
            If lstParcelas.ListItems(i).SubItems(5) >= 1 Then
                Desconto = Desconto + lstParcelas.ListItems(i).SubItems(9)
            Else
                Desconto = Desconto + 0
            End If
            Saldo = Saldo + lstParcelas.ListItems(i).SubItems(10)
        End If
    Next
End Sub


Private Sub chkSelecionarTodos_Click(Value As Integer)
    Dim Item As Integer
    For Item = 1 To lstParcelas.ListItems.Count
        lstParcelas.ListItems(Item).Checked = Value
    Next
    txtValorPago = "0,0"
    txtdebitoRestante = txtTotalParc
    For Item = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(Item).Checked = True Then
                txtValorPago = Format(txtValorPago + CDbl(lstParcelas.ListItems(Item).SubItems(7)), Const_Monetario)
                If txtdebitoRestante = "" Then
                    txtdebitoRestante = 0
                End If
                txtdebitoRestante = Format(txtdebitoRestante - lstParcelas.ListItems(Item).SubItems(7), Const_Monetario)
        End If
    Next
End Sub

Private Sub Check1_Click()
    Dim Item As Integer
    lstParcelas.MarcarTodos Check1
    txtTotalParc = Nvl(txtTotalParc, 0)
    If Check1.Value Then
        txtdebitoRestante = Format(0, Const_Monetario)
    Else
        txtdebitoRestante = Format(txtTotalParc, Const_Monetario)
    End If
    txtValorPago = Format(CDbl(txtTotalParc) - CDbl(txtdebitoRestante), Const_Monetario)
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo TrataErro
    Dim Data As String
    If Trim(txtDtVence.Text) = "" Then
        Util.Avisa "Informe a data de vencimento"
        txtDtVence.SetFocus
        Exit Sub
    End If
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Data = DataServidor
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Data = Date
    End If
    If DateDiff("d", Data, CDate(txtDtVence.Text)) < 0 Then
        Util.Avisa "A data de vencimento não pode ser menor que a data atual"
        txtDtVence.SetFocus
        Exit Sub
    End If
    If Trim(txtInscricao.Text) = "" And txtImovel = "" Then
        Util.Avisa "Informe uma inscrição válida"
        txtInscricao.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11
    BuscarDebitos
    Screen.MousePointer = 0
    
    Exit Sub
TrataErro:
    Util.Erro Err.Description
    Exit Sub
    Resume
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    lstParcelas.ListItems.Clear
    cmdParcela.Enabled = True
    txtDtVence.SetFocus
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    Dim Barra As Boolean
    Dim Cobranca As New VSCobranca
    Dim CgcPref As String
    Dim M As Boolean
    Dim i As Integer
    
    M = False
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked Then
            M = True
            Exit For
        End If
    Next
    
    
    Barra = False
    If CodPagamento = 0 Then
        Informa "Não há extrato para ser impresso."
        Exit Sub
    End If
    Screen.MousePointer = 11
    'Avisa "Extrato deve ser impresso em 02(duas) vias."
'    If Confirma("O extrato será usado para pagamento do débito?") Then
    If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra_TITULO.rpt") Then
                Avisa "Arquivo do extrato não foi encontrado."
                Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtratoBarra.rpt") Then
                Avisa "Arquivo do extrato não foi encontrado."
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        Barra = True
'    Else
'        If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TDAMExtrato.rpt") Then Exit Sub
'    End If
    With Rpt
        'Formulas do Dam...
        .Formulas "DATAVENCIMENTO", txtDtVence
        .Formulas "PARCELA", "UNICA"
        .Formulas "NODOCUMENTO", CodPagamento
        If UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
            .Formulas "TXDAM", TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXTDAM"), 0), ".", ",")
        Else
            .Formulas "TXDAM", " "
        End If
        .Formulas "VENCIMENTONORMAL", txtDtVence
        .Formulas "EXTRATO", "Extrato de Negociação Nº " & CStr(CodPagamento)
        .Formulas "NOSSONUMERO", CodPagamento
        .Formulas "CODIGOTRIBUTO", Const_Extrato
        .Formulas "PERIDO", Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
        .Formulas "VALORTOTAL", txtValorPago
        .Formulas "EMISSAO", Format(Date, "DD/MM/YYYY")
        'DADOS DO SACADO...
        If txtInscricao <> "" Then
            .Formulas "NOME", txtInscricao & " - " & txtrazao
        Else
            .Formulas "NOME", txtImovel & " - " & txtrazao
        End If
        .Formulas "ENDERECO", txtEndereco
        'Atualizo os valores...
        Call AtualizaValores
         .Formulas "VALORTRIBUTO", CStr(Format(CDbl(txtValorPago), "###,###,###,##0.00"))
         '.Formulas "DEDUCAO", CStr(Format(Desconto, "###,###,###,##0.00"))
         .Formulas "DEDUCAO", "0,00"
         .Formulas "TAXAEXPEDIENTE", Format(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0), "###,###,###,##0.00")
         'COLOCO A DESCRIÇÃO DA TXDAM...
         If UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
            .Formulas "MENSAGEM1", "[ TXDAM - " & TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",") & " ]"
         End If
         .Formulas "VT_CARTEIRA", Temp.PegaParametro(Bdados, "CARTEIRA")
         '.Formulas "VALORJUROS", CStr(Format(Juros, "###,###,###,##0.00"))
         .Formulas "VALORJUROS", "0,00"
         '.Formulas "VALORMULTA", CStr(Format(Multa, "###,###,###,##0.00"))
         .Formulas "VALORMULTA", "0,00"
         '.Formulas "CORRECAO", CStr(Format(Atualizacao, "###,###,###,##0.00"))
         .Formulas "CORRECAO", "0,00"
        ',tcc_imposto_original + tcc_juros_atual + tcc_multa_atual + tcc_correcao_monetaria - tcc_desconto_concedido
         '.Formulas "VALORTOTAL", CStr(Format(CDbl(Nvl(Trim(txtValorPago), 0)) + Juros + Atualizacao + Multa + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)), "###,###,###,##0.00"))
         .Formulas "VALORTOTAL", CStr(Format(CDbl(Nvl(Trim(txtValorPago), 0)) + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)), "###,###,###,##0.00"))
         'FORMULAS DO EXTRATO...
        .Formulas "VT_EXTRATO ", "EXTRATO Nº " & CStr(CodPagamento)
        .Formulas "VT_PRAZO ", txtDtVence
        'Atualizo os Dados da Observação referente as Período e Tributo...
        Call Pega_Dados
        .Formulas "MENSAGEM2", "Tributo(s) - " & Tributo
        .Formulas "MENSAGEM3", "Período(s) - " & Periodo
        .Formulas "VT_CONTRIBUINTE", IIf(Trim(txtInscricao) = "", txtImovel, txtInscricao)
        .Formulas "VT_RAZAO", txtrazao
        .Formulas "VT_ENDERECO", txtEndereco
        .Formulas "VT_OBS_GERAL ", txtObservacao
        '.Formulas "RAZAO", txtrazao
            '.Formulas "ENDERECO", txtEndereco
        .Selecao = "{TAB_PAGAMENTO_EXTRATO.TPE_COD_PAGAMENTO_EXTRATO} =" & CodPagamento
        
        If Barra Then
            .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
            .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
            Cobranca.ImprimeDamBarra Rpt, txtInscricao, Const_Extrato, CStr(Format(CDbl(Nvl(Trim(txtValorPago), 0)) + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)), "###,###,###,##0.00")), Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), PicBarra, txtDtVence, 0, CodPagamento
        Else
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        End If
        .Titulo = "Extrato de Lançamento"
        .Arvore = False
        .Visualizar
    End With
    Screen.MousePointer = 0
    Set Rpt = Nothing
End Sub

Private Sub cmdParcela_Click()
    Dim Sql As String
    Dim CCorrente As New ContaCorrente
    Dim rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim i As Integer
    Dim Campos As String, Valores As String
    Dim Codigo As String
    Dim Contas As Integer
    On Error Resume Next
    Contas = 0
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked Then
            Marcou = True
            Contas = Contas + 1
            If Contas > 1 Then Exit For
        End If
    Next
'    If Contas < 2 Then
'        Avisa "O numero de documentos não pode ser inferior a 02(dois) em um extrato."
'        txtInscricao.SetFocus
'        Exit Sub
'    End If
    If txtImovel <> "" Then
        Codigo = txtImovel
    Else
        Codigo = txtInscricao
    End If
    If Marcou = False Then Util.Avisa "Selecione um débito para impressão.": Exit Sub
    txtInscricao.Tag = ""
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    
    
    If Aplicacoes.municipio = "PETROLINA" Then
        If (Fix(txtdebitoRestante) <> 0) And Trim(txtImovel) <> "" Then
            Sql = "SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
            If Bdados.AbreTabela(Sql, rs) Then
                Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(0), "TOC_DESCONTO", _
                    "TOC_COD_OBRIGACAO =" & rs!TOC_COD_OBRIGACAO
                Conta.ExecutaAtualizacao txtImovel, etiImovel, , rs!TOC_COD_OBRIGACAO, , txtDtVence
                Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(0), "TOC_DESCONTO", _
                    "TOC_COD_OBRIGACAO =" & rs!TOC_COD_OBRIGACAO
            End If
        End If
    End If
    
    CodPagamento = Conta.GeraCodPagamento(EtsExtratoPagamento)
    ListaDocs = ""
    Campos = "TPE_INSCRICAO, TPE_COD_PAGAMENTO_EXTRATO, TPE_TGT_COD_PAGAMENTO,TPE_TIP_COD_IMPOSTO,TPE_SUB_VALOR,TPE_TIPO_DOCUMENTO,TPE_SUB_PERIODO"
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked Then
            
            With lstParcelas.ListItems(i)
                Valores = Bdados.PreparaValor(Trim(.SubItems(1)), CodPagamento, .Text, .SubItems(11), Bdados.Converte(Trim(.SubItems(10)), TCMonetario), 1, .SubItems(3))
                Bdados.GravaDados "TAB_PAGAMENTO_EXTRATO", Valores, Campos, "TPE_TGT_COD_PAGAMENTO=" & .Text & " and TPE_COD_PAGAMENTO_EXTRATO=" & CodPagamento
            End With
        End If
    Next
    
    ListaDocs = Left(ListaDocs, Len(ListaDocs) - 3)
    Conta.GeraPagamento Codigo, "", Const_Extrato, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtDtVence, CDbl(txtValorPago), 0, 0, CodPagamento, 0, 0, 0, , EtcCreditoTributario
    cmdParcela.Enabled = False
    Screen.MousePointer = 0
    Bdados.FechaTabela rs
    If Util.Confirma("Extrato " & CodPagamento & " gerado com sucesso. Deseja Imprimir?") Then
        cmdImprime_Click
    Else
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub txtCotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub lstParcelas_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    If Item.Checked Then
        If Nvl(txtdebitoRestante, 0) = txtTotalParc Then
            txtdebitoRestante = Format(CDbl(txtTotalParc) - CDbl(Item.SubItems(10)), Const_Monetario)
        Else
            txtdebitoRestante = Format(CDbl(txtdebitoRestante) - CDbl(Item.SubItems(10)), Const_Monetario)
        End If
    Else
        txtdebitoRestante = Format(CDbl(txtdebitoRestante) + CDbl(Item.SubItems(10)), Const_Monetario)
    End If
    txtValorPago = Format(CDbl(txtTotalParc) - CDbl(txtdebitoRestante), Const_Monetario)
End Sub

Private Sub txtDtVence_LostFocus()
    txtDtVence = Edita.FormataTexto(txtDtVence, Data)
    If Trim(txtDtVence) = "" Then Exit Sub
    If CDate(txtDtVence) < Date Then
        Util.Avisa "Data inválida."
        txtDtVence.SetFocus
    End If
End Sub


Private Sub BuscarDebitos()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim ValorTotal As Double
    Dim ja As Boolean
    Dim Obrig As New Obrigacao
        
    If txtInscricao <> "" Then
        Conta.ExecutaAtualizacao txtInscricao, etiContribuinte, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), , , txtDtVence, CDbl(Nvl(txtOriginal, 0)), CDbl(Nvl(txtAtualizacao, 0)), CDbl(Nvl(txtMulta, 0)), CDbl(Nvl(txtJuros, 0))
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, txtInscricao, , , etlNaoPagos, , etiContribuinte, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Util.Avisa "Consulta sem resultados."
        End If
    Else
        If Aplicacoes.municipio = "PETROLINA" Then
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(50), "TOC_DESCONTO", _
                "TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
        End If
        Conta.ExecutaAtualizacao txtImovel, etiImovel, , , , txtDtVence, CDbl(Nvl(txtOriginal, 0)), CDbl(Nvl(txtAtualizacao, 0)), CDbl(Nvl(txtMulta, 0)), CDbl(Nvl(txtJuros, 0))
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, txtImovel, , , etlNaoPagos, , etiImovel) Then
            Util.Avisa "Consulta sem resultados."
        End If
        If Aplicacoes.municipio = "PETROLINA" Then
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(0), "TOC_DESCONTO", _
                "TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
        End If
    End If
    lstParcelas.AtualizarQtd
    If lstParcelas.ListItems.Count > 0 Then ValorTotal = Format(lstParcelas.Colunas(11).Soma, Const_Monetario)
    
    Bdados.FechaTabela rs
    Check1.Value = 1
    chkSelecionarTodos_Click True: lstParcelas.MarcarTodos True
    txtTotalParc = Format(ValorTotal, Const_Monetario)
    txtdebitoRestante = Format(0, Const_Monetario)
    txtValorPago = ValorTotal
    Screen.MousePointer = 0
End Sub

Private Sub txtImovel_LostFocus()
  Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtrazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtImovel.SetFocus
        End If
    End If
End Sub

Private Sub txtInscricao_LostFocus()
    If txtInscricao = "" Then Exit Sub
    txtInscricao = BuscaContribuinte(txtInscricao, txtrazao, txtEndereco)
End Sub

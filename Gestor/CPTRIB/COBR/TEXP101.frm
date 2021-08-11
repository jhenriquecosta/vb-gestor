VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Begin VB.Form TEXP101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuinte"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   30
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TEXP101.frx":0000
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
      TabIndex        =   29
      Top             =   5730
      Width           =   2655
   End
   Begin VTOcx.fraVISUAL fraVISUAL4 
      Height          =   765
      Left            =   30
      TabIndex        =   23
      Top             =   6000
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
         TabIndex        =   14
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
         TabIndex        =   15
         Tag             =   "Valor Pago"
         Top             =   345
         Width           =   1185
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   7
         Left            =   2535
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   16
         Tag             =   "Debito Pendente"
         Top             =   345
         Width           =   1215
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2220
      Left            =   30
      TabIndex        =   21
      Top             =   660
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   3916
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   930
         TabIndex        =   3
         Top             =   1800
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
         CorFundo        =   -2147483644
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   7470
         TabIndex        =   31
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
         CorFundo        =   -2147483629
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
         TabIndex        =   4
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
         TabIndex        =   24
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
      TabIndex        =   20
      Top             =   7935
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   7590
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   120
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
         TabIndex        =   11
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
         TabIndex        =   12
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
      TabIndex        =   18
      Top             =   -525
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   17
      Top             =   -360
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   1138
      Icone           =   "TEXP101.frx":2123
   End
   Begin VTOcx.grdVISUAL lstParcelas 
      Height          =   2955
      Left            =   30
      TabIndex        =   13
      Top             =   3015
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   5212
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
      TabIndex        =   22
      Top             =   6795
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
         TabIndex        =   28
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
         TabIndex        =   10
         Top             =   375
         Width           =   9210
      End
   End
End
Attribute VB_Name = "TEXP101"
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
            Desconto = Desconto + lstParcelas.ListItems(i).SubItems(9)
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
    Dim Data As Date
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
    On Error GoTo TRATA
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
        .Formulas "VENCIMENTONORMAL", txtDtVence
        .Formulas "EXTRATO", "Extrato Nº " & CodPagamento
        .Formulas "NOSSONUMERO", CodPagamento
        .Formulas "CODIGOTRIBUTO", Const_Extrato
        If UCase(AplicacoesVTFuncoes.municipio) = "PETROLINA" Then
            .Formulas "TXDAM", TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ",")
        End If
        .Formulas "PERIDO", Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
'        .Formulas "VALORTOTAL", CDbl(Nvl(Trim(txtValorPago), 0)) '+ Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)
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
       '  '.Formulas "VALORTOTAL", Trim(txtValorPago)
       '  'FORMULAS DO EXTRATO...
        .Formulas "VT_EXTRATO ", "EXTRATO Nº " & CStr(CodPagamento)
        .Formulas "VT_PRAZO ", txtDtVence
        'Atualizo os Dados da Observação referente as Período e Tributo...
        Call Pega_Dados
        .Formulas "MENSAGEM2", "Tributo(s) - " & Tributo
        .Formulas "MENSAGEM3", "Período(s) - " & Periodo
        If Trim(txtInscricao) = "" Then
            txtImovel = Trim(txtImovel)
            If Len(Trim(txtImovel)) = 18 Then
                .Formulas "VT_CONTRIBUINTE", Left(txtImovel, 1) & "." & Mid(txtImovel, 2, 4) & "." & Mid(txtImovel, 6, 3) & "." & Mid(txtImovel, 9, 2) & "." & Mid(txtImovel, 11, 4) & "." & Right(txtImovel, 4)
            Else
                .Formulas "VT_CONTRIBUINTE", txtImovel
            End If
        Else
            
            If Len(txtInscricao) = 18 Then
                .Formulas "VT_CONTRIBUINTE", Left(txtInscricao, 1) & "." & Mid(txtInscricao, 2, 4) & "." & Mid(txtInscricao, 6, 3) & "." & Mid(txtInscricao, 9, 2) & "." & Mid(txtInscricao, 11, 4) & "." & Right(txtInscricao, 4)
            Else
                .Formulas "VT_CONTRIBUINTE", txtInscricao
            End If
        End If
        
        
        .Formulas "VT_RAZAO", txtrazao
        .Formulas "VT_ENDERECO", txtEndereco
        .Formulas "VT_OBS_GERAL ", txtObservacao
        '.Formulas "RAZAO", txtrazao
        
        .Selecao = "{TAB_PAGAMENTO_EXTRATO.TPE_COD_PAGAMENTO_EXTRATO} =" & CodPagamento
        If Barra Then
            .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                .Formulas "VT_NOME_DAM", Temp.PegaParametro(Bdados, "NOME DAM")
            End If
            CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
            .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
            Cobranca.ImprimeDamBarra Rpt, txtInscricao, Const_Extrato, CStr(Format(CDbl(Nvl(Trim(txtValorPago), 0)) + CDbl(Nvl(TrocaPic(Temp.PegaParametro(Bdados, "TXTDAM"), ".", ","), 0)), "###,###,###,##0.00")), Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), PicBarra, txtDtVence, 0, CodPagamento
        Else
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        End If
        'Bdados.GravaDados "TAB_PARAMETRO_TEXTO", ListaDocs, "TPT_TEXTO", "TPT_PARAMETRO = 'DOCUMENTOS EXTRATO'"
        .Titulo = "Extrato de Lançamento"
        .Arvore = False
        .Visualizar
        'Bdados.GravaDados "TAB_PARAMETRO_TEXTO", "", "TPT_TEXTO", "TPT_PARAMETRO = 'DOCUMENTOS EXTRATO'"
    End With
    Screen.MousePointer = 0
    Set Rpt = Nothing
    Exit Sub
TRATA:
    Avisa Err.Number & " - " & Err.Description
    Exit Sub
    Resume
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
    Dim TotalIPTU As Double
    Dim TotalIptuPago As Double
    Dim CodIPTU As String
    
    On Error Resume Next
    If Len(cboImposto) = 0 Then
        Mensagem "A conclusão desta rotina é feita com tributos de MESMO tipo, selecione o tributo e click em BUSCAR"
        cboImposto.SetFocus
        Exit Sub
    End If
    Contas = 0
    For i = 1 To lstParcelas.ListItems.Count
        If lstParcelas.ListItems(i).Checked Then
            Marcou = True
            Contas = Contas + 1
            If Contas > 1 Then Exit For
        End If
    Next
    If Contas < CInt(Nvl(Temp.PegaParametro(Bdados, "MINIMO PARCELAS EXTRATO"), 0)) Then
        Avisa "O numero de documentos não pode ser inferior a 02(dois) em um extrato."
        txtInscricao.SetFocus
        Exit Sub
    End If
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
        TotalIptuPago = 0
        TotalIPTU = 0
        CodIPTU = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
        For i = 1 To lstParcelas.ListItems.Count
            If lstParcelas.ListItems(i).SubItems(11) = CodIPTU Then
                TotalIPTU = TotalIPTU + lstParcelas.ListItems(i).SubItems(10)
                If lstParcelas.ListItems(i).Checked Then
                    TotalIptuPago = TotalIptuPago + lstParcelas.ListItems(i).SubItems(10)
                End If
            End If
        Next
        If (TotalIPTU <> TotalIptuPago) And Trim(txtImovel) <> "" Then
            Sql = "SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200' and TOC_STATUS_OBRIGACAO IN (" & Const_NaoPagos & ")"
            If Bdados.AbreTabela(Sql, rs) Then
                Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(0), "TOC_DESCONTO", _
                    "TOC_COD_OBRIGACAO =" & rs!TOC_COD_OBRIGACAO
                Conta.ExecutaAtualizacao txtImovel, etiImovel, , rs!TOC_COD_OBRIGACAO, , txtDtVence, , , , , "" & cboImposto.Coluna(0).Valor
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
                Valores = Bdados.PreparaValor(Bdados.Converte(Trim(.SubItems(1)), tctexto), CodPagamento, .Text, .SubItems(11), Bdados.Converte(Trim(.SubItems(10)), TCMonetario), 1, .SubItems(3))
                Bdados.GravaDados "TAB_PAGAMENTO_EXTRATO", Valores, Campos, "TPE_TGT_COD_PAGAMENTO=" & .Text & " and TPE_COD_PAGAMENTO_EXTRATO=" & CodPagamento
                '
                Bdados.Executa ("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_REMESSA=1, TOC_STATUS_OBRIGACAO=16 WHERE TOC_COD_OBRIGACAO=" & lstParcelas.ListItems(i))
                Bdados.Executa ("UPDATE TAB_CONTA_CONTRIBUINTE SET TCC_STATUS_CONTA=6 WHERE TCC_CODIGO_CONTA=" & lstParcelas.ListItems(i))
            End With
        End If
    Next
                'BCP
                Dim ob As New obrigacao
                Dim p As String
                p = Format(Now, "YYYY")
                If ob.CriaObrigacao(CStr(cboImposto.Coluna(0).Valor), p, _
                    p, Codigo, CCur(txtValorPago), etsCreditoOriginalAberto, etsCriaNova, _
                    txtDtVence, , , , , , , 0, 0, txtImovel, IIf(Len(txtImovel) = 0, etiContribuinte, etiImovel)) Then
                End If
                'FIM BCP
    ListaDocs = Left(ListaDocs, Len(ListaDocs) - 3)
    Conta.GeraPagamento txtInscricao, txtImovel, Const_Extrato, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtDtVence, CDbl(txtValorPago), 0, 0, CodPagamento, 0, 0, 0, , EtcCreditoTributario
'    cmdParcela.Enabled = False
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
    Dim Obrig As New obrigacao
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Obrig.PreencheComboTributo cboImposto, False
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
    On Error GoTo TRATA
    txtDtVence = Edita.FormataTexto(txtDtVence, Data)
    If Trim(txtDtVence) = "" Then Exit Sub
    If CDate(txtDtVence) < Date Then
        Util.Avisa "Data inválida."
        txtDtVence.SetFocus
    End If
    Exit Sub
TRATA:
    If Err.Number = 13 Then
        Avisa "Data Inválida."
    End If
End Sub


Private Sub BuscarDebitos()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Conta As New ContaCorrente
    Dim ValorTotal As Double
    Dim ja As Boolean
    Dim Obrig As New obrigacao
    On Error GoTo TRATA
    If txtInscricao <> "" Then
        Conta.ExecutaAtualizacao txtInscricao, etiContribuinte, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False), , , txtDtVence, , , , , "" & cboImposto.Coluna(0).Valor
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, txtInscricao, CStr("" & cboImposto.Coluna(0).Valor), , etlNaoPagos, , etiContribuinte, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Util.Avisa "Consulta sem resultados."
        End If
    Else
        If Aplicacoes.municipio = "PETROLINA" Then
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(50), "TOC_DESCONTO", _
                "TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
        End If
        Conta.ExecutaAtualizacao txtImovel, etiImovel, , , , txtDtVence, , , , , "" & cboImposto.Coluna(0).Valor
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstParcelas, txtImovel, CStr("" & cboImposto.Coluna(0).Valor), , etlNaoPagos, , etiImovel, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
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
    txtValorPago = Format(ValorTotal, Const_Monetario)
    Screen.MousePointer = 0
    cmdParcela.Enabled = True
    
    Exit Sub
TRATA:
    Avisa Err.Number & " - " & Err.Description
    Exit Sub
    Resume
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

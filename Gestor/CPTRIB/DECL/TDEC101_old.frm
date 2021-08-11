VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.1#0"; "VTControles.ocx"
Begin VB.Form TDEC101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DECL101"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cboVISUAL cboTipo 
      Height          =   315
      Left            =   2130
      TabIndex        =   2
      Top             =   1050
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      Caption         =   "Tipo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   6240
      Left            =   60
      TabIndex        =   36
      Top             =   1440
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   11007
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TDEC101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5850
         Index           =   0
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   10319
         _Version        =   131082
         TabGuid         =   "TDEC101.frx":00C8
         Begin VTOcx.fraFUTURO fraFUTURO1 
            Height          =   7000
            Index           =   0
            Left            =   -60
            TabIndex        =   40
            Top             =   -60
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   12356
            Caption         =   "Apuração"
            Descricao       =   "Informe o intervalo das notas fiscais emitidas no período"
            corFaixa        =   16384
            corFundo        =   14737632
            corTexto        =   16384
            Icone           =   "TDEC101.frx":00F0
            Ocultavel       =   0   'False
            Altura          =   2000
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   0
               Left            =   1110
               TabIndex        =   50
               Top             =   2280
               Width           =   2385
               _ExtentX        =   4207
               _ExtentY        =   556
               Caption         =   "Aliquota"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   5
               Left            =   3840
               TabIndex        =   8
               Top             =   2670
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   556
               Caption         =   "Total imposto retido"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   7
               Left            =   4200
               TabIndex        =   10
               Top             =   3030
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   556
               Caption         =   "Imposto devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   6
               Left            =   450
               TabIndex        =   9
               Top             =   3030
               Width           =   3045
               _ExtentX        =   5371
               _ExtentY        =   556
               Caption         =   "Saldo tributável"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   3
               Left            =   60
               TabIndex        =   6
               Top             =   1140
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   556
               Caption         =   "Valor total em notas"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   4
               Left            =   60
               TabIndex        =   7
               Top             =   2670
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               Caption         =   "Total sujeito a ICMS"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   2
               Left            =   4710
               TabIndex        =   5
               Top             =   780
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   556
               Caption         =   "Nota Final"
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   4
               Top             =   780
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               Caption         =   "Nota Inicial"
               Text            =   ""
               Restricao       =   2
            End
            Begin VTOcx.txtVISUAL txtItemDecl 
               Height          =   945
               Index           =   8
               Left            =   2070
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   5910
               Width           =   5685
               _ExtentX        =   10028
               _ExtentY        =   1667
               Caption         =   "Obs"
               Text            =   ""
               AlinhamentoRotuloVertical=   0
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5850
         Index           =   1
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   10319
         _Version        =   131082
         TabGuid         =   "TDEC101.frx":09CA
         Begin VTOcx.fraFUTURO fraFUTURO1 
            Height          =   6075
            Index           =   1
            Left            =   -60
            TabIndex        =   43
            Top             =   -60
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   10716
            Caption         =   "Notas Recebidas"
            Descricao       =   "Informe as notas fiscais de servicos recebidas no período"
            corFaixa        =   16384
            corFundo        =   14737632
            corTexto        =   16384
            Icone           =   "TDEC101.frx":09F2
            Ocultavel       =   0   'False
            Altura          =   2000
            Begin VTOcx.txtVISUAL txtSTImposto 
               Height          =   525
               Left            =   5760
               TabIndex        =   48
               Top             =   3360
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   926
               Caption         =   "Imposto retido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cmdVISUAL cmdRetirarNotaST 
               Height          =   375
               Left            =   6780
               TabIndex        =   25
               ToolTipText     =   "Excluir"
               Top             =   3900
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   2
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdAdicionarNotaST 
               Height          =   375
               Left            =   6330
               TabIndex        =   24
               ToolTipText     =   "Adicionar"
               Top             =   3900
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.grdVISUAL grdST 
               Height          =   1875
               Left            =   90
               TabIndex        =   46
               Top             =   4290
               Width           =   7380
               _ExtentX        =   13018
               _ExtentY        =   3307
               Caption         =   "Notas recebidas"
               CorTitulo       =   5346129
               CorCaption      =   16777215
               CorDica         =   192
            End
            Begin VTOcx.txtVISUAL txtSTSaldo 
               Height          =   525
               Left            =   4260
               TabIndex        =   45
               Top             =   3360
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   926
               Caption         =   "Saldo tributável"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtSTIcms 
               Height          =   525
               Left            =   1440
               TabIndex        =   23
               Top             =   3360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   926
               Caption         =   "Valor Sujeito ICMS"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtSTValor 
               Height          =   525
               Left            =   5760
               TabIndex        =   22
               Top             =   2820
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   926
               Caption         =   "Valor"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtSTRecolhimento 
               Height          =   525
               Left            =   4260
               TabIndex        =   21
               Top             =   2820
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   926
               Caption         =   "Data Pago"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtSTNumNota 
               Height          =   525
               Left            =   1440
               TabIndex        =   19
               Top             =   2820
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   926
               Caption         =   "Nº Nota"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VTOcx.txtVISUAL txtSTEmissao 
               Height          =   525
               Left            =   2760
               TabIndex        =   20
               Top             =   2820
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   926
               Caption         =   "Data Emissão"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               AlinhamentoTexto=   1
            End
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               Caption         =   "Prestador de Serviço"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   2085
               Index           =   0
               Left            =   120
               TabIndex        =   44
               Top             =   690
               Width           =   7335
               Begin VTOcx.txtVISUAL txtSTRazao 
                  Height          =   315
                  Left            =   810
                  TabIndex        =   13
                  Top             =   600
                  Width           =   6435
                  _ExtentX        =   11351
                  _ExtentY        =   556
                  Caption         =   "Nome"
                  Text            =   ""
               End
               Begin VTOcx.txtVISUAL txtSTInscricao 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   12
                  Top             =   240
                  Width           =   3405
                  _ExtentX        =   6006
                  _ExtentY        =   556
                  Caption         =   "CPF/CNPJ/IM"
                  Text            =   ""
                  Restricao       =   2
                  AgruparValores  =   0   'False
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtSTEndereco 
                  Height          =   315
                  Left            =   510
                  TabIndex        =   14
                  Top             =   960
                  Width           =   6735
                  _ExtentX        =   11880
                  _ExtentY        =   556
                  Caption         =   "Endereço"
                  Text            =   ""
               End
               Begin VTOcx.txtVISUAL txtSTBairro 
                  Height          =   315
                  Left            =   780
                  TabIndex        =   15
                  Top             =   1320
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   556
                  Caption         =   "Bairro"
                  Text            =   ""
               End
               Begin VTOcx.txtVISUAL txtSTCep 
                  Height          =   315
                  Left            =   5100
                  TabIndex        =   16
                  Top             =   1320
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   556
                  Caption         =   "CEP"
                  Text            =   ""
                  Formato         =   4
                  Restricao       =   2
               End
               Begin VTOcx.txtVISUAL txtSTCidade 
                  Height          =   315
                  Left            =   690
                  TabIndex        =   17
                  Top             =   1680
                  Width           =   4245
                  _ExtentX        =   7488
                  _ExtentY        =   556
                  Caption         =   "Cidade"
                  Text            =   ""
               End
               Begin VTOcx.cboVISUAL cboSTUf 
                  Height          =   315
                  Left            =   5250
                  TabIndex        =   18
                  Top             =   1680
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   556
                  Caption         =   "UF"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
               End
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5850
         Index           =   2
         Left            =   30
         TabIndex        =   39
         Top             =   30
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   10319
         _Version        =   131082
         TabGuid         =   "TDEC101.frx":12CC
         Begin VTOcx.fraFUTURO fraFUTURO1 
            Height          =   7000
            Index           =   2
            Left            =   -60
            TabIndex        =   42
            Top             =   -60
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   12356
            Caption         =   "Notas Fiscais"
            Descricao       =   "Informe as notas fiscais canceladas e com imposto retido"
            corFaixa        =   16384
            corFundo        =   14737632
            corTexto        =   16384
            Icone           =   "TDEC101.frx":12F4
            Ocultavel       =   0   'False
            Altura          =   2000
            Begin VTOcx.txtVISUAL txtNotaRetida 
               Height          =   315
               Left            =   180
               TabIndex        =   29
               Top             =   3360
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               Caption         =   "Nº Nota"
               Text            =   ""
            End
            Begin VTOcx.grdVISUAL grdVISUAL1 
               Height          =   2355
               Left            =   90
               TabIndex        =   49
               Top             =   3780
               Width           =   7380
               _ExtentX        =   13018
               _ExtentY        =   4154
               Caption         =   "Notas com imposto retido"
               CorTitulo       =   5346129
               CorCaption      =   16777215
               CorDica         =   192
               OcultarRodape   =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdAdicionarNotaRetida 
               Height          =   375
               Left            =   2160
               TabIndex        =   30
               ToolTipText     =   "Adicionar"
               Top             =   3360
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdRetirarNotaRetida 
               Height          =   375
               Left            =   2610
               TabIndex        =   31
               ToolTipText     =   "Excluir"
               Top             =   3360
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   2
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdRetirarNotaCancelada 
               Height          =   375
               Left            =   2610
               TabIndex        =   28
               ToolTipText     =   "Excluir"
               Top             =   720
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   2
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdAdicionarNotaCancelada 
               Height          =   375
               Left            =   2160
               TabIndex        =   27
               ToolTipText     =   "Adicionar"
               Top             =   720
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   661
               Caption         =   ""
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.grdVISUAL grdCanceladas 
               Height          =   2355
               Left            =   90
               TabIndex        =   47
               Top             =   1140
               Width           =   7380
               _ExtentX        =   13018
               _ExtentY        =   4154
               Caption         =   "Notas canceladas"
               CorTitulo       =   5346129
               CorCaption      =   16777215
               CorDica         =   192
               OcultarRodape   =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtNotaCancelada 
               Height          =   315
               Left            =   210
               TabIndex        =   26
               Top             =   720
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               Caption         =   "Nº Nota"
               Text            =   ""
            End
         End
      End
   End
   Begin VTOcx.txtVISUAL txtIM 
      Height          =   315
      Left            =   510
      TabIndex        =   0
      Top             =   690
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "IM"
      Text            =   ""
      Restricao       =   2
      AgruparValores  =   0   'False
      RetirarMascara  =   0   'False
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   34
      Top             =   7740
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5790
         TabIndex        =   32
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6780
         TabIndex        =   41
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   1138
      Icone           =   "TDEC101.frx":160E
   End
   Begin VTOcx.txtVISUAL txtRazao 
      Height          =   315
      Left            =   2130
      TabIndex        =   35
      Top             =   690
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      CorFundo        =   -2147483626
   End
   Begin VTOcx.txtVISUAL txtPeriodo 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   1050
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      Caption         =   "Período"
      Text            =   ""
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtData 
      Height          =   315
      Left            =   5910
      TabIndex        =   3
      Top             =   1050
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      Caption         =   "Data"
      Text            =   ""
      Formato         =   0
   End
End
Attribute VB_Name = "TDEC101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declaracao As VsTFuncoes.cDeclaracao
Private GerarIM As Boolean
Private AliqISSQN As Double, ISSQNFixo As Double
Private AliqISSST As Double, ISSSTFixo As Double
Dim Atividade As New VsTEcon.Atividade

Private Sub cboTipo_Click()
    Dim Temp As Byte
    
    
    Temp = cboTipo.Coluna(1).Valor
    If Temp = 0 Then Exit Sub
    If Trim$(txtIM) <> "" And Trim$(txtPeriodo) <> "" Then
        If Declaracao.Buscar(txtIM, txtPeriodo, perMMAAAA) Then
            If Temp = decNormal Then
                Util.Avisa "Contribuinte já possui declaração no período."
               ' txtPeriodo.SetFocus
            End If
        Else
            If Temp = decSubstitutiva Then
                Util.Avisa "Contribuinte não possui declaração no período."
                'txtPeriodo.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub cmdAdicionarNotaCancelada_Click()
    AdicionarCancelada grdCanceladas, txtNotaCancelada
End Sub

Private Sub AdicionarCancelada(ByRef grd As Object, ByRef Nota As Object)
    If Trim$(Nota) <> "" Then
        grd.ListItems.Add , , Nota
        Nota = ""
        Nota.SetFocus
    End If
End Sub

Private Sub cmdAdicionarNotaST_Click()
    AdicionarST grdST, txtSTInscricao, txtSTNumNota, txtSTEmissao, txtSTRecolhimento, txtSTValor, txtSTIcms, txtSTSaldo, txtSTImposto
End Sub

Private Sub AdicionarST(ByRef grd As Object, ByRef Im As Object, ByRef Nota As Object, ByRef Emissao As Object, ByRef Recolhimento As Object, ByRef Valor As Object, ByRef ICMS As Object, ByRef Saldo As Object, ByRef Imposto As Object)
    Dim Linha As Object
    
    If Trim$(Im) <> "" And Trim$(Nota) <> "" Then
        Set Linha = grdST.ListItems.Add(, , Im)
        Linha.SubItems(1) = Nota
        Linha.SubItems(2) = Emissao
        Linha.SubItems(3) = Recolhimento
        Linha.SubItems(4) = Valor
        Linha.SubItems(5) = ICMS
        Linha.SubItems(6) = Saldo
        Linha.SubItems(7) = Imposto
    
        Nota = ""
        Emissao = ""
        Recolhimento = ""
        Valor = ""
        ICMS = ""
        Saldo = ""
        Imposto = ""
        Im.SetFocus
    End If

End Sub

Private Sub cmdRetirarNotaCancelada_Click()
    If Not grdCanceladas.SelectedItem Is Nothing Then
        grdCanceladas.ListItems.Remove grdCanceladas.SelectedItem
    End If
End Sub

Private Sub cmdRetirarNotaST_Click()
    If Not grdST.SelectedItem Is Nothing Then
        If Util.Confirma("Retirar nota " & grdST.SelectedItem.SubItems(1) & " ?") Then
            grdST.ListItems.Remove grdST.SelectedItem
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If txtData = "" Then txtData = Date
    If txtIM = "" Then Exit Sub
    Declaracao.Im = txtIM
    Declaracao.Periodo = txtPeriodo
    Declaracao.Data = txtData
    Declaracao.Origem = orgSistema
    Declaracao.Recepcao = Date
    Declaracao.Versao = "1"
    'Declaracao.Tipo = Right(Me.Tag, 1) POIS NAO GRAVAVA SUBSTITUTIVA
    Declaracao.Tipo = cboTipo.Coluna(1).Valor
    carregarItens
    
    If Declaracao.Gravar(FormatoPeriodo:=perMMAAAA) Then
        'Util.Avisa "Declaração gravada com sucesso."
        Edita.LimpaCampos Me
        txtIM.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    cabVISUAL1.Exibir Bdados, Me.Tag, App.Path
    rodVISUAL1.Exibir Bdados, Me.Tag
    
    Set Imposto = New VsTFuncoes.VSImposto
    
    prepararGridST
    prepararGridCancelada
    Set Declaracao = New cDeclaracao
    cboTipo.PreencherGeral Bdados, "TIPO DECLARACAO"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Declaracao = Nothing
End Sub

Private Sub txtData_LostFocus()
    If Trim$(txtData) = "" Then txtData = Date
End Sub

Public Sub carregarItens()
    Dim Controle As Object
    
    Dim Item As cItemDeclaracao
        
    Declaracao.Itens.Limpar
    
    For Each Controle In txtItemDecl
        Set Item = New cItemDeclaracao
        Item.Numero = Controle.Index
        Item.Valor = Util.Nvl(Controle.Text, 0)
        
        Declaracao.Itens.Adicionar Item
    Next


    For Each Linha In grdCanceladas.ListItems
        For Each Coluna In grdCanceladas.ColumnHeaders
            Set Item = New cItemDeclaracao
            Item.Numero = Util.ParseString(Coluna.Key, ":", 2)
            If Coluna.Index = 1 Then
                Item.Valor = Linha.Text
            Else
                Item.Valor = Linha.ListSubItems(Coluna.Index - 1).Text
            End If
            Declaracao.Itens.Adicionar Item
        Next
    Next
End Sub

Private Sub txtIM_LostFocus()
    
    If Trim$(txtIM) <> "" Then
        If Not buscarContribuinte(txtIM, txtRazao) Then
            Util.Avisa "Contribuinte não encontrado."
            txtIM = "": txtRazao = ""
            txtIM.SetFocus
        Else
            AliqISSQN = Atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
            Declaracao.tciAtividade = Atividade.Nome
            Set Atividade = Nothing
        End If
    End If
End Sub


Private Sub txtItemDecl_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4, 5
            CalcularImposto txtItemDecl(3), txtItemDecl(5), txtItemDecl(4), txtItemDecl(6), txtItemDecl(7), txtItemDecl(0)
        Case 8
            SSActiveTabs1.Tabs(2).Selected = True
            txtSTInscricao.SetFocus
    End Select
End Sub

Private Sub CalcularImposto(ByRef Total As Object, ByRef Retido As Object, ByRef ICMS As Object, ByRef Tributavel As Object, ByRef Imposto As Object, ByRef Aliquota As Object)
    Total = Util.Nvl(Trim$(Total.Text), 0)
    Retido = Util.Nvl(Trim$(Retido.Text), 0)
    ICMS = Util.Nvl(Trim$(ICMS.Text), 0)
    
    Aliquota = AliqISSQN * 100
    Tributavel = Total - ICMS
    If AliqISSQN > 0 Then
        Imposto = Tributavel * AliqISSQN - Retido
    Else
        Imposto = ISSQNFixo - Retido
    End If
End Sub
Private Sub txtPeriodo_LostFocus()
    If Trim(txtPeriodo) = "" Then Exit Sub
    txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
    AliqISSQN = Atividade.BuscaAliquotaAtividade(Bdados, txtIM, txtPeriodo, ISSQNFixo)
    cboTipo_Click
End Sub

Private Sub txtSTEmissao_LostFocus()
    If Trim$(txtSTEmissao) = "" Then txtSTEmissao = Date
End Sub

Private Sub txtSTIcms_LostFocus()
    txtSTIcms = Util.Nvl(Trim$(txtSTIcms), 0)
    txtSTSaldo = Util.Nvl(txtSTValor, 0) - txtSTIcms
    If AliqISSST > 0 Then
        txtSTImposto = txtSTSaldo * AliqISSST
    Else
        txtSTImposto = ISSSTFixo
    End If
End Sub

Private Function buscarContribuinte(ByRef Inscricao As Object, ByRef Nome As Object, Optional ByRef Endereco As Object, _
                    Optional ByRef Bairro As Object, Optional ByRef CEP As Object, Optional ByRef Cidade As Object, Optional ByRef UF As Object) As Boolean
    Dim Im As Boolean
    
    Im = False
    If Trim(Inscricao) = "" Then Exit Function
    Inscricao.Text = Edita.TiraTudo(Inscricao.Text)
    Select Case Len(Inscricao.Text)
        Case 10
            Inscricao.Text = Imposto.FormataInscricao(Inscricao.Text, InscContrib): Im = True
        Case 11
            Inscricao = Edita.FormataTexto(Inscricao, Cpf)
        Case 14
            Inscricao = Edita.FormataTexto(Inscricao, Cgc)
    End Select
    If Trim(Inscricao) = "" Then Exit Function
    
    Dim Sql As String, Rs As Object
    Sql = "SELECT tci_im, tci_nome, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento, tci_bairro, tci_cep, tci_cidade, tci_UF " & _
            " FROM TAB_CONTRIBUINTE"
    If Im Then
        Sql = Sql & " WHERE TCI_IM='" & Inscricao & "'"
    Else
        Sql = Sql & " WHERE TCI_CGC_CPF='" & Inscricao & "'"
    End If
    
    If Bdados.AbreTabela(Sql, Rs) Then
        Inscricao = "" & Rs!tci_im
        Nome = "" & Rs!tci_nome
        If Not Endereco Is Nothing Then Endereco = "" & Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & ", " & Rs!tci_numero & " " & Rs!tci_complemento
        If Not Bairro Is Nothing Then Bairro = "" & Rs!tci_bairro
        If Not CEP Is Nothing Then CEP = "" & Rs!tci_cep
        If Not Cidade Is Nothing Then Cidade = "" & Rs!tci_cidade
        If Not UF Is Nothing Then UF = "" & Rs!tci_UF
        
        With Declaracao
            .tciNome = "" & Rs!tci_nome
            .tciEndereco = "" & Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & ", " & Rs!tci_numero & " " & Rs!tci_complemento
            .tciBairro = "" & Rs!tci_bairro
            .tciCEP = "" & Rs!tci_cep
            .tciCidade = "" & Rs!tci_cidade
            .tciUF = "" & Rs!tci_UF
        End With
        buscarContribuinte = True
    End If
    Bdados.FechaTabela Rs
End Function

Private Sub txtSTInscricao_LostFocus()
    
    On Error Resume Next
    If buscarContribuinte(txtSTInscricao, txtSTRazao, txtSTEndereco, txtSTBairro, txtSTCep, txtSTCidade, cboSTUf) Then
        If txtSTInscricao <> txtIM Then
            AliqISSST = Atividade.BuscaAliquotaAtividade(Bdados, txtSTInscricao, txtPeriodo, ISSSTFixo)
            Set Atividade = Nothing
            
            GerarIM = False
            txtSTNumNota.SetFocus
        Else
            Util.Avisa "Contribuinte de fato e de direito não podem ser o mesmo."
            txtSTInscricao = ""
            txtSTRazao = ""
            txtSTEndereco = ""
            txtSTBairro = ""
            txtSTCep = ""
            txtSTCidade = ""
            cboSTUf = ""
            txtSTInscricao.SetFocus
        End If
    Else
        txtSTRazao = ""
        txtSTEndereco = ""
        txtSTBairro = ""
        txtSTCep = ""
        txtSTCidade = ""
        cboSTUf = ""
        GerarIM = True
        txtSTRazao.SetFocus
    End If
End Sub

Private Sub prepararGridST()
    Dim UltimoIndice As Byte
    
    UltimoIndice = 8
    grdST.ColumnHeaders.Clear
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Inscricao": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Nota": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Emissao": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Recolhimento": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Valor": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "ICMS": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Tributavel": UltimoIndice = UltimoIndice + 1
    grdST.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Imposto": UltimoIndice = UltimoIndice + 1
End Sub


Private Sub prepararGridCancelada()
    Dim UltimoIndice As Byte
    
    UltimoIndice = 16
    grdCanceladas.ColumnHeaders.Clear
    grdCanceladas.ColumnHeaders.Add , "Item:" & UltimoIndice + 1, "Nota": UltimoIndice = UltimoIndice + 1
End Sub

Private Sub txtSTRecolhimento_LostFocus()
    If Trim$(txtSTRecolhimento) = "" Then txtSTRecolhimento = Date
End Sub

Private Sub txtSTValor_LostFocus()
    txtSTSaldo = Util.Nvl(txtSTValor, 0) - Util.Nvl(txtSTIcms, 0)
    If AliqISSST > 0 Then
        txtSTImposto = txtSTSaldo * AliqISSST
    Else
        txtSTImposto = ISSSTFixo
    End If
    
End Sub

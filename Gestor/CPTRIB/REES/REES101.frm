VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form REES101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REES101"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   3660
      Left            =   30
      TabIndex        =   4
      Tag             =   "Documento gerencial"
      Top             =   1725
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6456
      _Version        =   131082
      TabCount        =   4
      TabOrientation  =   2
      Tabs            =   "REES101.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   14
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "REES101.frx":00FF
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   3285
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5794
            Altura          =   1905
            Caption         =   " Nota Fiscal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtNota 
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
               Height          =   2955
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   300
               Width           =   9060
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "REES101.frx":0127
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   3285
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5794
            Altura          =   1905
            Caption         =   " Livro Fiscal (Modelos Diferentes)"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtLivro 
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
               Height          =   2955
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   300
               Width           =   9060
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   20
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "REES101.frx":014F
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   3285
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5794
            Altura          =   1905
            Caption         =   " Declaração Fiscal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtDeclaracao 
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
               Height          =   2955
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Top             =   300
               Width           =   9060
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   21
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "REES101.frx":0177
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   3285
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5794
            Altura          =   1905
            Caption         =   " Documento Gerencial"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtDocumento 
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
               Height          =   2955
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   8
               Top             =   300
               Width           =   9060
            End
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   16
      Top             =   6285
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8115
         TabIndex        =   13
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   5925
         TabIndex        =   11
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7020
         TabIndex        =   12
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1138
      Icone           =   "REES101.frx":019F
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   855
      Left            =   30
      TabIndex        =   18
      Top             =   5415
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1508
      Altura          =   1905
      Caption         =   " Dados do Responsável"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtCPF 
         Height          =   480
         Left            =   5910
         TabIndex        =   10
         Tag             =   "CPF Responsável"
         Top             =   315
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   847
         Caption         =   "CPF"
         Text            =   ""
         Formato         =   1
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   15
      End
      Begin VTOcx.txtVISUAL txtResp 
         Height          =   480
         Left            =   75
         TabIndex        =   9
         Tag             =   "Responsável"
         Top             =   315
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   847
         Caption         =   "Responsável:"
         Text            =   ""
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   50
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1020
      Left            =   30
      TabIndex        =   19
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1799
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   285
         Left            =   450
         TabIndex        =   2
         Top             =   690
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   503
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Tag             =   "Insc. Municipal"
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   3
         Top             =   375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
End
Attribute VB_Name = "REES101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GeraCod As New ContaCorrente

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub





Private Sub txtCPF_LostFocus()
      If Trim(txtCPF) = "" Then Exit Sub
    
    If txtCPF = "11111111111" Or txtCPF = "111.111.111-11" Or txtCPF = "22222222222" Or txtCPF = "222.222.222-22" Or txtCPF = "33333333333" Or txtCPF = "333.333.333-33" Or txtCPF = "44444444444" Or txtCPF = "444.444.444-44" Or txtCPF = "55555555555" Or txtCPF = "555.555.555-55" Or txtCPF = "66666666666" Or txtCPF = "666.666.666-66" Or txtCPF = "77777777777" Or txtCPF = "777.777.777-77" Or txtCPF = "88888888888" Or txtCPF = "888.888.888-88" Or txtCPF = "99999999999" Or txtCPF = "999.999.999-99" Or txtCPF = "00000000000" Or txtCPF = "000.000.000-00" Or txtCPF = "111.111.111-11" Or txtCPF = "11111111111" Then
        Util.Avisa "Valor do CPF inválido."
        txtCPF.SetFocus
    End If
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub

Private Sub cmdSalvar_Click()
    Dim camposRe As String
    Dim camposPr As String
    Dim ValoresRe As String
    Dim ValoresPr As String
    Dim Descricao As String
    Dim Codigo As Double
    Dim tipoProcesso As Integer
    Dim status As Integer
    
    Descricao = "INSTAURAÇÃO DE REGIME ESPECIAL"
    
    tipoProcesso = 3
    status = 1
    Codigo = GeraCod.GeraCodPagamento(40)
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtNota = "" And txtLivro = "" And txtDeclaracao = "" And txtDocumento = "" Then
        Avisa "Identifique um dos documentos"
        Exit Sub
    End If
    ' status_processo = 1 - aberto TipoProcesso Regime Especial = 3
    ' status regime = 3
    camposPr = "TPR_NUMERO_PROCESSO,TPR_INSCRICAO,TPR_DESCRICAO_PEDIDO,TPR_PEDIDO_REPR_PREPOSTO,TPR_PEDIDO_REPR_PREPOSTO_CPF,TPR_PEDIDO_DATA,TPR_TIPO_PROCESSO,TPR_STATUS"
    ValoresPr = Bdados.PreparaValor(Codigo, txtIm, Descricao, txtResp, txtCPF, Date, tipoProcesso, status)
    If Bdados.InsereDados("TAB_PROCESSO", ValoresPr, camposPr) Then
        camposRe = "TRE_TCI_IM,TRE_TPR_NUMERO_PROCESSO,TRE_Descricao_Declaracao,TRE_Descricao_Documento_Fiscal,TRE_Descricao_Nota_Fiscal,TRE_Livros_Fiscais_Modelos,TRE_STATUS_CONCLUSAO"
        ValoresRe = Bdados.PreparaValor(txtIm, Codigo, txtDeclaracao, txtDocumento, txtNota, txtLivro, 3)
        If Bdados.InsereDados("TAB_REGIME_ESPECIAL", ValoresRe, camposRe) Then
            Avisa "Dados gravados com sucesso"
            cmdLimpar_Click
        End If
    End If
End Sub



Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
    
End Sub






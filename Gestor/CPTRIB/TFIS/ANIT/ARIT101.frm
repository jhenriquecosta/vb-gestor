VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form ARIT101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ARIT101"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1138
      Icone           =   "ARIT101.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   9
      Top             =   7500
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   6555
         TabIndex        =   5
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   8505
         TabIndex        =   7
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   7530
         TabIndex        =   6
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   3105
      Left            =   45
      TabIndex        =   10
      Tag             =   "Documento gerencial"
      Top             =   3060
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5477
      _Version        =   131082
      TabCount        =   5
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "ARIT101.frx":0C0A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
         Height          =   2715
         Left            =   -99969
         TabIndex        =   34
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4789
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0D62
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   2715
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4789
            Altura          =   1905
            Caption         =   " Fundamentação Legal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtFundamentacao 
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
               Height          =   2370
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   36
               Top             =   300
               Width           =   9195
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2715
         Left            =   -99969
         TabIndex        =   11
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4789
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0D8A
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   2715
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4789
            Altura          =   1905
            Caption         =   " Atestado"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtAtestado 
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
               Height          =   2370
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Top             =   300
               Width           =   9195
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2715
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4789
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0DB2
         Begin VTOcx.fraVISUAL fraVISUAL9 
            Height          =   2715
            Left            =   -15
            TabIndex        =   43
            Top             =   0
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4789
            Altura          =   1905
            Caption         =   " Cumprimento das Exigências Legais"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtExigencias 
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
               Height          =   2370
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   44
               Top             =   300
               Width           =   9195
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   2715
         Left            =   -99969
         TabIndex        =   13
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4789
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0DDA
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   2715
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4789
            Altura          =   1905
            Caption         =   " Documentos Originários"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtDocOriginarios 
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
               Height          =   2370
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   300
               Width           =   9195
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   2715
         Left            =   -99969
         TabIndex        =   14
         Top             =   30
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   4789
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0E02
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   2715
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4789
            Altura          =   1905
            Caption         =   " Documentos Decorrentes"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VB.TextBox txtDocDecorrente 
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
               Height          =   2370
               Left            =   30
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   38
               Top             =   300
               Width           =   9195
            End
         End
      End
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1620
      Left            =   45
      TabIndex        =   15
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   2858
      Altura          =   1905
      Caption         =   " Qualificação do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   3270
         TabIndex        =   19
         Top             =   315
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3615
         TabIndex        =   18
         Top             =   315
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   540
         TabIndex        =   0
         Tag             =   "Ins. Municipal"
         Top             =   315
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   915
         TabIndex        =   17
         Top             =   615
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtNFiscalizacao 
         Height          =   285
         Left            =   5895
         TabIndex        =   1
         Tag             =   "Nº Fiscalização"
         Top             =   1275
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   503
         Caption         =   "Nº Fiscalização"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         AlinhamentoRotuloVertical=   0
         CorRotulo       =   16384
         CorTexto        =   4194304
         MaxLen          =   100
      End
      Begin VTOcx.cboVISUAL cboAtividade 
         Height          =   315
         Left            =   135
         TabIndex        =   16
         Tag             =   "Perfil Constitucional da Imunidade Tributária"
         Top             =   930
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   556
         Caption         =   "Atividade Principal"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.fraVISUAL fraHorario 
      Height          =   735
      Left            =   30
      TabIndex        =   20
      Top             =   2295
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   1296
      Altura          =   1905
      Caption         =   " Perfil Constitucional da Imunidade Tributária"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboPerfil 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Tag             =   "Perfil Constitucional da Imunidade Tributária"
         Top             =   345
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         Caption         =   "Perfil"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin ActiveTabs.SSActiveTabs SSActiveTabs1 
      Height          =   1305
      Left            =   45
      TabIndex        =   21
      Tag             =   "Documento gerencial"
      Top             =   6180
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   2302
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "ARIT101.frx":0E2A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel8 
         Height          =   915
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   1614
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0EBA
         Begin VTOcx.fraVISUAL fraVISUAL8 
            Height          =   855
            Left            =   15
            TabIndex        =   31
            Top             =   15
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " Dados do Responsável"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtMatricula 
               Height          =   480
               Left            =   7065
               TabIndex        =   33
               Tag             =   "Matrícula"
               Top             =   285
               Width           =   2085
               _ExtentX        =   3678
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
            Begin VTOcx.txtVISUAL txtVISUAL1 
               Height          =   480
               Left            =   75
               TabIndex        =   32
               Tag             =   "Responsável"
               Top             =   285
               Width           =   7005
               _ExtentX        =   12356
               _ExtentY        =   847
               Caption         =   "Responsável"
               Text            =   ""
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   50
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   915
         Left            =   -99969
         TabIndex        =   27
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   1614
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0EE2
         Begin VTOcx.fraVISUAL fraVISUAL7 
            Height          =   855
            Left            =   0
            TabIndex        =   28
            Top             =   15
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " Autoridade Fiscal"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboAutoridade 
               Height          =   315
               Left            =   150
               TabIndex        =   29
               Tag             =   "Autoridade Fiscal"
               Top             =   375
               Width           =   8955
               _ExtentX        =   15796
               _ExtentY        =   556
               Caption         =   "Autoridade Fiscal"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorRotulo       =   16384
               CorTexto        =   4194304
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   915
         Left            =   -99969
         TabIndex        =   25
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   1614
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0F0A
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   855
            Left            =   0
            TabIndex        =   26
            Top             =   15
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " Dados Responsável"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCpf 
               Height          =   480
               Left            =   6495
               TabIndex        =   4
               Tag             =   "Cpf"
               Top             =   285
               Width           =   2685
               _ExtentX        =   4736
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
               Left            =   60
               TabIndex        =   3
               Tag             =   "Responsável"
               Top             =   300
               Width           =   6390
               _ExtentX        =   11271
               _ExtentY        =   847
               Caption         =   "Responsável"
               Text            =   ""
               AlinhamentoRotulo=   1
               CorRotulo       =   4210752
               CorTexto        =   4194304
               MaxLen          =   50
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   22
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "ARIT101.frx":0F32
         Begin VTOcx.fraVISUAL fraVISUAL6 
            Height          =   3285
            Left            =   0
            TabIndex        =   23
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
               TabIndex        =   24
               Top             =   300
               Width           =   9060
            End
         End
      End
   End
End
Attribute VB_Name = "ARIT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Atarit As eAtarit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub
Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub
Private Sub cmdSalvar_Click()
    Dim sql As String
    Dim rs As VSRecordset
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    sql = "SELECT * FROM TAB_ATA_NIT WHERE tan_TCI_IM = '" & txtIm & "'"
    
    Screen.MousePointer = 11
    
    If Bdados.AbreTabela(sql, rs) Then
        Avisa "ATA-NIT já emitido para o contribuinte atual"
        Exit Sub
    End If
    
    With Atarit
        .CodPerfilConstitucional = cboPerfil.Coluna(1).VALOR
        .CumprimentoExigencias = txtExigencias
        .procedimento.Atributos.NumeroFiscalizacao = txtNFiscalizacao
        .procedimento.Atributos.NumeroProcedimento = ""
        .procedimento.Atributos.Contribuinte = txtIm
        .procedimento.Atributos.RepresentantePassivoNome = txtResp
        .procedimento.Atributos.RepresentantePassivoCPF = txtCpf
        .procedimento.Atributos.DocumentosOriginarios = txtDocOriginarios
        .procedimento.Atributos.DocumentosDecorrentes = txtDocDecorrente
        .procedimento.Atributos.FundamentacaoLegal = txtFundamentacao
        .procedimento.Atributos.DescricaoMotivos = txtAtestado
        .procedimento.Atributos.TipoProcedimento = tpProcedimentoATA_RIT
        .procedimento.Atributos.Autoridade.Atributos.Matricula = cboAutoridade.Coluna(1).VALOR
        If .Grava Then
            Avisa "Dados Salvos com sucesso."
            cmdLimpar_Click
        End If
    End With
    Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
    cboPerfil.Preencher Bdados, "Select * from Tab_Perfil_Constitucional"
    cboAutoridade.Preencher Bdados, ""
End Sub

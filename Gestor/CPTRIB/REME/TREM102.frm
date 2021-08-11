VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#1.1#0"; "VTControles.ocx"
Begin VB.Form TREM102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
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
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabParametro 
      Height          =   2640
      Left            =   30
      TabIndex        =   4
      Top             =   660
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   4657
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TREM102.frx":0000
      Images          =   "TREM102.frx":00B5
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2220
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   3916
         _Version        =   131082
         TabGuid         =   "TREM102.frx":1188
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            Caption         =   "Data Vencimento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1155
            Index           =   3
            Left            =   1170
            TabIndex        =   31
            Top             =   990
            Width           =   3345
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "AANNN"
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
               Height          =   195
               Index           =   9
               Left            =   1800
               TabIndex        =   38
               Top             =   900
               Width           =   1425
            End
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "DDMMAAAA"
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
               Height          =   195
               Index           =   7
               Left            =   1800
               TabIndex        =   37
               Top             =   690
               Width           =   1425
            End
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "DDMMAA"
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
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   36
               Top             =   240
               Width           =   1425
            End
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "AAMMDD"
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
               Height          =   195
               Index           =   3
               Left            =   30
               TabIndex        =   35
               Top             =   465
               Width           =   1425
            End
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "AAAAMMDD"
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
               Height          =   195
               Index           =   4
               Left            =   30
               TabIndex        =   34
               Top             =   690
               Width           =   1425
            End
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "ANNN"
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
               Height          =   195
               Index           =   5
               Left            =   1800
               TabIndex        =   33
               Top             =   240
               Width           =   1425
            End
            Begin VB.OptionButton optVencto 
               Appearance      =   0  'Flat
               Caption         =   "NNNA"
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
               Height          =   195
               Index           =   6
               Left            =   1800
               TabIndex        =   32
               Top             =   465
               Width           =   1425
            End
         End
         Begin VTOcx.txtVISUAL txtNumConvenio 
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Tag             =   "Numero Convenio"
            Top             =   30
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
            Caption         =   "Nº Convenio"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   6
         End
         Begin VTOcx.txtVISUAL txtFebraban 
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Tag             =   "Identificao FEBRABAN"
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            Caption         =   "FEBRABAN"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   8
         End
         Begin VTOcx.txtVISUAL txtLayout 
            Height          =   285
            Left            =   570
            TabIndex        =   30
            Top             =   690
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            Caption         =   "Layout"
            Text            =   ""
            MaxLen          =   7
            RetirarMascara  =   0   'False
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2220
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   3916
         _Version        =   131082
         TabGuid         =   "TREM102.frx":11B0
         Begin VTOcx.txtVISUAL txtCNPJ 
            Height          =   285
            Left            =   390
            TabIndex        =   10
            Tag             =   "CNPJ Convenente"
            Top             =   30
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            Caption         =   "CNPJ"
            Text            =   ""
            Formato         =   2
            Restricao       =   2
            MaxLen          =   18
         End
         Begin VTOcx.txtVISUAL txtConvenente 
            Height          =   285
            Left            =   330
            TabIndex        =   11
            Tag             =   "Nome Convenente"
            Top             =   390
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   503
            Caption         =   "Nome"
            Text            =   ""
            MaxLen          =   50
         End
         Begin VTOcx.txtVISUAL txtEndereco 
            Height          =   285
            Left            =   30
            TabIndex        =   12
            Tag             =   "Endereco Convenente"
            Top             =   750
            Width           =   5625
            _ExtentX        =   9922
            _ExtentY        =   503
            Caption         =   "Endereço"
            Text            =   ""
            MaxLen          =   35
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   285
            Left            =   300
            TabIndex        =   13
            Tag             =   "Bairro Convenente"
            Top             =   1110
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   503
            Caption         =   "Bairro"
            Text            =   ""
            MaxLen          =   30
         End
         Begin VTOcx.txtVISUAL txtCEP 
            Height          =   285
            Left            =   3810
            TabIndex        =   14
            Tag             =   "CEP Convenente"
            Top             =   1110
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   503
            Caption         =   "CEP"
            Text            =   ""
            Formato         =   4
            Restricao       =   2
            MaxLen          =   9
         End
         Begin VTOcx.txtVISUAL txtCidade 
            Height          =   285
            Left            =   210
            TabIndex        =   15
            Tag             =   "Cidade Convenente"
            Top             =   1470
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   503
            Caption         =   "Cidade"
            Text            =   ""
            MaxLen          =   30
         End
         Begin VTOcx.txtVISUAL txtUF 
            Height          =   285
            Left            =   5040
            TabIndex        =   16
            Tag             =   "UF Convenente"
            Top             =   1470
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Caption         =   "UF"
            Text            =   ""
            Restricao       =   1
            MaxLen          =   2
         End
         Begin VTOcx.txtVISUAL txtUnidade 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Tag             =   "Nome Convenente"
            Top             =   1830
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   503
            Caption         =   "Unidade"
            Text            =   ""
            MaxLen          =   40
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   2220
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   3916
         _Version        =   131082
         TabGuid         =   "TREM102.frx":11D8
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            Caption         =   "Receber após vencto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   705
            Index           =   2
            Left            =   1860
            TabIndex        =   26
            Top             =   1380
            Width           =   2175
            Begin VB.OptionButton optReceberVencto 
               Appearance      =   0  'Flat
               Caption         =   "Sim"
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
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   28
               Top             =   240
               Width           =   1665
            End
            Begin VB.OptionButton optReceberVencto 
               Appearance      =   0  'Flat
               Caption         =   "Não"
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
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   27
               Top             =   465
               Width           =   1665
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            Caption         =   "Formulário"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1365
            Index           =   0
            Left            =   60
            TabIndex        =   20
            Top             =   0
            Width           =   3915
            Begin VB.OptionButton optTipoFormulario 
               Appearance      =   0  'Flat
               Caption         =   "Guia com pagamento único (2 vias envelopado)"
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
               Height          =   195
               Index           =   4
               Left            =   30
               TabIndex        =   25
               Top             =   1140
               Width           =   3855
            End
            Begin VB.OptionButton optTipoFormulario 
               Appearance      =   0  'Flat
               Caption         =   "Guia com pagamento único (4 vias)"
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
               Height          =   195
               Index           =   3
               Left            =   30
               TabIndex        =   24
               Top             =   915
               Width           =   3855
            End
            Begin VB.OptionButton optTipoFormulario 
               Appearance      =   0  'Flat
               Caption         =   "Guia com pagamento único (2 vias)"
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
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   23
               Top             =   690
               Width           =   3855
            End
            Begin VB.OptionButton optTipoFormulario 
               Appearance      =   0  'Flat
               Caption         =   "Guia com parcelas"
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
               Height          =   195
               Index           =   0
               Left            =   30
               TabIndex        =   22
               Top             =   465
               Width           =   3855
            End
            Begin VB.OptionButton optTipoFormulario 
               Appearance      =   0  'Flat
               Caption         =   "Carnê"
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
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   21
               Top             =   240
               Width           =   3855
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            Caption         =   "Remessa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   705
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   1380
            Width           =   1725
            Begin VB.OptionButton optTipoRemessa 
               Appearance      =   0  'Flat
               Caption         =   "Produção"
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
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   19
               Top             =   465
               Width           =   1665
            End
            Begin VB.OptionButton optTipoRemessa 
               Appearance      =   0  'Flat
               Caption         =   "Teste"
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
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   18
               Top             =   240
               Width           =   1665
            End
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabCabecalho 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   1138
      Formulario      =   "CODIGO"
      Icone           =   "TREM102.frx":1200
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   3330
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   926
      CorFundo        =   -2147483632
      CorFrente       =   -2147483633
      Begin VTOcx.cmdVISUAL cmdGravar 
         Height          =   405
         Left            =   3990
         TabIndex        =   0
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   714
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   4980
         TabIndex        =   1
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
End
Attribute VB_Name = "TREM102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intTipoFormulario As Integer
Private intTipoRemessa As Integer
Private intReceberVencto As Integer
Private intFormatoVencto As Integer

Private Sub cmdGravar_Click()
    Dim Campos As String, Valores As String

    If Edita.CriticaCampos(Me) Then
        Campos = "TPR_NUMERO_CONVENIO, TPR_IDENTIFICACAO_FEBRABAN, TPR_CNPJ_CONVENENTE, " & _
                " TPR_CONVENENTE, TPR_ENDERECO_CONVENENTE, TPR_CEP_CONVENENTE, TPR_CIDADE_CONVENENTE, " & _
                " TPR_BAIRRO_CONVENENTE, TPR_UF_CONVENENTE, TPR_UNIDADE_CONVENENTE, TPR_TIPO_FORMULARIO, " & _
                " TPR_TIPO_REMESSA, TPR_RECEBER_VENCIDO, TPR_LAYOUT, TPR_FORMATO_VENCTO"
        Valores = Bdados.PreparaValor(txtNumConvenio, txtFebraban, Bdados.Converte(Edita.TiraTudo(txtCNPJ), TCTexto), _
                txtConvenente, Edita.TiraTudo(txtEndereco), Edita.TiraTudo(txtCEP), txtCidade, _
                txtBairro, txtUF, txtUnidade, intTipoFormulario, intTipoRemessa, intReceberVencto, _
                Bdados.Converte(txtLayout, TCTexto), intFormatoVencto)
        Bdados.DeletaDados "TAB_PARAMETRO_REMESSA"
        If Bdados.InsereDados("TAB_PARAMETRO_REMESSA", Valores, Campos) Then
            Informa "Parâmetros registrados com sucesso."
            cmdSair.SetFocus
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabCabecalho.Exibir Bdados, Me.Name, App.Path
    rodRodape.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Exibir
End Sub

Private Sub Exibir()
    Dim rs As VSRecordset
    
    Edita.LimpaCampos Me
    If Bdados.AbreTabela("SELECT * FROM TAB_PARAMETRO_REMESSA", rs) Then
        txtNumConvenio = "" & rs!TPR_NUMERO_CONVENIO
        txtFebraban = "" & rs!TPR_IDENTIFICACAO_FEBRABAN
        txtLayout = "" & rs!TPR_LAYOUT
        txtCNPJ = "" & rs!TPR_CNPJ_CONVENENTE
        txtConvenente = "" & rs!TPR_CONVENENTE
        txtEndereco = "" & rs!TPR_ENDERECO_CONVENENTE
        txtCEP = "" & rs!TPR_CEP_CONVENENTE
        txtCidade = "" & rs!TPR_CIDADE_CONVENENTE
        txtBairro = "" & rs!TPR_BAIRRO_CONVENENTE
        txtUF = "" & rs!TPR_UF_CONVENENTE
        txtUnidade = "" & rs!TPR_UNIDADE_CONVENENTE
        If "" & rs!TPR_TIPO_FORMULARIO <> "" Then optTipoFormulario(rs!TPR_TIPO_FORMULARIO).Value = True
        If "" & rs!TPR_TIPO_REMESSA <> "" Then optTipoRemessa(rs!TPR_TIPO_REMESSA).Value = True
        If "" & rs!TPR_RECEBER_VENCIDO <> "" Then optReceberVencto(rs!TPR_RECEBER_VENCIDO + 1).Value = True
        If "" & rs!TPR_FORMATO_VENCTO <> "" Then optVencto(rs!TPR_FORMATO_VENCTO).Value = True
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub optTipoFormulario_Click(Index As Integer)
    Dim Opcao As OptionButton
    
    For Each Opcao In optTipoFormulario
        Opcao.Font.Bold = False
    Next
    
    optTipoFormulario(Index).Font.Bold = True
    intTipoFormulario = Index
End Sub

Private Sub optTipoRemessa_Click(Index As Integer)
    Dim Opcao As OptionButton
    
    For Each Opcao In optTipoRemessa
        Opcao.Font.Bold = False
    Next
    
    optTipoRemessa(Index).Font.Bold = True
    intTipoRemessa = Index
End Sub

Private Sub optReceberVencto_Click(Index As Integer)
    Dim Opcao As OptionButton
    
    For Each Opcao In optReceberVencto
        Opcao.Font.Bold = False
    Next
    
    optReceberVencto(Index).Font.Bold = True
    intReceberVencto = Index - 1
End Sub

Private Sub optVencto_Click(Index As Integer)
    Dim Opcao As OptionButton
    
    For Each Opcao In optVencto
        Opcao.Font.Bold = False
    Next
    
    optVencto(Index).Font.Bold = True
    intFormatoVencto = Index
End Sub


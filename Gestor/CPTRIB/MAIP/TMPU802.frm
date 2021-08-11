VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TMPU802 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdTrecho 
      Height          =   2415
      Left            =   60
      TabIndex        =   27
      Top             =   3060
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4260
      CorTitulo       =   16711680
      CorCaption      =   16777215
      OcultarRodape   =   -1  'True
   End
   Begin VB.ComboBox cboRelatorio 
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
      ItemData        =   "TMPU802.frx":0000
      Left            =   3435
      List            =   "TMPU802.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2205
      Width           =   3075
   End
   Begin Threed.SSFrame fra 
      Height          =   1485
      Index           =   0
      Left            =   45
      TabIndex        =   17
      Top             =   675
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   2619
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
      Begin VB.TextBox txtDist 
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
         Left            =   870
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "Trecho"
         Top             =   990
         Width           =   510
      End
      Begin VB.TextBox txtCodBairro 
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "Bairro"
         Top             =   990
         Width           =   600
      End
      Begin VB.TextBox txtLogrInicial 
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
         Left            =   3195
         MaxLength       =   10
         TabIndex        =   6
         Top             =   990
         Width           =   1035
      End
      Begin VB.TextBox txtLogrFinal 
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
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   7
         Top             =   990
         Width           =   1035
      End
      Begin VB.TextBox txtSetor 
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
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "Setor"
         Top             =   990
         Width           =   495
      End
      Begin VB.TextBox txtQuadra 
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
         Left            =   2550
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "Quadra"
         Top             =   990
         Width           =   585
      End
      Begin VB.TextBox txtNumTrecho 
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
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Trecho"
         Top             =   990
         Width           =   720
      End
      Begin VB.TextBox txtValor 
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
         Left            =   5445
         MaxLength       =   6
         TabIndex        =   8
         Top             =   990
         Width           =   915
      End
      Begin VB.TextBox txtCodLogr 
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
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Cod Logr"
         Top             =   390
         Width           =   825
      End
      Begin VB.TextBox txtBairro 
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
         Left            =   4110
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "Valor"
         Top             =   390
         Width           =   2265
      End
      Begin VB.TextBox txtLogr 
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
         Left            =   990
         MaxLength       =   50
         TabIndex        =   15
         Tag             =   "Valor"
         Top             =   390
         Width           =   3030
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   150
         Width           =   765
         _ExtentX        =   1349
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
         Caption         =   "Cod Logr"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Command1"
         Default         =   -1  'True
         Height          =   255
         Left            =   3090
         TabIndex        =   19
         Top             =   2550
         Width           =   375
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   20
         Top             =   780
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Trecho"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   7
         Left            =   1395
         TabIndex        =   21
         Top             =   750
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Set."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   8
         Left            =   2565
         TabIndex        =   22
         Top             =   750
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Qd."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   9
         Left            =   3210
         TabIndex        =   23
         Top             =   750
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Logr Inicial"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   10
         Left            =   4335
         TabIndex        =   24
         Top             =   750
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Logr Final"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   0
         Left            =   1935
         TabIndex        =   25
         Top             =   750
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   196610
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
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   3
         Left            =   5430
         TabIndex        =   26
         Top             =   750
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Valor"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   4
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   225
         Index           =   4
         Left            =   870
         TabIndex        =   28
         Top             =   780
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   196610
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
         Caption         =   "Dist."
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   1
         Alignment       =   0
         RoundedCorners  =   0   'False
      End
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   5385
      TabIndex        =   10
      Top             =   2595
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
      Left            =   4185
      TabIndex        =   9
      Top             =   2595
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdImprime 
      Height          =   375
      Left            =   1785
      TabIndex        =   12
      Top             =   2595
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Imprimir"
      Acao            =   4
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdLimpar 
      Height          =   375
      Left            =   585
      TabIndex        =   11
      Top             =   2595
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin VTOcx.cmdVISUAL cmdExcluir 
      Height          =   375
      Left            =   2985
      TabIndex        =   13
      Top             =   2595
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777088
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   1138
      Icone           =   "TMPU802.frx":0040
   End
End
Attribute VB_Name = "TMPU802"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto
Dim Click As Boolean
Dim Relatorio As VSRelatorio

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Function VerificaTEMInfraEstrutura(CodLogr As String, NumTrecho As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select tdl_tlg_cod_logradouro,tdl_num_trecho, tdl_tgl_cod_grupo" & _
            " From tab_detalhe_logradouro" & _
            " where tdl_tlg_cod_logradouro = '" & CodLogr & "' and" & _
            " tdl_num_trecho = '" & NumTrecho & "'"
    If Bdados.AbreTabela(Sql, rs) Then VerificaTEMInfraEstrutura = True
End Function

Private Sub cmdExcluir_Click()
    If Trim(txtNumTrecho) = "" Then Exit Sub
    If Trim(txtCodLogr) = "" Then Exit Sub
    If Util.Confirma("Confirma exclusão do trecho " & txtNumTrecho & " ?") Then
        Screen.MousePointer = 11
        If VerificaTEMInfraEstrutura(txtCodLogr, txtNumTrecho) = True Then
            Util.Informa ("Existe infra-estrutura cadastrada para esse trecho. Não é possivel excluí-lo.")
            txtNumTrecho.SetFocus
            Screen.MousePointer = 0
        Else
            If Bdados.DeletaDados("tab_trecho", "TTC_TLG_COD_LOGRADOURO = '" & txtCodLogr & "' and TTC_COD_TRECHO = '" & txtNumTrecho & "'") Then
                Util.Informa "Trecho " & txtNumTrecho & " excluído."
                cmdLimpar_Click
                Screen.MousePointer = 0
            End If
        End If
    End If
End Sub

Private Sub cmdImprime_Click()
    On Error GoTo trata
Set Relatorio = New VSRelatorio
    With Relatorio
        
    Select Case cboRelatorio.ListIndex
        Case 0 'PENDENTES
            If Not .DefinirArquivo(Bdados, App.Path & "\TTrechorpt.rpt") Then Exit Sub
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        
        Case 1 'TODOS
            If .DefinirArquivo(Bdados, App.Path & "\TTrechoDetalhado.rpt") Then
                .Formulas "PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
                .Formulas "SECRETARIA", Temp.PegaParametro(Bdados, "SEMFAZ")
                .Formulas "SETOR", Temp.PegaParametro(Bdados, "SETOR")
            End If
        Case 2 'INFRA
            If .DefinirArquivo(Bdados, App.Path & "\TInfraEstrutura.rpt") Then
                .Formulas "PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
                .Formulas "SECRETARIA", Temp.PegaParametro(Bdados, "SEMFAZ")
                .Formulas "SETOR", Temp.PegaParametro(Bdados, "SETOR")
                Dim TODASELECT As String
                TODASELECT = IIf(Trim(txtCodBairro) <> "", "{TAB_BAIRRO.TBA_COD_BAIRRO} = " & txtCodBairro, "")
                If Trim(txtCodLogr) <> "" Then
                    .Selecao = TODASELECT & IIf(Trim(txtCodBairro) <> "", " and {VIS_TRECHO.TTC_TLG_COD_LOGRADOURO} = '" & txtCodLogr & "'", "{VIS_TRECHO.TTC_TLG_COD_LOGRADOURO} = '" & txtCodLogr & "'")
                End If
                .Arvore = False
                .Visualizar
            End If
            DoEvents
            Set Relatorio = Nothing
            Exit Sub
        Case 3 'TRECHO
            If .DefinirArquivo(Bdados, App.Path & "\TTrechoSetor.rpt") Then
                .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name, Horizontal
                .Titulo = "Trechos do Setor " & IIf(Trim(txtSetor) <> "", txtSetor, " - TODOS")
                If Trim(txtSetor) <> "" Then .Selecao = "cdbl({TAB_TRECHO.TTC_SETOR}) = cdbl(" & txtSetor & ")"
                .Arvore = False
                .Visualizar
                DoEvents
            End If
            Set Relatorio = Nothing
            Exit Sub
        Case Else
            Avisa "Informe o tipo do relatório."
            cboRelatorio.SetFocus
            Exit Sub
    End Select
    .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
    .Titulo = "Componentes do Cadastro Imobiliário"
    If Trim(txtCodLogr) <> "" Then .Selecao = "{TAB_LOGRADOURO.tlg_cod_logradouro} = '" & txtCodLogr & "'"
    .Arvore = False
    .Visualizar
    DoEvents
    End With
    Set Relatorio = Nothing
trata:
    Erro Err.Description
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdTrecho.ListItems.Clear
    txtCodLogr.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim CodLogradouro As Long
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Campos = "TTC_TLG_COD_LOGRADOURO,TTC_COD_TRECHO,TTC_DISTRITO,TTC_SETOR,TTC_QUADRA,TTC_TBA_COD_BAIRRO,TTC_LOGR_INICIAL,TTC_LOGR_FINAL,TTC_VALOR"
    Valores = Bdados.PreparaValor(CDbl(txtCodLogr), txtNumTrecho, Bdados.Converte(Format(txtDist, "00"), tctexto), Bdados.Converte(Format(txtSetor, "00"), tctexto), Bdados.Converte(Format(txtQuadra, "000"), tctexto), CDbl(txtCodBairro), CDbl(Nvl(txtLogrInicial, 0)), CDbl(Nvl(txtLogrFinal, 0)), Bdados.Converte(Nvl(Trim(txtValor), 0), TCDuplo))
    If Bdados.GravaDados("TAB_TRECHO", Valores, Campos, "TTC_TLG_COD_LOGRADOURO='" & txtCodLogr & "' and TTC_COD_TRECHO='" & txtNumTrecho & "'") Then
        Informa "Trecho gravado."
        txtNumTrecho = CStr(CInt(Nvl(Mid(txtNumTrecho, 1, Len(txtNumTrecho) - 1), 0)) + 1) & Right(txtNumTrecho, 1)
        txtQuadra = ""
        txtLogrInicial = ""
        txtLogrFinal = ""
        'cmdLimpar_Click
        txtSetor = ""
        txtDist = ""
        txtCodBairro = ""
        txtValor = ""
        Screen.MousePointer = 0
        txtNumTrecho.SetFocus
    Else
        Erro "Problemas ao gravar o trecho."
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub Form_Activate()
    Set cadastro = New VSImposto
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    
End Sub


Private Sub grdTrecho_DblClick()
    If grdTrecho.SelectedItem Is Nothing Then Exit Sub
    CarregaDetalhesTrecho
    txtNumTrecho.SetFocus
End Sub

Private Sub txtCodBairro_Change()
    If Len(txtCodBairro) = txtCodBairro.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCodBairro_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtCodLogr_LostFocus()
    Dim rs As VSRecordset
    Dim Sql As String
    Dim codTrecho As String
    Dim CodLogradouro As String
    If Trim(txtCodLogr) <> "" Then
        Sql = "Select TTL_NOME,tlg_nome,tba_nome from tab_logradouro, tab_tipo_logr,tab_bairro where tlg_cod_logradouro='" & CDbl(txtCodLogr) & "' and tlg_ttl_cod_tip_logr = TTL_COD_TIP_LOGR and tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " AND TLG_TBA_COD_BAIRRO = TBA_COD_BAIRRO and tlg_tmu_cod_municipio = tba_tmu_cod_municipio "
        If Bdados.AbreTabela(Sql, rs) Then
            CodLogradouro = txtCodLogr
            Edita.LimpaCampos Me
            txtCodLogr = CodLogradouro
            txtLogr = rs(0) & " " & rs(1)
            txtBairro = rs(2)
            PreencherGridTrecho grdTrecho, txtCodLogr
        Else
            Avisa "Código de logradouro inexistente."
            grdTrecho.Preencher Bdados, ""
            txtLogr.SetFocus
        End If
        Bdados.FechaTabela rs
    End If
End Sub

Private Sub txtDist_Change()
    If Len(txtDist) = txtDist.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDist_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLogrFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtLogrInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtNumTrecho_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumTrecho_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    If Trim(txtCodLogr) = "" Then Exit Sub
    If Trim(txtNumTrecho) <> "" Then
        Sql = "Select * from tab_trecho where (not ttc_setor is null ) and TTC_TLG_COD_LOGRADOURO ='" & _
            CDbl(txtCodLogr) & "' and TTC_COD_TRECHO ='" & txtNumTrecho & "'"
        If Bdados.AbreTabela(Sql, rs, Registros) Then
            txtDist = "" & rs!TTC_DISTRITO
            txtSetor = "" & rs!TTC_SETOR
            txtQuadra = "" & rs!TTC_QUADRA
            txtCodBairro = "" & rs!TTC_TBA_COD_BAIRRO
            txtLogrInicial = "" & rs!TTC_LOGR_INICIAL
            txtLogrFinal = "" & rs!TTC_LOGR_FINAL
            txtValor = Format("" & rs!TTC_VALOR, Const_Monetario)
        End If
        Bdados.FechaTabela rs
    End If
    
End Sub

Private Sub txtQuadra_Change()
     If Len(txtQuadra) = txtQuadra.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtSetor_Change()
    If Len(txtSetor) = txtSetor.MaxLength Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSetor_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
Private Function PreencherGridTrecho(grd As Object, CodLogr As String) As Boolean
    Dim Sql As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "select TTC_TLG_COD_LOGRADOURO as [Cod. Logr.]," & _
               "  TTC_COD_TRECHO as [Cod. Trecho]," & _
               " TTC_DISTRITO as Distrito," & _
               " TTC_SETOR as [Set.]," & _
               " TTC_QUADRA as [Qd.]," & _
               " TTC_TBA_COD_BAIRRO as Bairro," & _
               " TTC_LOGR_INICIAL as [Logr. Inicial]," & _
               " TTC_LOGR_FINAL as [Logr. Final]," & _
               " TTC_VALOR As Valor" & _
               " From tab_trecho" & _
               " where TTC_TLG_COD_LOGRADOURO = '" & CodLogr & "' order by ttc_seq_trecho asc"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "select TTC_TLG_COD_LOGRADOURO as Cod_Logr," & _
               "  TTC_COD_TRECHO as Cod_Trecho," & _
               " TTC_DISTRITO as Distrito," & _
               " TTC_SETOR as Setor," & _
               " TTC_QUADRA as Qd," & _
               " TTC_TBA_COD_BAIRRO as Bairro," & _
               " TTC_LOGR_INICIAL as Logr_Inicial," & _
               " TTC_LOGR_FINAL as Logr_Final," & _
               " TTC_VALOR As Valor" & _
               " From tab_trecho" & _
               " where TTC_TLG_COD_LOGRADOURO = '" & CodLogr & "' order by ttc_seq_trecho asc"
    End If
    grd.Preencher Bdados, Sql
End Function

Private Sub CarregaDetalhesTrecho()
    txtCodLogr = grdTrecho.SelectedItem
    txtNumTrecho = grdTrecho.SelectedItem.SubItems(1)
    txtSetor = grdTrecho.SelectedItem.SubItems(3)
    txtCodBairro = grdTrecho.SelectedItem.SubItems(3)
    txtQuadra = grdTrecho.SelectedItem.SubItems(4)
    txtLogrInicial = grdTrecho.SelectedItem.SubItems(6)
    txtLogrFinal = grdTrecho.SelectedItem.SubItems(7)
    txtValor = grdTrecho.SelectedItem.SubItems(8)
    txtDist = grdTrecho.SelectedItem.SubItems(2)
End Sub

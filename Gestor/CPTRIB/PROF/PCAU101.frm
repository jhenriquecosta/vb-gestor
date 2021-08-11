VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form PCAU101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario-Permissao de Usuarios"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PCAU101.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabAcessos 
      Height          =   4590
      Left            =   60
      TabIndex        =   9
      Top             =   1800
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   8096
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Tabs            =   "PCAU101.frx":08CA
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4200
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   7408
         _Version        =   131082
         TabGuid         =   "PCAU101.frx":0949
         Begin MSComctlLib.TreeView treeSistema 
            Height          =   4095
            Left            =   60
            TabIndex        =   2
            Top             =   45
            Width           =   7290
            _ExtentX        =   12859
            _ExtentY        =   7223
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   538
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   6
            Checkboxes      =   -1  'True
            HotTracking     =   -1  'True
            ImageList       =   "imlMenu"
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4200
         Left            =   -99969
         TabIndex        =   11
         Top             =   30
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   7408
         _Version        =   131082
         TabGuid         =   "PCAU101.frx":0971
         Begin MSComctlLib.TreeView treeLotacoes 
            Height          =   4095
            Left            =   60
            TabIndex        =   6
            Top             =   45
            Width           =   7290
            _ExtentX        =   12859
            _ExtentY        =   7223
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   538
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   6
            Checkboxes      =   -1  'True
            HotTracking     =   -1  'True
            ImageList       =   "imlMenu"
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVisual 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   7
      Top             =   6540
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   926
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   405
         Left            =   4410
         TabIndex        =   3
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   5445
         TabIndex        =   4
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   6480
         TabIndex        =   5
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1080
      Left            =   45
      TabIndex        =   8
      Top             =   690
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1905
      Altura          =   1905
      Caption         =   " Usuário"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtNome 
         Height          =   285
         Left            =   345
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   697
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   503
         Caption         =   "Nome"
         Text            =   ""
         Restricao       =   1
         MaxLen          =   200
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Tag             =   "Código"
         Top             =   360
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   503
         Caption         =   "Codigo"
         Text            =   ""
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1138
      Icone           =   "PCAU101.frx":0999
   End
End
Attribute VB_Name = "PCAU101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Pode As Boolean
Dim Lotacao As pLotacao
Dim Acesso As pAcesso


Private Sub GravaDados(User As String)
    On Error GoTo Trata
    Dim Rs As Object
    Dim Sql As String
    Dim I As Integer
    Dim Condicao As String
    Dim Texto As String
    Dim Valores As String
    Dim Campos As String
    Dim Sistema As String
    
    BDados.DeletaDados "TAB_ACESSO_USUARIO", "TAU_TUS_COD_USUARIO = '" & User & "'"
    
    For I = 1 To treeSistema.Nodes.Count
        Texto = treeSistema.Nodes.Item(I).Key
        If IsNumeric(Right(Texto, 3)) Then
            If treeSistema.Nodes.Item(I).Checked Then
                Sql = "SELECT TMO_TSI_COD_SISTEMA FROM TAB_MODULO WHERE TMO_COD_MODULO = '" & Left(Texto, 4) & "'"
                If BDados.AbreTabela(Sql, Rs) Then
                    Sistema = Rs(0)
                
                    Valores = BDados.PreparaValor(User, Sistema, Left(Texto, 4), Right(Texto, 3))
                    Campos = "TAU_TUS_COD_USUARIO,TAU_TSI_COD_SISTEMA," & _
                    "TAU_TMO_COD_MODULO,TAU_TFO_COD_FORMULARIO"
                    Call BDados.InsereDados("TAB_ACESSO_USUARIO", Valores, Campos)
                End If
            End If
        End If
    Next
    Call Util.Informa("Registro gravado com sucesso.")
    
    Exit Sub
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    Set Acesso = New pAcesso
        Acesso.ExibeSistema treeSistema, tpAcesso
    Set Acesso = Nothing
    
    Set Lotacao = New pLotacao
        Lotacao.ExibirEstrutura treeLotacoes
    Set Lotacao = Nothing
    
    treeSistema.Enabled = False
    tabAcessos.Tabs(1).Selected = True
    Pode = False
    txtCodigo.SetFocus
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
    
    If Pode Then
        Screen.MousePointer = 11
        
        Set Acesso = New pAcesso
            Acesso.GravarAcessoSistema treeSistema, txtCodigo, tpAcesso
            Acesso.GravarAcessoLotacao treeLotacoes, txtCodigo
            Acesso.ExibeSistema treeSistema, tpAcesso
        Set Acesso = Nothing
        
        Call cmdLimpar_Click
        Screen.MousePointer = 0
    Else
        Util.Avisa "Não existe alteração para ser Gravada."
        tabAcessos.Tabs(1).Selected = True
        treeSistema.SetFocus
    End If
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    cabVisual.Exibir BDados, Me.Name, App.Path
    rodVisual.Exibir BDados, Me.Name, App.Major, App.Minor, App.Revision
    Me.Show
    Set Acesso = New pAcesso
        Acesso.ExibeSistema treeSistema, tpAcesso
    Set Acesso = Nothing
    
    Set Lotacao = New pLotacao
        Lotacao.ExibirEstrutura treeLotacoes
    Set Lotacao = Nothing
    
    
    Screen.MousePointer = 0
    Exit Sub
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub treeLotacoes_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo Trata
    
    Util.Marcar Node, Node, Node.Checked
    ' LINHA COMENTADA PARA PERMITIR QUE ALGUÉM TENHA ACESSO SOMENTE PARA DENTRO DO NÓ
    'Util.Integridade treeLotacoes.Nodes.Item(1)
    Pode = True

Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub


Private Sub treeSistema_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error GoTo Trata
    
    Util.Marcar Node, Node, Node.Checked
    Util.Integridade treeSistema.Nodes.Item(1)
    Pode = True
    
Trata:
    If ERR.Number <> 0 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    
    If txtCodigo = "" Then Exit Sub
    Set Acesso = New pAcesso
        Acesso.ExibeSistema treeSistema, tpAcesso
    Set Acesso = Nothing
    
    Set Lotacao = New pLotacao
'        Lotacao.ExibirEstrutura treeLotacoes
    Set Lotacao = Nothing
    
    txtNome = UCase(Seguranca.ExisteUsuario(BDados, txtCodigo))
    If Trim(txtNome) = "" Then
        Util.Informa "Usuário não Cadastrado."
        txtCodigo = ""
        txtCodigo.SetFocus
    Else
        Screen.MousePointer = 11
        treeSistema.Enabled = True
        
        Set Acesso = New pAcesso
        If Acesso.MarcaAcessos(treeSistema, txtCodigo, taSistema, tpAcesso) Then
'            Acesso.MarcaAcessos treeLotacoes, txtCodigo, taLotacao
        Else
            LimpaCampos Me
            DoEvents
        End If
        Set Acesso = Nothing
        
        Util.Integridade treeSistema.Nodes.Item(1)
        treeSistema.Refresh
        tabAcessos.Tabs(1).Selected = True
        treeSistema.SetFocus
        Screen.MousePointer = 0
    End If
End Sub


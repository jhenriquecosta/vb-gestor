VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form REES104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REES104"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
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
      TabCount        =   2
      TabOrientation  =   2
      Tabs            =   "REES104.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3270
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "REES104.frx":008C
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   3285
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5794
            Altura          =   1905
            Caption         =   " Despacho"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtDesp 
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
               Tag             =   "Despacho"
               Top             =   300
               Width           =   9060
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "REES104.frx":00B4
         Begin VTOcx.grdVISUAL grdDados 
            Height          =   3180
            Left            =   15
            TabIndex        =   16
            Top             =   90
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   5609
            CorBorda        =   32768
            Caption         =   "Processos em Andamento"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            CheckBox        =   -1  'True
            MarcaUnico      =   -1  'True
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   11
      Top             =   6345
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8115
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   6
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
      TabIndex        =   12
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1138
      Icone           =   "REES104.frx":00DC
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   855
      Left            =   45
      TabIndex        =   13
      Top             =   5415
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1508
      Altura          =   1905
      Caption         =   " Autoridade Fiscal"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtData 
         Height          =   480
         Left            =   7080
         TabIndex        =   19
         Tag             =   "Data"
         Top             =   315
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   847
         Caption         =   "Data"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   15
      End
      Begin VTOcx.txtVISUAL txtMatricula 
         Height          =   480
         Left            =   5055
         TabIndex        =   18
         Tag             =   "Matrícula"
         Top             =   315
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   847
         Caption         =   "Matrícula"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorRotulo       =   4210752
         CorTexto        =   4194304
         MaxLen          =   15
      End
      Begin VTOcx.txtVISUAL txtResp 
         Height          =   480
         Left            =   90
         TabIndex        =   17
         Tag             =   "Responsável"
         Top             =   315
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   847
         Caption         =   "Responsável"
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
      TabIndex        =   14
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
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
Attribute VB_Name = "REES104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GeraCod As New ContaCorrente

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdDados.ListItems.Clear
    grdDados.Enabled = True
    TabDados.Tabs(2).Enabled = False
    
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

Private Sub grdDados_ItemCheck(ByVal Item As MSComctlLib.IListItem)
   TabDados.Tabs(2).Enabled = True
   TabDados.Tabs(2).Selected = True
   txtDesp.SetFocus
   grdDados.Enabled = False
   txtDesp = grdDados.SelectedItem.SubItems(4)
   txtResp = grdDados.SelectedItem.SubItems(5)
   txtMatricula = grdDados.SelectedItem.SubItems(6)
   txtData = grdDados.SelectedItem.SubItems(7)
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
    carregaProcesso
End Sub

Private Sub cmdSalvar_Click()
     ' status = 1 - aberto TipoProcesso Regime Especial = 3
    Dim camposPr As String
    Dim ValoresPr As String
    Dim Condicao As String
   
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Condicao = " TPR_NUMERO_PROCESSO = '" & grdDados.SelectedItem & "'"
    camposPr = " TPR_FUNCIONARIO_DESPACHO,TPR_FUNCIONARIO_MATRICULA,TPR_FUNCIONARIO_NOME,TPR_FUNCIONARIO_DATA_VISTO"
    ValoresPr = Bdados.PreparaValor(txtDesp, txtMatricula, txtResp, txtData)
    If Bdados.AtualizaDados("TAB_PROCESSO", ValoresPr, camposPr, Condicao) Then
        Avisa "Dados Salvos com Sucesso"
        carregaProcesso
        txtDesp = ""
    End If
  
End Sub



Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
     txtData = Date
End Sub


Private Sub carregaProcesso()
    Dim sql As String
    
    
    sql = "select TPR_NUMERO_PROCESSO as Processo, "
    sql = sql & " TPR_INSCRICAO as Inscrição, "
    sql = sql & " TPR_DESCRICAO_PEDIDO as Descrição, "
    sql = sql & " TGE_NOME   As Status, "
    sql = sql & " TPR_FUNCIONARIO_DESPACHO ,"
    sql = sql & " TPR_FUNCIONARIO_NOME,"
    sql = sql & " TPR_FUNCIONARIO_MATRICULA,"
    sql = sql & " TPR_FUNCIONARIO_DATA_VISTO"
    sql = sql & " From tab_processo, vis_status_Processo "
    sql = sql & " Where TPR_TIPO_PROCESSO = 3 And TPR_STATUS = 1 and TPR_STATUS = TGE_CODIGO and TPR_INSCRICAO = '" & txtIm & "'"
    
    If Not grdDados.Preencher(Bdados, sql, 1200, 1200, 5500, 1500, 0, 0, 0, 0) Then
        Avisa "Busca sem resultados"
    End If
    grdDados.Enabled = True
    TabDados.Tabs(2).Enabled = False

End Sub



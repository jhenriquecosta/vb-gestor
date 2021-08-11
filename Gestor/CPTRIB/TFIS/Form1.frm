VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   19
      Top             =   5700
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   9075
         TabIndex        =   16
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   10050
         TabIndex        =   17
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   345
         Left            =   8100
         TabIndex        =   15
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1138
      Icone           =   "Form1.frx":0000
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   4980
      Left            =   15
      TabIndex        =   20
      Tag             =   "Documento gerencial"
      Top             =   675
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   8784
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "Form1.frx":282A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4590
         Left            =   -99969
         TabIndex        =   21
         Top             =   30
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   8096
         _Version        =   131082
         TabGuid         =   "Form1.frx":28B5
         Begin VTOcx.cmdVISUAL CmdAdicionar 
            Height          =   345
            Left            =   8415
            TabIndex        =   13
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            Caption         =   "Adicionar"
            Acao            =   1
         End
         Begin VTOcx.grdVISUAL GrdDados 
            Height          =   3075
            Left            =   105
            TabIndex        =   32
            Top             =   1470
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   5424
         End
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   930
            Left            =   135
            TabIndex        =   33
            Top             =   90
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   1640
            Altura          =   1905
            Caption         =   " Apuração"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtValorRetido 
               Height          =   495
               Left            =   9045
               TabIndex        =   12
               Tag             =   "Valor Retido"
               Top             =   345
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   873
               Caption         =   "Valor Retido"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtValorImposto 
               Height          =   495
               Left            =   7695
               TabIndex        =   11
               Tag             =   "Valor Imposto"
               Top             =   345
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   873
               Caption         =   "Valor Imposto"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtAliquota 
               Height          =   495
               Left            =   6750
               TabIndex        =   10
               Tag             =   "Aliquota"
               Top             =   345
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   873
               Caption         =   "Aliquota"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtValorDocumento 
               Height          =   495
               Left            =   5130
               TabIndex        =   9
               Tag             =   "Valor Documento"
               Top             =   345
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   873
               Caption         =   "Valor Documento"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDataNota 
               Height          =   495
               Left            =   3825
               TabIndex        =   8
               Tag             =   "Data"
               Top             =   345
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   873
               Caption         =   "Data"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDocumento 
               Height          =   495
               Left            =   2505
               TabIndex        =   7
               Tag             =   "Nº Documento"
               Top             =   345
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   873
               Caption         =   "Nº Documento"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cboVISUAL CboTipoDocumento 
               Height          =   510
               Left            =   300
               TabIndex        =   6
               Tag             =   "Tipo de Documento"
               Top             =   345
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   900
               Caption         =   "Tipo de Documento"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.cmdVISUAL CmdExcluir 
            Height          =   345
            Left            =   9600
            TabIndex        =   14
            Top             =   1065
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            Caption         =   "Excluir"
            Acao            =   2
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4590
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   8096
         _Version        =   131082
         TabGuid         =   "Form1.frx":28DD
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1440
            Left            =   135
            TabIndex        =   23
            Top             =   1470
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   2540
            Altura          =   1905
            Caption         =   " Tomador"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   285
               Left            =   3510
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   390
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtInscricaoTomador 
               Height          =   285
               Left            =   240
               TabIndex        =   3
               Tag             =   "Inscricao Tomador"
               Top             =   390
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   503
               Caption         =   "Inscricao Tomador"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtEnderecoTomador 
               Height          =   285
               Left            =   1050
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1020
               Width           =   9465
               _ExtentX        =   16695
               _ExtentY        =   503
               Caption         =   "Endereço"
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   -2147483644
            End
            Begin VTOcx.txtVISUAL txtRazaoTomador 
               Height          =   285
               Left            =   1320
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   705
               Width           =   9195
               _ExtentX        =   16219
               _ExtentY        =   503
               Caption         =   "Razão"
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   -2147483644
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   1440
            Left            =   135
            TabIndex        =   27
            Top             =   3030
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   2540
            Altura          =   1905
            Caption         =   " Prestador"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtCNPJ 
               Height          =   300
               Left            =   8010
               TabIndex        =   5
               Top             =   375
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   529
               Caption         =   "CNPJ"
               Text            =   ""
               Formato         =   2
            End
            Begin VTOcx.txtVISUAL txtRazaoPrestador 
               Height          =   285
               Left            =   1335
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   705
               Width           =   9180
               _ExtentX        =   16193
               _ExtentY        =   503
               Caption         =   "Razão"
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   -2147483644
            End
            Begin VTOcx.txtVISUAL txtEnderecoPrestador 
               Height          =   285
               Left            =   1065
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   1020
               Width           =   9450
               _ExtentX        =   16669
               _ExtentY        =   503
               Caption         =   "Endereço"
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   -2147483644
            End
            Begin VTOcx.txtVISUAL txtInscricaoPrestador 
               Height          =   285
               Left            =   195
               TabIndex        =   4
               Tag             =   "Inscricao Prestador"
               Top             =   390
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   503
               Caption         =   "Inscricao Prestador"
               Text            =   ""
               Restricao       =   2
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL1 
               Height          =   285
               Left            =   3510
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   390
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   1215
            Left            =   135
            TabIndex        =   31
            Top             =   150
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   2143
            Altura          =   1905
            Caption         =   " Dados da Operação"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtDataLevantamento 
               Height          =   285
               Left            =   7455
               TabIndex        =   2
               Tag             =   "Data Levantamento"
               Top             =   795
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   503
               Caption         =   "Data Levantamento"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboFiscalizacao 
               Height          =   315
               Left            =   645
               TabIndex        =   0
               Tag             =   "Nº Fiscalização"
               Top             =   420
               Width           =   9915
               _ExtentX        =   17489
               _ExtentY        =   556
               Caption         =   "Nº Fiscalização"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL CboNatureza 
               Height          =   315
               Left            =   1155
               TabIndex        =   1
               Tag             =   "Natureza"
               Top             =   780
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   556
               Caption         =   "Natureza"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdicionar_Click()
    Dim Index As Integer
    
    If CriticaCampos(Me) = False Then Exit Sub
        
    '"Tipo Cocumento"  =  0
    '"Nº Documento = 1
    '"Data" = 2
    '"Valor Documento" = 3
    '"Aliquota" = 4
    '"Valor Impost" =5
    '"Valor Retido" = 6
    
        
    Index = GrdDados.ListItems.Count + 1
    
    GrdDados.ListItems.Add Index, , CboTipoDocumento.Coluna(0).VALOR & " - " & CboTipoDocumento.Text
    GrdDados.ListItems(Index).SubItems(1) = txtDocumento
    GrdDados.ListItems(Index).SubItems(2) = txtDataNota
    GrdDados.ListItems(Index).SubItems(3) = txtValorDocumento
    GrdDados.ListItems(Index).SubItems(4) = txtAliquota
    GrdDados.ListItems(Index).SubItems(5) = txtValorImposto
    GrdDados.ListItems(Index).SubItems(6) = txtValorRetido
    
    txtDocumento = ""
    txtDataNota = ""
    txtValorDocumento = "0,00"
    txtAliquota = "0,00"
    txtValorImposto = "0,00"
    txtValorRetido = "0,00"
    txtDocumento.SetFocus


End Sub

Private Sub CmdExcluir_Click()
    If GrdDados.ListItems.Count >= 1 Then
        GrdDados.ListItems.Remove GrdDados.SelectedItem.Index
    End If
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores  As String
    Dim Campos   As String
    Dim Condicao As String
    Dim contador As Integer
    
    'INDEX NA TABELA...
    'Tipo Cocumento   =  0
    'Nº Documento     =  1
    'Data             =  2
    'Valor Documento  =  3
    'Aliquota         =  4
    'Valor Impost     =  5
    'Valor Retido     =  6
    
    If Not CriticaCampos(Me) Then Exit Sub
    
    For contador = 1 To GrdDados.ListItems.Count
          Campos = "TAI_TFI_COD_FISCALIZACAO ,"
          Campos = Campos & "TAI_TCI_IM ,"
          Campos = Campos & "TAI_ISSQN_NATUREZA ,"
          Campos = Campos & "TAI_TIPO_DOCUMENTO ,"
          Campos = Campos & "TAI_NUM_DOCUMENTO ,"
          Campos = Campos & "TAI_ALIQUOTA ,"
          Campos = Campos & "TAI_VALOR_DOCUMENTO ,"
          Campos = Campos & "TAI_VALOR_IMPOSTO ,"
          Campos = Campos & "TAI_VALOR_RETIDO  ,"
          Campos = Campos & "TAI_PERIODO ,"
          Campos = Campos & "TAI_DATA_DOCUMENTO ,"
          Campos = Campos & "TAI_DATA_LEVANTAMENTO ,"
          Campos = Campos & "TAI_TCI_IM_ENVOLVIDO ,"
          Campos = Campos & "TAI_CNPJ_ENVOLVIDO "
          
'          Valores = Bdados.PreparaValor ( cboFiscalizacao.Coluna(0).Valor,Bdados.Converte(txtInscricaoTomador,tctexto)
    Next
    
    
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    CboNatureza.PreencherGeral Bdados, "NATUREZA APURACAO IMPOSTO"
    CboTipoDocumento.PreencherGeral Bdados, "TIPO DOC APURACAO IMPOSTO"
    With GrdDados.ColumnHeaders
        .Add , , "Tipo Cocumento"
        .Add , , "Nº Documento"
        .Add , , "Data"
        .Add , , "Valor Documento"
        .Add , , "Aliquota"
        .Add , , "Valor Imposto"
        .Add , , "Valor Retido"
    End With
End Sub

Private Sub GrdDados_Click()
    If GrdDados.ListItems.Count >= 1 Then
        txtDocumento = GrdDados.SelectedItem.SubItems(1)
        txtDataNota = GrdDados.SelectedItem.SubItems(2)
        txtValorDocumento = GrdDados.SelectedItem.SubItems(3)
        txtAliquota = GrdDados.SelectedItem.SubItems(4)
        txtValorImposto = GrdDados.SelectedItem.SubItems(5)
        txtValorRetido = GrdDados.SelectedItem.SubItems(6)
        GrdDados.ListItems.Remove GrdDados.SelectedItem.Index
    End If
End Sub

Private Sub TabStrip1_Click()

End Sub

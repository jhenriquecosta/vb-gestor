VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form THOF104 
   Caption         =   "THOF104"
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
      TabIndex        =   16
      Top             =   5700
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   9075
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   12
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
      TabIndex        =   15
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1138
      Icone           =   "THOF104.frx":0000
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   4980
      Left            =   15
      TabIndex        =   17
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
      Tabs            =   "THOF104.frx":282A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4590
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   8096
         _Version        =   131082
         TabGuid         =   "THOF104.frx":28B5
         Begin VTOcx.grdVISUAL GrdDados 
            Height          =   2835
            Left            =   120
            TabIndex        =   25
            Top             =   1755
            Width           =   10770
            _ExtentX        =   18997
            _ExtentY        =   5001
         End
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   1665
            Left            =   135
            TabIndex        =   26
            Top             =   75
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   2937
            Altura          =   1905
            Caption         =   " Apuração"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL CmdExcluir 
               Height          =   345
               Left            =   9435
               TabIndex        =   32
               Top             =   1275
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   609
               Caption         =   "Excluir"
               Acao            =   2
            End
            Begin VTOcx.cmdVISUAL CmdAdicionar 
               Height          =   345
               Left            =   8250
               TabIndex        =   31
               Top             =   1275
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   609
               Caption         =   "Adicionar"
               Acao            =   1
            End
            Begin VTOcx.fraVISUAL fraVISUAL2 
               Height          =   690
               Left            =   165
               TabIndex        =   28
               Top             =   900
               Width           =   4515
               _ExtentX        =   7964
               _ExtentY        =   1217
               Altura          =   1905
               Caption         =   " Prestador/Tomador"
               CorTexto        =   0
               CorFaixa        =   32768
               CorFundo        =   -2147483644
               Ocultavel       =   0   'False
               Begin VTOcx.txtVISUAL txtCGCEnvolvido 
                  Height          =   285
                  Left            =   1845
                  TabIndex        =   30
                  Top             =   345
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   503
                  Caption         =   "CPF/CNPJ"
                  Text            =   ""
                  Formato         =   2
                  CorFundo        =   -2147483644
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtImEnvolvido 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   29
                  Top             =   345
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   503
                  Caption         =   "Inscrição"
                  Text            =   ""
                  Restricao       =   2
                  CorFundo        =   -2147483644
                  MaxLen          =   20
                  RetirarMascara  =   0   'False
               End
            End
            Begin VTOcx.txtVISUAL txtValorDevido 
               Height          =   495
               Left            =   6810
               TabIndex        =   27
               Tag             =   "Valor Retido"
               Top             =   345
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   873
               Caption         =   "Valor Devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorFundo        =   -2147483644
               MaxLen          =   20
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtValorRetido 
               Height          =   495
               Left            =   9435
               TabIndex        =   11
               Top             =   345
               Width           =   1215
               _ExtentX        =   2143
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
            Begin VTOcx.txtVISUAL txtValorRecolhido 
               Height          =   495
               Left            =   8040
               TabIndex        =   10
               Top             =   345
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   873
               Caption         =   "Valor Recolhido"
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
               Left            =   6060
               TabIndex        =   9
               Tag             =   "Aliquota"
               Top             =   345
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   873
               Caption         =   "Aliq(%)"
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
               Left            =   4860
               TabIndex        =   8
               Tag             =   "Valor Documento"
               Top             =   345
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   873
               Caption         =   "Valor Doc."
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
               Left            =   3675
               TabIndex        =   7
               Tag             =   "Data"
               Top             =   345
               Width           =   1200
               _ExtentX        =   2117
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
               Left            =   2355
               TabIndex        =   6
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
               Left            =   150
               TabIndex        =   5
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
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4590
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   8096
         _Version        =   131082
         TabGuid         =   "THOF104.frx":28DD
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1800
            Left            =   135
            TabIndex        =   20
            Top             =   2220
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   3175
            Altura          =   1905
            Caption         =   " Contribuinte"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   285
               Left            =   3510
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   585
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   503
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtInscricaoTomador 
               Height          =   285
               Left            =   240
               TabIndex        =   4
               Top             =   585
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
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1215
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
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   900
               Width           =   9195
               _ExtentX        =   16219
               _ExtentY        =   503
               Caption         =   "Razão"
               Text            =   ""
               Enabled         =   0   'False
               CorFundo        =   -2147483644
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   1545
            Left            =   135
            TabIndex        =   24
            Top             =   405
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   2725
            Altura          =   1905
            Caption         =   " Dados da Operação"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtOF 
               Height          =   285
               Left            =   645
               TabIndex        =   0
               Top             =   630
               Width           =   2520
               _ExtentX        =   4445
               _ExtentY        =   503
               Caption         =   "Nº Fiscalização"
               Text            =   ""
               CorFundo        =   -2147483644
            End
            Begin VTOcx.txtVISUAL txtDataLevantamento 
               Height          =   285
               Left            =   7455
               TabIndex        =   3
               Top             =   960
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
               Left            =   3180
               TabIndex        =   1
               Top             =   585
               Visible         =   0   'False
               Width           =   7365
               _ExtentX        =   12991
               _ExtentY        =   556
               Caption         =   "Nº Fiscalização"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   -2147483644
            End
            Begin VTOcx.cboVISUAL CboNatureza 
               Height          =   315
               Left            =   1155
               TabIndex        =   2
               Top             =   945
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   556
               Caption         =   "Natureza"
               Text            =   ""
               AutoFocaliza    =   0   'False
               CorFundo        =   -2147483644
            End
         End
      End
   End
End
Attribute VB_Name = "THOF104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Calcula_Aliquota_BC_Recolhida() As Double
    Dim Base           As Double
    Dim ValorRecolhido As Double
    Dim X              As Double
    
    Base = txtValorDocumento
    ValorRecolhido = txtValorRecolhido
    
    X = (100 * ValorRecolhido) / Base
    Calcula_Aliquota_BC_Recolhida = X
    
End Function
Public Function Calcula_Aliquota_BC_Retida()
    Dim Base           As Double
    Dim ValorRetido As Double
    Dim X              As Double
    
    Base = txtValorDocumento
    ValorRetido = txtValorRetido
    
    X = (100 * ValorRetido) / Base
    Calcula_Aliquota_BC_Retida = X
    
End Function


Private Sub CmdAdicionar_Click()
    Dim Index As Integer
    txtDataLevantamento.Tag = ""
    If CriticaCampos(Me) = False Then Exit Sub
        
    '"Tipo Cocumento"  =  0
    '"Nº Documento = 1
    '"Data" = 2
    '"Valor Documento" = 3
    '"Aliquota" = 4
    '"Valor Impost" =5
    '"Valor Retido" = 6
    
    If txtImEnvolvido = "" And txtCGCEnvolvido = "" Then
        Avisa "Informe dados do prestador."
         If txtCGCEnvolvido = "" Then
            txtCGCEnvolvido.SetFocus
        Else
            txtImEnvolvido.SetFocus
        End If
        Exit Sub
    End If
        
    Index = GrdDados.ListItems.Count + 1
    
    GrdDados.ListItems.Add Index, , CboTipoDocumento.Coluna(1).Valor & " - " & CboTipoDocumento.Text
    GrdDados.ListItems(Index).SubItems(1) = txtDocumento
    GrdDados.ListItems(Index).SubItems(2) = txtDataNota
    GrdDados.ListItems(Index).SubItems(3) = txtValorDocumento
    GrdDados.ListItems(Index).SubItems(4) = txtAliquota
    GrdDados.ListItems(Index).SubItems(5) = txtValorDevido
    GrdDados.ListItems(Index).SubItems(6) = txtValorRecolhido
    GrdDados.ListItems(Index).SubItems(7) = txtValorRetido
    GrdDados.ListItems(Index).SubItems(8) = txtImEnvolvido
    GrdDados.ListItems(Index).SubItems(9) = "" & txtCGCEnvolvido
    GrdDados.ListItems(Index).SubItems(10) = "" & txtValorRecolhido
    GrdDados.ListItems(Index).SubItems(11) = Calcula_Aliquota_BC_Recolhida()
    GrdDados.ListItems(Index).SubItems(12) = Format(Calcula_Imposto_Recolhido, Const_Monetario)
    GrdDados.ListItems(Index).SubItems(13) = txtValorRetido
    GrdDados.ListItems(Index).SubItems(14) = Calcula_Aliquota_BC_Retida()
    GrdDados.ListItems(Index).SubItems(15) = Format(Calcula_Imposto_Retido, Const_Monetario)
    
    txtDocumento = ""
    txtDataNota = ""
    txtValorDocumento = "0,00"
    txtValorRecolhido = "0,00"
    txtValorRetido = "0,00"
    txtAliquota = "0,00"
    txtValorImposto = "0,00"
    txtValorRetido = "0,00"
    txtImEnvolvido = ""
    txtCGCEnvolvido = ""
    txtDocumento.SetFocus


End Sub
Private Function Calcula_Imposto_Retido()
    Calcula_Imposto_Retido = (Calcula_Aliquota_BC_Retida * txtValorDocumento) / 100
End Function
Private Sub CmdExcluir_Click()
    If GrdDados.ListItems.Count >= 1 Then
        GrdDados.ListItems.Remove GrdDados.SelectedItem.Index
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    GrdDados.ListItems.Clear
    
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscricaoTomador
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()

    Dim Valores            As String
    Dim Campos             As String
    Dim Condicao           As String
    Dim contador           As Integer
    Dim PosDocumento       As Integer
    Dim IRegistrosAfetados As Integer
    
    
    
    'INDEX NA TABELA...
    'Tipo Cocumento   =  0
    'Nº Documento     =  1
    'Data             =  2
    'Valor Documento  =  3
    'Aliquota         =  4
    'Valor Impost     =  5
    'Valor Retido     =  6
    
    If txtOF = "" Then
        Avisa "Informe o nº da ordem de fiscalização."
        
        Exit Sub
    End If
    
    If CboNatureza.ListIndex = -1 Then
        Avisa "Selecione Natureza."
        CboNatureza.SetFocus
        Exit Sub
    End If
    
    If txtDataLevantamento = "" Then
        Avisa "Informe " & txtDataLevantamento.Caption
        txtDataLevantamento.SetFocus
        Exit Sub
    End If
    If txtInscricaoTomador = "" Then
        Avisa "Informe " & txtInscricaoTomador.Caption
        txtInscricaoTomador.SetFocus
        Exit Sub
    End If
    

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
          
          PosDocumento = InStr(GrdDados.ListItems(contador), " - ")
          Valores = Bdados.PreparaValor(txtOF, _
          Bdados.Converte(txtInscricaoTomador, tctexto), _
          CboNatureza.Coluna(1).Valor, _
          Left(GrdDados.ListItems(contador), _
          PosDocumento - 1), _
          GrdDados.ListItems(contador).SubItems(1), _
          Bdados.Converte(GrdDados.ListItems(contador).SubItems(4), TCMonetario), _
          Bdados.Converte(GrdDados.ListItems(contador).SubItems(3), TCMonetario), _
          Bdados.Converte(GrdDados.ListItems(contador).SubItems(5), TCMonetario), _
          Bdados.Converte(GrdDados.ListItems(contador).SubItems(6), TCMonetario), _
          Month(GrdDados.ListItems(contador).SubItems(2)) & Year(GrdDados.ListItems(contador).SubItems(2)), _
          GrdDados.ListItems(contador).SubItems(2), _
          txtDataLevantamento, _
          Bdados.Converte(txtInscricaoPrestador, tctexto), _
          Bdados.Converte(txtCNPJ, tctexto))
          
          Condicao = "TAI_TFI_COD_FISCALIZACAO = " & txtOF
          Condicao = Condicao & " and TAI_TIPO_DOCUMENTO = " & Left(GrdDados.ListItems(contador), PosDocumento - 1)
          Condicao = Condicao & " and TAI_NUM_DOCUMENTO = " & GrdDados.ListItems(contador).SubItems(1)
          If Bdados.GravaDados("TAB_APURACAO_IMPOSTO", Valores, Campos, Condicao) Then
                IRegistrosAfetados = IRegistrosAfetados + 1
          End If
    Next
        
        If IRegistrosAfetados = GrdDados.ListItems.Count Then
            Avisa "Operação concluída com sucesso."
            cmdLimpar_Click
        Else
            Erro "Erro ao gravar Apuração."
        End If
    
    
End Sub


Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    CboNatureza.PreencherGeral Bdados, "NATUREZA APURACAO IMPOSTO"
    CboTipoDocumento.PreencherGeral Bdados, "TIPO DOC APURACAO IMPOSTO"
    With GrdDados.ColumnHeaders
        .Add , , "Tipo Cocumento", 1500
        .Add , , "Nº Documento"
        .Add , , "Data"
        .Add , , "Valor Documento"
        .Add , , "Aliquota"
        .Add , , "Valor Devido"
        .Add , , "Valor Recolhido"
        .Add , , "Valor Retido"
        .Add , , "Inscrição"
        .Add , , "CNPJ"
        '------------------------------
        '------------------------------
        .Add , , "BC Recolhida"
        .Add , , "Aliq."
        .Add , , "Imposto Recolhido"
        '-----------------------------
        .Add , , "BC Retida"
        .Add , , "Aliq."
        .Add , , "Imposto Retido"
        
    End With
End Sub
Private Sub CalculaImposto()
    On Error Resume Next
    txtValorDevido = (Nvl(txtAliquota, 0) * Nvl(txtValorDocumento, 0)) / 100
End Sub
Private Sub TabStrip1_Click()

End Sub

Private Sub GrdDados_DblClick()
    If GrdDados.ListItems.Count >= 1 Then
    
        txtDocumento = GrdDados.SelectedItem.SubItems(1)
        txtDataNota = GrdDados.SelectedItem.SubItems(2)
        txtValorDocumento = GrdDados.SelectedItem.SubItems(3)
        txtAliquota = GrdDados.SelectedItem.SubItems(4)
        txtValorImposto = GrdDados.SelectedItem.SubItems(5)
        txtValorRetido = GrdDados.SelectedItem.SubItems(6)
        txtValorRetido = GrdDados.SelectedItem.SubItems(7)
        txtImEnvolvido = GrdDados.SelectedItem.SubItems(8)
        txtCGCEnvolvido = GrdDados.SelectedItem.SubItems(9)
        txtValorRecolhido = GrdDados.SelectedItem.SubItems(10)
        txtValorRetido = GrdDados.SelectedItem.SubItems(13)
        GrdDados.ListItems.Remove GrdDados.SelectedItem.Index
        
    End If
End Sub
Private Function Calcula_Imposto_Recolhido() As Double

     Calcula_Imposto_Recolhido = (Calcula_Aliquota_BC_Recolhida * txtValorDocumento) / 100
     
End Function
Private Sub txtAliquota_Change()

    CalculaImposto
    
End Sub


Private Sub txtInscricaoTomador_LostFocus()

    If Trim(txtInscricaoTomador) = "" Then Exit Sub
    txtInscricaoTomador = BuscaContribuinte(txtInscricaoTomador, txtRazaoTomador, txtEnderecoTomador, , etiContribuinte)
    
End Sub

Private Sub txtValorDocumento_Change()
    CalculaImposto
End Sub


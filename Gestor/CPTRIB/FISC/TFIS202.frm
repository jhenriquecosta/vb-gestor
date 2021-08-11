VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TFIS202 
   Caption         =   "TPRT101"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1138
      Icone           =   "TFIS202.frx":0000
      Codigo          =   "txtObs"
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   5820
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   767
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   330
         Left            =   5820
         TabIndex        =   22
         Top             =   90
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   7020
         TabIndex        =   3
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8025
         TabIndex        =   5
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
      End
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   5070
      Left            =   0
      TabIndex        =   8
      Tag             =   "Documento gerencial"
      Top             =   690
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   8943
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "TFIS202.frx":031A
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4680
         Left            =   -99969
         TabIndex        =   9
         Top             =   30
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   8255
         _Version        =   131082
         TabGuid         =   "TFIS202.frx":03B5
         Begin VTOcx.fraVISUAL fraVISUAL4 
            Height          =   1395
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   60
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   2461
            Altura          =   1905
            Caption         =   " Fiscalização"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Enabled         =   0   'False
            Begin VTOcx.txtVISUAL txtNomeContrib 
               Height          =   285
               Left            =   3150
               TabIndex        =   16
               Top             =   690
               Width           =   5850
               _ExtentX        =   10319
               _ExtentY        =   503
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               CorRotulo       =   0
               CorTexto        =   4194304
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtIM 
               Height          =   285
               Left            =   300
               TabIndex        =   15
               Top             =   690
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   503
               Caption         =   "Contribuinte"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCodFiscalizacao 
               Height          =   285
               Left            =   90
               TabIndex        =   14
               Top             =   360
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   503
               Caption         =   "Nº Fiscalização"
               Text            =   ""
               Enabled         =   0   'False
               Restricao       =   2
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtEndereco 
               Height          =   285
               Left            =   540
               TabIndex        =   13
               Top             =   1050
               Width           =   8460
               _ExtentX        =   14923
               _ExtentY        =   503
               Caption         =   "Endereço"
               Text            =   ""
               Enabled         =   0   'False
               CorRotulo       =   0
               CorTexto        =   4194304
               RetirarMascara  =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL3 
            Height          =   885
            Left            =   0
            TabIndex        =   17
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   1440
            Width           =   9030
            _ExtentX        =   15928
            _ExtentY        =   1561
            Altura          =   1905
            Caption         =   " Levantamento Homologatório"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtFim 
               Height          =   495
               Left            =   1560
               TabIndex        =   19
               Top             =   300
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   873
               Caption         =   "Prazo Final"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtInicio 
               Height          =   495
               Left            =   60
               TabIndex        =   4
               Top             =   300
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   873
               Caption         =   "Data Inicial"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtProcedimentoAtual 
               Height          =   495
               Left            =   3630
               TabIndex        =   18
               Top             =   300
               Width           =   5325
               _ExtentX        =   9393
               _ExtentY        =   873
               Caption         =   "Procedimento Atual"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
            End
         End
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   885
            Left            =   0
            TabIndex        =   20
            ToolTipText     =   "Pesquisa Contribuintes"
            Top             =   2370
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   1561
            Altura          =   1905
            Caption         =   " Dados do Procedimento"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtData 
               Height          =   495
               Left            =   60
               TabIndex        =   0
               Top             =   300
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   873
               Caption         =   "Data"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.cboVISUAL cboAutoridade 
               Height          =   510
               Left            =   4890
               TabIndex        =   2
               Top             =   300
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   900
               Caption         =   "Autoridade Fiscal Responsável"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.txtVISUAL txtDtMaxEncerra 
               Height          =   495
               Left            =   3000
               TabIndex        =   21
               Top             =   300
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   873
               Caption         =   "Prazo Encerramento"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtDtEncerra 
               Height          =   495
               Left            =   1230
               TabIndex        =   1
               Top             =   300
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   873
               Caption         =   "Data Encerramento"
               Text            =   ""
               Formato         =   0
               Restricao       =   2
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               AgruparValores  =   0   'False
               RetirarMascara  =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4680
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   8255
         _Version        =   131082
         TabGuid         =   "TFIS202.frx":03DD
         Begin VTOcx.fraVISUAL fraVISUAL2 
            Height          =   6675
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   11460
            _ExtentX        =   20214
            _ExtentY        =   11774
            Altura          =   1905
            Caption         =   " txtDocumentos"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
            Begin VB.TextBox txtObs 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   1095
               Left            =   120
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Tag             =   "Descrição"
               Top             =   3240
               Width           =   8820
            End
            Begin VB.TextBox txtInformacoes 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   1095
               Left            =   90
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Tag             =   "Descrição"
               Top             =   1890
               Width           =   8820
            End
            Begin VB.TextBox txtDocumentos 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   1095
               Left            =   90
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Tag             =   "Descrição"
               Top             =   540
               Width           =   8820
            End
            Begin VB.Label Label1 
               Caption         =   "Relatos/Fundamentos/Informações Fiscais"
               Height          =   165
               Index           =   2
               Left            =   90
               TabIndex        =   28
               Top             =   3000
               Width           =   3255
            End
            Begin VB.Label Label1 
               Caption         =   "Informações Requeridas"
               Height          =   165
               Index           =   1
               Left            =   90
               TabIndex        =   27
               Top             =   1680
               Width           =   2235
            End
            Begin VB.Label Label1 
               Caption         =   "Documentos Solicitados"
               Height          =   165
               Index           =   0
               Left            =   90
               TabIndex        =   26
               Top             =   330
               Width           =   2235
            End
         End
      End
   End
End
Attribute VB_Name = "TFIS202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio
Dim Fisc As New Fiscalizacao
Dim Fase As Integer
Dim Procedimento As Integer
Dim CodAndamento As String
Private Sub Cancelar()
    Dim Motivo As String
    
    If Trim(txtDtEncerra) <> "" Then
        Avisa "Fiscalização já encerrada em " & Trim(txtCodFiscalizacao) & " não poderá ser cancelada. "
        Exit Sub
    End If
    If Confirma("Confirma o cancelamento da Fiscalização nº " & txtCodFiscalizacao & "?") Then
        Motivo = Trim(Util.Entrada("Informe o Motivo do cancelamento", "Informação obrigatória"))
        If Trim(Motivo) = "" Then
            Erro "Cancelamento de Fiscalização não foi concluído por falta de motivo."
            Exit Sub
        End If
        If Bdados.AtualizaDados("TAB_FISCALIZACAO", Bdados.PreparaValor(2, Bdados.Converte(Date, TCDataHora), AplicacoesVTFuncoes.Usuario), _
            "TFI_STATUS,TFI_DATA_CANCELAMENTO,TIF_USUARIO_CANCELAMENTO", _
            "TFI_COD_FISCALIZACAO = " & txtCodFiscalizacao) Then
            Avisa "Fiscaliazação cancelada com sucesso."
            Unload Me
        End If
    End If
End Sub

Private Sub Encerrar()
    Dim Data As String
    If Trim(txtCodFiscalizacao) <> "" Then
        Avisa "Fiscalização já encerrada em " & Trim(txtCodFiscalizacao) & "."
        Exit Sub
    End If
    If Confirma("Confirma o encerramento da Fiscalização nº " & txtCodFiscalizacao & "?") Then
        Data = Trim(Util.Entrada("Informe a Data do encerramento (dd/mm/aaaa)", "Informação obrigatória(dd/mm/aaaa)"))
        If Not IsDate(Trim(Data)) Then
            Avisa "Data inválida."
            Exit Sub
        End If
        If Trim(Data) = "" Then
            Erro "encerramento de Fiscalização não foi concluído por falta da data."
            Exit Sub
        End If
        If DateDiff("d", txtCodFiscalizacao, Data) < 0 Then
            Erro "Data de encerramento da ação fiscal não pode ser menor do que a data de início da ação fiscal."
            Exit Sub
        End If
        If Bdados.AtualizaDados("TAB_FISCALIZACAO", Bdados.PreparaValor(3, Bdados.Converte(Data, TCDataHora), AplicacoesVTFuncoes.Usuario), "TFI_STATUS,TFI_DATA_TEAF,TIF_USUARIO_TEAF", "TFI_COD_FISCALIZACAO = " & txtCodFiscalizacao) Then
            Avisa "Fiscaliazação encerrada com sucesso."
            If Confirma("Deseja imprimir o TEAF agora?") Then
                
            End If
            Unload Me
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
End Sub

Private Sub cmdImprimir_Click()
    If Confirma("Deseja imprimir o documento agora?") Then
        With Rpt
            If Not .DefinirArquivo(Bdados, Fisc.Rede.rCaminhoRpt) Then Exit Sub
            .Formulas "VT_PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            .Formulas "VT_SECRETARIA", UCase(Temp.PegaParametro(Bdados, "SECRETARIA"))
            .Formulas "VT_COD_FISCALIZACAO", txtCodFiscalizacao
            .Formulas "VT_COD_SEQUENCIA_PROCEDIMENTO", Trim(CodAndamento)
            .Formulas "DOCUMENTO", UCase(Util.ParseString(Me.Caption, "|", 2))
            .Formulas "VT_NUM_DOC", " Nº " & Left(Trim(CodAndamento), 2) & "." & _
                Mid(Trim(CodAndamento), 3, 3) & "." & Mid(Trim(CodAndamento), 6, 3)
            .Arvore = False
            .Visualizar
            Set Rpt = Nothing
        End With
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Conta As New ContaCorrente
    
    CodAndamento = Trim(Util.ParseString(Me.Caption, "|", 3))
    If Fisc.Andamento.GravaAndamentoProcesso(txtCodFiscalizacao, Trim(Util.ParseString(Me.Caption, "|", 1)), txtData, txtDtEncerra, _
         CInt(Nvl(CStr(cboAutoridade.Coluna(0).Valor), 0)), CodAndamento, txtDocumentos, txtInformacoes, txtObs) Then
        Avisa "Procedimento gravado com sucesso."
        If Trim(Fisc.Rede.rCaminhoRpt) <> "" Then
            cmdImprimir_Click
        End If
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Fisc.Funcionario.PreencheComboFuncionario cboAutoridade
    
End Sub

Private Sub Form_Resize()
    If Trim(Me.Tag) <> "" Then
        Set Fisc = Nothing
        Set Fisc = New Fiscalizacao
        txtCodFiscalizacao = Me.Tag
        txtCodFiscalizacao_LostFocus
        Me.Tag = ""
        If Trim(Fisc.Rede.rCaminhoRpt) <> "" And Fisc.Andamento.vCodSequenciaProcedimento <> "" Then
            cmdImprimir.Visible = True
        End If
    End If
End Sub

Private Sub txtData_LostFocus()
    If Trim(txtData) = "" Then txtData = Date
    If Not IsDate(txtData) Then Exit Sub
    If DateDiff("d", txtData, Date) < 0 Then
        Avisa "Data inválida"
        If txtData.Enabled = True Then txtData.SetFocus
    End If
    If Fisc.Rede.rPrazo > 0 Then txtDtMaxEncerra = DateAdd("d", Fisc.Rede.rPrazo, txtData)
End Sub

Private Sub txtDtEncerra_LostFocus()
    If Trim(txtDtEncerra) = "" Then Exit Sub
    If DateDiff("d", txtDtEncerra, Date) < 0 Then
        Avisa "Data inválida"
        txtDtEncerra.SetFocus
    End If
    
    If DateDiff("d", txtData, txtDtEncerra) < 0 Then
        Avisa "Data inválida"
        txtDtEncerra.SetFocus
    End If
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIM) = 10 Or Len(txtIM) = 11 Then
            Ic = Imposto.FormataInscricao(txtIM, InscContrib)
        Else
            Ic = txtIM
        End If
    Else
            Ic = txtIM
    End If
    txtIM = BuscaContribuinte(Ic, txtNomeContrib, txtEndereco)
End Sub

Private Sub txtCodFiscalizacao_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Titulo As String
    
    Set Fisc = Nothing
    Set Fisc = New Fiscalizacao
    Titulo = Me.Caption
    Fisc.Andamento.vCodSequenciaProcedimento = Util.ParseString(Me.Caption, "|", 3)
    If Fisc.CarregaDadosFiscalizacao(txtCodFiscalizacao) Then
        txtIM = Fisc.vIm
        txtIm_LostFocus
        txtInicio = Fisc.vDataInicio
        txtFim = Fisc.vDataFim
        txtData = Fisc.Andamento.vDataAbertura
        If Trim(txtData) = "" Then
            Fisc.Andamento.vEtapa.CarregaDadosRede Util.ParseString(Me.Caption, "|", 1)
            If Fisc.Andamento.vEtapa.rOrdem > 0 Then
                CodAndamento = Fisc.Andamento.BuscaCodigoAndamento(txtCodFiscalizacao, Fisc.Andamento.vEtapa.rCodEtapa)
                If Trim(CodAndamento) <> "" Then Fisc.Andamento.CarregaAndamentoFiscalizacao txtCodFiscalizacao, CodAndamento
                txtData = Fisc.Andamento.vDataAbertura
                If Trim(CodAndamento) <> "" Then Me.Caption = Me.Caption & " | " & CodAndamento
            End If
        End If
        If Trim(CodAndamento) = "" Then CodAndamento = Util.ParseString(Me.Caption, "|", 3)
        txtDtEncerra = Fisc.Andamento.vDataConclusao
        txtObs = Fisc.Andamento.vRelato
        txtDocumentos = Fisc.Andamento.vDocumentos
        txtInformacoes = Fisc.Andamento.vInformacoes
        Fisc.Rede.CarregaDadosRede Trim(Util.ParseString(Titulo, "|", 1))
        If Trim(Fisc.Andamento.vCodSequenciaProcedimento) = "" Then
            Fisc.Rede.ParametrosTexto.CarregaDadosParametro Fisc.Rede.rCodParametroFundamento
            txtObs = Fisc.Rede.ParametrosTexto.vDescricao

        End If
        txtFim = Fisc.vDataFim
        cboAutoridade.SetarLinha Fisc.vCodFuncionario, 0
        
        txtProcedimentoAtual = Fisc.Rede.rDescricao
        If Trim(txtData) <> "" Then txtData_LostFocus
        cmdSalvar.Enabled = IIf(Trim(Fisc.Andamento.vDataConclusao) <> "", False, True)
        txtDtEncerra.Enabled = IIf(Trim(Fisc.Andamento.vDataConclusao) <> "", False, True)
        cboAutoridade.Enabled = IIf(Trim(Fisc.Andamento.vDataConclusao) <> "", False, True)
        txtObs.Enabled = IIf(Trim(Fisc.Andamento.vDataConclusao) <> "", False, True)
        txtDocumentos.Enabled = IIf(Trim(Fisc.Andamento.vDataConclusao) <> "", False, True)
        txtInformacoes.Enabled = IIf(Trim(Fisc.Andamento.vDataConclusao) <> "", False, True)
        txtData.Enabled = IIf(Trim(txtData) <> "", False, True)
        If Not txtData.Enabled And txtDtEncerra.Enabled Then txtDtEncerra.SetFocus
    End If
End Sub


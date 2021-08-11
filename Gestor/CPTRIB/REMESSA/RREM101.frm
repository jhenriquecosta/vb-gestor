VERSION 5.00
Object = "{741D44DD-BF8E-4BC8-85FF-338C9BF39DFB}#1.0#0"; "Cabecalho.ocx"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{467EEF11-5281-4102-AFD3-AD54F754C329}#1.5#0"; "VTControles.ocx"
Begin VB.Form RREM101 
   BackColor       =   &H00FBEDE8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RREM101"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FBEDE8&
      Caption         =   "Marcar Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   18
      Top             =   5475
      Width           =   1905
   End
   Begin VTOcx.cboVISUAL CboBradesco_Identificacao_Ocorrencia 
      Height          =   510
      Left            =   2580
      TabIndex        =   17
      Top             =   5400
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   900
      Caption         =   "Identificação Ocorrência"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Alinhamento     =   1
      TipoCampo       =   ""
      PictureFundo    =   "RREM101.frx":0000
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   1138
      Icone           =   "RREM101.frx":001C
      ImagemFundo     =   "RREM101.frx":0336
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   14
      Top             =   6165
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   900
      CorFundo        =   -2147483633
      ImagemFundo     =   "RREM101.frx":14290
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   390
         Left            =   6015
         TabIndex        =   9
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   688
         Caption         =   "Gerar Arquivo Bradesco"
         Acao            =   3
      End
      Begin VTOcx.cmdVISUAL cmdSalvar_Especial 
         Height          =   390
         Left            =   6615
         TabIndex        =   10
         Top             =   525
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         Caption         =   "Gerar Arquivo BB"
         Acao            =   3
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   390
         Left            =   8445
         TabIndex        =   11
         Top             =   60
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   688
         Caption         =   "Limpar"
         Acao            =   6
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9450
         TabIndex        =   12
         Top             =   60
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
      End
      Begin VTOcx.cmdVISUAL cmdBuscar_Especial 
         Height          =   390
         Left            =   5010
         TabIndex        =   8
         Top             =   60
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   688
         Caption         =   "Buscar"
         Acao            =   5
      End
   End
   Begin VTOcx.grdVISUAL Grid 
      Height          =   3300
      Left            =   90
      TabIndex        =   15
      Top             =   2430
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   5821
      Caption         =   "Dados"
      CheckBox        =   -1  'True
      PictureBarra    =   "RREM101.frx":3048E
   End
   Begin VTOcx.fraVISUAL txt 
      Height          =   1725
      Left            =   90
      TabIndex        =   16
      Top             =   675
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3043
      Altura          =   1905
      Caption         =   " Consultar Por:"
      CorTexto        =   0
      CorFaixa        =   12632256
      CorFundo        =   16510440
      Ocultavel       =   0   'False
      BackStyle       =   0
      Picture         =   "RREM101.frx":304AA
      Picture2        =   "RREM101.frx":35894
      Begin VTOcx.cboVISUAL cboServicoConsulta 
         Height          =   315
         Left            =   510
         TabIndex        =   2
         Top             =   645
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   556
         Caption         =   "Serviço"
         Text            =   ""
         TipoCampo       =   ""
         PictureFundo    =   "RREM101.frx":358B0
      End
      Begin VTOcx.cboVISUAL cboAlunoConsulta 
         Height          =   315
         Left            =   2430
         TabIndex        =   1
         Top             =   315
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   556
         Caption         =   "Aluno"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   0
         TipoCampo       =   ""
         PictureFundo    =   "RREM101.frx":358CC
      End
      Begin VTOcx.txtVISUAL txtNotaConsulta 
         Height          =   285
         Left            =   4845
         TabIndex        =   5
         Top             =   1020
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   503
         Caption         =   "Matricula"
         Text            =   ""
         PictureFundo    =   "RREM101.frx":358E8
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   7005
         TabIndex        =   6
         Top             =   990
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         CorRotulo       =   0
         Enabled         =   0   'False
         TipoCampo       =   ""
         PictureFundo    =   "RREM101.frx":35904
      End
      Begin VTOcx.txtVISUAL txtNumero 
         Height          =   285
         Left            =   480
         TabIndex        =   0
         Top             =   315
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   503
         Caption         =   "Número"
         Text            =   ""
         PictureFundo    =   "RREM101.frx":35920
      End
      Begin VTOcx.cboVISUAL cboRepresentante_Financeiro 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Tag             =   "Representante Financeiro"
         Top             =   1350
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   556
         Caption         =   "Representante Financeiro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         TipoLetras      =   0
         TipoCampo       =   ""
         PictureFundo    =   "RREM101.frx":3593C
      End
      Begin VTOcx.txtVISUAL txtInicio 
         Height          =   285
         Left            =   165
         TabIndex        =   3
         Top             =   1005
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         Caption         =   "Vencimento"
         Text            =   ""
         Formato         =   0
         PictureFundo    =   "RREM101.frx":35958
      End
      Begin VTOcx.txtVISUAL txtFim 
         Height          =   285
         Left            =   2715
         TabIndex        =   4
         Top             =   1005
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         Caption         =   "Até"
         Text            =   ""
         Formato         =   0
         PictureFundo    =   "RREM101.frx":35974
      End
   End
   Begin ActiveResizer.SSResizer SSResizer2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196610
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   10275
      DesignHeight    =   6675
   End
End
Attribute VB_Name = "RREM101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboTurma_LostFocus()

    'If cboTurma.Text = "" Then Exit Sub
    
End Sub

Private Sub Check1_Click()
    Grid.MarcarTodos Check1.Value
End Sub

Public Sub cmdBuscar_Especial_Click()
    Dim Sql As String

    Sql = "Select * from VIS_CONTA_receber where 1 = 1 "
    If cboAlunoConsulta.Text <> "" Then
        Sql = Sql & " AND Aluno = '" & cboAlunoConsulta.Text & "'"
    End If
    If txtNumero <> "" Then
        Sql = Sql & " and código = " & txtNumero
    End If
    If cboServicoConsulta <> "" Then
        Sql = Sql & " AND [Curso / Serviço] = '" & cboServicoConsulta.Text & "'"
    End If
    If txtNotaConsulta <> "" Then
        Sql = Sql & " and Matricula = " & txtNotaConsulta
    End If
    If cboStatus <> "" Then
        Sql = Sql & " and Status = '" & cboStatus.Text & "'"
    End If
    If cboRepresentante_Financeiro.Text <> "" Then
        Sql = Sql & " and [Representante Financeiro]  like '" & Trim(Left(cboRepresentante_Financeiro.Text, 30)) & "%'"
    End If
    If txtInicio <> "" And txtFim <> "" Then
        Sql = Sql & " and Vencimento >= " & Bdados.Converte(txtInicio, TCDataHora)
        Sql = Sql & " and Vencimento <= " & Bdados.Converte(txtFim, TCDataHora)
    ElseIf txtInicio <> "" And txtFim = "" Then
        Sql = Sql & " and Vencimento >= " & Bdados.Converte(txtInicio, TCDataHora)
        Sql = Sql & " and Vencimento <= " & Bdados.Converte(txtInicio, TCDataHora)
    End If
    'Sql = Sql & " order by TCR_descricao"
    Sql = Sql & " and left([status arquivo bradesco],1) = 2 "
    'Sql = Sql & " and  [status arquivo bradesco] is null"
    If Grid.Preencher(Bdados, Sql) Then
        'Grid.Mensagem = "Tot Original : " & Format(Grid.Colunas(9).Soma, "currency") & "| Tot Desconto : " & Format(Grid.Colunas(10).Soma, "currency") & "| Tot a Pagar : " & Format(Grid.Colunas(11).Soma, "currency") & "| Total Pago : " & Format(Grid.Colunas(12).Soma, "currency") & "| Saldo Devedor : " & Format(Grid.Colunas(13).Soma, "currency")
    Else
        Avisa "Consulta sem resultados."
        'Grid.Mensagem = "Nenhum registro encontrado."
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    Grid.ListItems.Clear
    cboStatus.SetarLinha esrAberto, 1
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Public Sub cmdSalvar_Especial_Click()
    Dim Arquivos_Remessa        As New Remessa
    Dim NomeArquivo                As String
    Dim i                                  As Integer
    Dim Sql As String
    Dim rs As VSRecordset
    
    If Grid.ListItems.Count >= 1 Then
    If Me.Tag <> "" Then
        GoTo Vai
    End If
        If Confirma("Confirma geração do arquivos remessa?") Then
Vai:
            With Arquivos_Remessa
                .PrefixoAgencia = PegaConfiguracaoEscola(PREFIXO_AGENCIA)
                .SequencialRemessa Bdados
                .DigitoVerificadorPrefixoAgencia = PegaConfiguracaoEscola(DV_PREFIXO_AGENCIA)
                .CodigoCedente = PegaConfiguracaoEscola(CODIGO_CEDENTE)
                .DigitoVerificadorCodigoCedente = PegaConfiguracaoEscola(DV_CODIGO_CEDENTE)
                .NumeroConvenente = PegaConfiguracaoEscola(NUMERO_CONVENIO)
                .NomeEmpresa = RetornaDadosEscola(PegaConfiguracaoEscola(Escola), TES_NOME)
                .DataGravacao = Date
                NomeArquivo = .Gera_HEADER
                For i = 1 To Grid.ListItems.Count
'                    .TipoInscricaoCedente = "02"
'                    .InscricaoCedente = "41373291000140"
'                    .PrefixoAgencia = "1312"
'                    .DigitoVerificadorPrefixoAgencia = 9
'                    .CodigoCedente = "24595"
'                    .DigitoVerificadorCodigoCedente = "X"
'                    .NumeroConvenio = "177655"
'                    .NumeroControleParticipante = "0000000000000000000059872" 'Ano + Matricula + Aluno
'                    .NossoNumero = "17765559872"
'                    .DigitoVerificadorNossoNumero = "8"
'                    .NumeroPrestacao = "00"
'                    .IndicativoSacador = " "
'                    .PrefixoTitulo = "AI"
'                    .VariacaoCaretira = " 01"
'                    .ContaCaucao = "9"
'                    .CodigoResponsabilidade = "00000"
'                    .DVCodigoResponsabilidade = "0"
'                    .NumeroBordero = "0"
'                    .Carteira = "     "
'                    .NumeroCarteira = "17"
'                    .Comando = "01"
'                    .SeuNumero = "0000059872"
'                    .DataVencimento = "050705"
'                    .ValorTitulo = "120"
'                    .NumeroBanco = "001"
'                    .PrefixoAgenciaCobradora = "0000"
'                    .DV_Pref_Agencia_Cobradora = "0"
'                    .EspecieTitulo = 12
'                    .Aceite = "N"
'                    .DataEmissao = "010705"
'                    .PrimeiraInstrucaoCodificada = "10"
'                    .SegundaInstrucaoCodificada = "00"
'                    .JurosMoraDia = "0,12"
'                    .DataLimiteConcessaoDesconto = "050705"
'                    .ValorDesconto = "12"
'                    .CampoEspecial_1 = "0000000000000"
'                    .ValorAbatimento = "0000000000000"
'                    .TipoInscricaoSacado = "01"
'                    .CPF_CNPJ_Sacado = "00028027841372"
'                    .NomeSacado = "WILSON FERNANDES LIMA"
'                    .EnderecoSacado = "RUA SAO FRANCISCO, S/N"
'                    .CepSacado = "65690000"
'                    .CidadeSacado = "COLINAS"
'                    .UFSacado = "MA"
'                    .Observacoes = "APOS VENCIMENTO: MULTA DE 2%"
'                    .DiasParaProtesto = "   "
'                    .SequencialRegistro = I + 1

                    'Dados do Aluno
                   Sql = "Select * from tab_alunos where tan_codigo = " & Grid.ListItems(i).SubItems(2)
                   If Bdados.AbreTabela(Sql) Then
                    .NumeroControleParticipante = Trim(Bdados.Tabela("tan_registro"))
                    .Observacoes = PegaConfiguracaoEscola(Observacoes)
                        If Bdados.AbreTabela("Select * from tab_representante where TRP_CODIGO = " & Bdados.Tabela("TAN_TRP_RESPONSAVEL_FINANCEIRO"), rs) Then
                            .TipoInscricaoSacado = "01"
                            .CPF_CNPJ_Sacado = rs.Fields("trp_doc")
                            .NomeSacado = Trim(rs.Fields("trp_nome"))
                            .EnderecoSacado = "" & rs.Fields("trp_endereco") & "nº" & rs.Fields("trp_numero")
                            .CepSacado = "" & rs.Fields("trp_cep")
                            .CidadeSacado = rs.Fields("trp_cidade")
                            .UFSacado = "MA"
                        End If
                    End If
                    'Dados Configuração
                    .Aceite = PegaConfiguracaoEscola(Aceite)
                    .PrimeiraInstrucaoCodificada = PegaConfiguracaoEscola(PRIMEIRA_INSTRUCAO_CODIFICADA)
                    .SegundaInstrucaoCodificada = PegaConfiguracaoEscola(SEGUNDA_INSTRUCAO_CODIFICADA)
                    .NumeroCarteira = PegaConfiguracaoEscola(NUMERO_CARTEIRA)
                    .Comando = PegaConfiguracaoEscola(Comando)
                    .EspecieTitulo = PegaConfiguracaoEscola(Especie_Titulo)
                    .PrefixoTitulo = PegaConfiguracaoEscola(NUMERO_CARTEIRA)
                    .PrefixoAgencia = PegaConfiguracaoEscola(PREFIXO_AGENCIA)
                    .DigitoVerificadorPrefixoAgencia = PegaConfiguracaoEscola(DV_PREFIXO_AGENCIA)
                    .CodigoCedente = PegaConfiguracaoEscola(CODIGO_CEDENTE)
                    .DigitoVerificadorCodigoCedente = PegaConfiguracaoEscola(DV_CODIGO_CEDENTE)
                    .NumeroConvenio = PegaConfiguracaoEscola(NUMERO_CONVENIO)
                    .IndicativoSacador = PegaConfiguracaoEscola(INDICATIVO_SACADOR)
                    .VariacaoCaretira = PegaConfiguracaoEscola(VARIACAO_CARTEIRA)
                    .ContaCaucao = PegaConfiguracaoEscola(CONTA_CAUCAO)
                    .CodigoResponsabilidade = PegaConfiguracaoEscola(CODIGO_RESPONSABILIDADE)
                    .DVCodigoResponsabilidade = PegaConfiguracaoEscola(DV_CODIGO_RESPONSABILIDADE)
                    .NumeroBordero = PegaConfiguracaoEscola(NUMERO_BORDEIRO)
                    .Carteira = PegaConfiguracaoEscola(CARTEIRAS)
                    .DiasParaProtesto = PegaConfiguracaoEscola(DIAS_PARA_PROTESTO)
                    .NumeroBanco = PegaConfiguracaoEscola(NUMERO_BANCO)
                    .PrefixoAgenciaCobradora = PegaConfiguracaoEscola(PREFIXO_AGENCIA_COBRADORA)
                    .DV_Pref_Agencia_Cobradora = PegaConfiguracaoEscola(DV_PREFIXO_AGENCIA_COBRADORA)
                    .DataLimiteConcessaoDesconto = Grid.ListItems(i).SubItems(7) 'PegaConfiguracaoEscola(DIAS_LIMITE_CONCESSAO_DESCONTO)
                    .TipoInscricaoCedente = "02"
                    .InscricaoCedente = RetornaDadosEscola(PegaConfiguracaoEscola(Escola), TES_DOC)
                    .CampoEspecial_1 = "0000000000000"
                    
                    .JurosMoraDia = (PegaConfiguracaoEscola(Juros) * Grid.ListItems(i).SubItems(8)) / 30
                    
                    'Dados do Débito
                    
                    .NossoNumero = Right(Grid.ListItems(i), 5)
                    .NumeroPrestacao = "00"
                    .SeuNumero = Grid.ListItems(i)
                    .DataVencimento = Grid.ListItems(i).SubItems(7)
                    .ValorTitulo = Format(Grid.ListItems(i).SubItems(8), Const_Monetario)
                    .DataEmissao = Date
                    
                    .ValorDesconto = Format(Grid.ListItems(i).SubItems(9), Const_Monetario)
                    .ValorAbatimento = "0,00"
                    
                    .SequencialRegistro = i + 1
                    .Gera_DETALHE NomeArquivo
                    Call MudaStatusArquivoDebito(Grid.ListItems(i), esadGerado)
                Next
                    .Gera_TRAILER NomeArquivo
            End With
        End If
    End If
    Avisa "Arquivos gerados com sucesso."
    If Me.Tag = "" Then
        cmdBuscar_Especial_Click
    Else
        Unload Me
    End If
End Sub

Private Sub cmdVISUAL1_Click()
    Dim Arquivos_Remessa        As New RemessaBradesco
    Dim NomeArquivo             As String
    Dim i                       As Integer
    Dim Sql                     As String
    Dim rs                      As VSRecordset
    Dim Marcou As Boolean
    
    Marcou = False
    
    For i = 1 To Grid.ListItems.Count
        If Grid.ListItems(i).Checked Then
            Marcou = True
            Exit For
        End If
    Next
    
    If Marcou = False Then
        Avisa "Selecione um recebimento para geração do arquivo remessa"
        Exit Sub
    End If
    
    
    If CboBradesco_Identificacao_Ocorrencia = "" Then
        Avisa "Selecione a ocorrência para o banco"
        CboBradesco_Identificacao_Ocorrencia.SetFocus
        Exit Sub
    End If
    If Grid.ListItems.Count >= 1 Then
    If Me.Tag <> "" Then
        GoTo Vai
    End If
        If Confirma("Confirma geração do arquivos remessa?") Then
Vai:
            With Arquivos_Remessa
                .Codigo_Empresa = PegaConfiguracaoEscola(TEC_BRADESCO_CODIGO_EMPRESA)
                .Identificador_Sistema = PegaConfiguracaoEscola(TEC_BRADESCO_IDENTIFICADOR_SISTEMA)
                .Numero_Bradesco = PegaConfiguracaoEscola(TEC_BRADESCO_NUMERO_BRADESCO)
                NomeArquivo = .Gera_HEADER
                If NomeArquivo = "" Then
                    Exit Sub
                End If
                
                For i = 1 To Grid.ListItems.Count
                    If Grid.ListItems(i).Checked Then
                         .Carteira = PegaConfiguracaoEscola(TEC_BRADESCO_CARTEIRA)
                         .Agencia = PegaConfiguracaoEscola(TEC_BRADESCO_AGENCIA)
                         .Conta_Corrente = PegaConfiguracaoEscola(TEC_BRADESCO_CONTA_CORRENTE)
                         .DV_Conta_Corrente = PegaConfiguracaoEscola(TEC_BRADESCO_DV_CONTA_CORRENTE)
                         'Dados do Aluno
                        Sql = "Select * from tab_alunos where tan_codigo = " & Grid.ListItems(i).SubItems(2)
                        If Bdados.AbreTabela(Sql) Then
                         .Numero_Controle_Participante = Trim(Bdados.Tabela("tan_registro"))
                             If Bdados.AbreTabela("Select * from tab_representante where TRP_CODIGO = " & Bdados.Tabela("TAN_TRP_RESPONSAVEL_FINANCEIRO"), rs) Then
                                 .Numero_Inscricao_Sacado = rs.Fields("trp_doc")
                                 .Nome_Sacado = Trim(rs.Fields("trp_nome"))
                                 .Endereco_Sacado = "" & rs.Fields("trp_endereco") & "nº" & rs.Fields("trp_numero")
                                 .Cep_Sacado = "" & rs.Fields("trp_cep")
                             End If
                         End If
                         .Identificador_Do_Titulo_Banco_Nosso_Numero = Grid.ListItems(i)
                         .Valor_Desconto_Bonificacao_Dia = PegaConfiguracaoEscola(TEC_BRADESCO_VALOR_DESCONTO_BONIFICACAO_DIA)
                         .Condicao_Emissao_Papeleta_Cobranca = PegaConfiguracaoEscola(TEC_BRADESCO_CONDICAO_EMISSAO_PAPELETA_COBRANCA)
                         .Identificacao_Ocorrencia = CboBradesco_Identificacao_Ocorrencia.Coluna(1).Valor
                         .Numero_Documento = Grid.ListItems(i).SubItems(2)
                         .Data_Vencimento_Titulo = Grid.ListItems(i).SubItems(7)
                         .Valor_Titulo = Format(Grid.ListItems(i).SubItems(8), Const_Monetario)
                         .Especie_Titulo = PegaConfiguracaoEscola(TEC_BRADESCO_ESPECIE_TITULO)
                         .Identificacao_Aceite = Left(RetornaDadosVisGeral(VIS_ACEITE, PegaConfiguracaoEscola(TEC_BRADESCO_IDENTIFICACAO_ACEITE), edvNome), 1)
                         .Data_Emissao_Titulo = Date
                         .Instrucao_1 = PegaConfiguracaoEscola(TEC_BRADESCO_1_INSTRUCAO)
                         .Instrucao_2 = PegaConfiguracaoEscola(TEC_BRADESCO_2_INSTRUCAO)
                         .Valor_Cobrado_Dia_Atraso = PegaConfiguracaoEscola(TEC_BRADESCO_VALOR_COBRADO_DIA_ATRASO)
                         .Data_LImite_Concessao_Desconto = Grid.ListItems(i).SubItems(7) 'VENCIMENTO
                         .Valor_Desconto = Format(Grid.ListItems(i).SubItems(9), Const_Monetario)
                         .Valor_IOF = String(13, "0")
                         .Mensagem_1 = PegaConfiguracaoEscola(TEC_BRADESCO_1_MENSAGEM)
                         .Mensagem_2 = PegaConfiguracaoEscola(TEC_BRADESCO_2_MENSAGEM)
                         
                         .Sequencia_Registro_Detalhe = i + 1
                         .Gera_DETALHE NomeArquivo
                         Call MudaStatusArquivoDebitoBradesco(Grid.ListItems(i), esadGerado)
                     End If
                Next
                    .Gera_TRAILER NomeArquivo
            End With
        End If
    End If
    Avisa "Arquivos gerados com sucesso."
    If Me.Tag = "" Then
        cmdBuscar_Especial_Click
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     cboStatus.PreencherGeral Bdados, "STATUS COTA"
    cboRepresentante_Financeiro.Preencher Bdados, "Select trp_codigo,trp_nome  + '  /  ' + trp_doc from tab_representante", 1
    cboAlunoConsulta.Preencher Bdados, "SELECT TAN_CODIGO,TAN_NOME FROM TAB_ALUNOS", 1
    cboServicoConsulta.Preencher Bdados, "SELECT tcu_codigo,tcu_nome FROM TAB_CURSOS", 1
    cboStatus.SetarLinha esrAberto, 1
    CboBradesco_Identificacao_Ocorrencia.PreencherGeral Bdados, "BRADESCO IDENTIFICAO OCORRENCIA"
End Sub


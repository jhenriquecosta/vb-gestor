VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemessaBradesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Codigo_Empresa                                        As String
Private m_Numero_Bradesco                                       As String
Private m_Identificador_Sistema                                 As String
Private CodBarra                                                As CodigoDeBarra

Private m_CARTEIRA                                              As String
Private m_AGENCIA                                               As String
Private m_CONTA_CORRENTE                                        As String
Private m_DV_CONTA_CORRENTE                                     As String
Private m_VALOR_DESCONTO_BONIFICACAO_DIA                        As String
Private m_CONDICAO_EMISSAO_PAPELETA_COBRANCA                    As String
Private m_Especie_Titulo                                        As String
Private m_IDENTIFICACAO_ACEITE                                  As String
Private m_1_INSTRUCAO                                           As String
Private m_2_INSTRUCAO                                           As String
Private m_VALOR_COBRADO_DIA_ATRASO                              As String
Private m_1_MENSAGEM                                            As String
Private m_2_MENSAGEM                                            As String

Private m_Identificacao_Registro                               As String
Private m_Agencia_Debito_Sacado                                As String
Private m_Digito_Agencia_Debito_Sacado                         As String
Private m_RAzao_Conta_Corrente_Sacado                          As String
Private m_Conta_Corrente_Sacado                                As String
Private m_Digito_Conta_Corrente_Sacado                         As String
Private m_Identificacao_Empresa_Banco                          As String
Private m_Numero_Controle_Participante                         As String
Private m_Codigo_Banco_A_Ser_Na_Camera_Compensacao             As String
Private m_Emitir_Papeleta_Debito_Automatico                    As String
Private m_Identificacao_Operacao_Banco                         As String
Private m_Indicador_Rateio_Credito                             As String
Private m_Enderecamento_Aviso_Debito_Automatico_Conta_Corrente As String
Private m_Branco_107_108                                       As String
Private m_Identificacao_Ocorrencia                             As String
Private m_Numero_Documento                                     As String
Private m_Data_Vencimento_Titulo                               As String
Private m_Valor_Titulo                                         As String
Private m_Banco_Encarregado_Cobranca                           As String
Private m_Agencia_Depositaria                                  As String
Private m_Data_Emissao_Titulo                                  As String
Private m_Valor_Desconto                                       As String
Private m_Valor_IOF                                            As String
Private m_Valor_Abatimento_Concedito_Cancelado                 As String
Private m_Tipo_Inscricao_Sacado                                As String
Private m_Numero_Inscricao_Sacado                              As String
Private m_Nome_Sacado                                          As String
Private m_Endereco_Sacado                                      As String
Private m_Cep_Sacado                                           As String
Private m_Sequencia_Registro_Detalhe                           As String
Private m_Identificador_Do_Titulo_Banco_Nosso_Numero           As String
Private m_Data_LImite_Concessao_Desconto                       As String

Public Property Get Sequencia_Registro_Detalhe() As String
   Sequencia_Registro_Detalhe = m_Sequencia_Registro_Detalhe
End Property

Public Property Let Sequencia_Registro_Detalhe(Valor As String)
    m_Sequencia_Registro_Detalhe = Format(Valor, String(6, "0"))
End Property


Public Property Get Cep_Sacado() As String
   Cep_Sacado = m_Cep_Sacado
End Property

Public Property Let Cep_Sacado(Valor As String)
    If Valor <> "" Then
        m_Cep_Sacado = Left(Valor, 5) & Right(Valor, 3)
    Else
        m_Cep_Sacado = String(8, "0")
    End If
End Property

Public Property Get Endereco_Sacado() As String
   Endereco_Sacado = m_Endereco_Sacado
End Property

Public Property Let Endereco_Sacado(Valor As String)
    Dim Nome As String
    Nome = Trim(Left(Valor, 40))
    If Len(Nome) < 40 Then
        m_Endereco_Sacado = Nome & String(40 - Len(Nome), " ")
    Else
        m_Endereco_Sacado = Nome
    End If
End Property


Public Property Get Nome_Sacado() As String
   Nome_Sacado = m_Nome_Sacado
End Property

Public Property Let Nome_Sacado(Valor As String)
    Dim Nome As String
    Nome = Trim(Left(Valor, 40))
    If Len(Nome) < 40 Then
        m_Nome_Sacado = Nome & String(40 - Len(Nome), " ")
    Else
        m_Nome_Sacado = Nome
    End If
End Property


Public Property Get Numero_Inscricao_Sacado() As String
   Numero_Inscricao_Sacado = m_Numero_Inscricao_Sacado
End Property

Public Property Let Numero_Inscricao_Sacado(Valor As String)
    m_Numero_Inscricao_Sacado = String(14 - Len(TiraTudo(Valor)), "0") & TiraTudo(Valor)
End Property



Public Property Get Valor_IOF() As String
   Valor_IOF = m_Valor_IOF
End Property

Public Property Let Valor_IOF(Valor As String)
    m_Valor_IOF = Valor
End Property


Public Property Get Valor_Desconto() As String
   Valor_Desconto = m_Valor_Desconto
End Property

Public Property Let Valor_Desconto(Valor As String)
    m_Valor_Desconto = Format(Nvl(TiraTudo(Valor), 0), String(13, "0"))
End Property



Public Property Get Data_LImite_Concessao_Desconto() As String
   Data_LImite_Concessao_Desconto = m_Data_LImite_Concessao_Desconto
End Property

Public Property Let Data_LImite_Concessao_Desconto(Valor As String)
    m_Data_LImite_Concessao_Desconto = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Year(Valor), 2)
End Property


Public Property Get Data_Emissao_Titulo() As String
   Data_Emissao_Titulo = m_Data_Emissao_Titulo
End Property

Public Property Let Data_Emissao_Titulo(Valor As String)
    m_Data_Emissao_Titulo = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Year(Valor), 2)
End Property

Public Property Get Identificador_Do_Titulo_Banco_Nosso_Numero() As String
   Identificador_Do_Titulo_Banco_Nosso_Numero = m_Identificador_Do_Titulo_Banco_Nosso_Numero & CodBarra.Calculo_DV11(m_CARTEIRA & m_Identificador_Do_Titulo_Banco_Nosso_Numero)
End Property

Public Property Let Identificador_Do_Titulo_Banco_Nosso_Numero(Valor As String)
    m_Identificador_Do_Titulo_Banco_Nosso_Numero = Format(Valor, String(11, "0"))
End Property


Public Property Get Tipo_Inscricao_Sacado() As String
   Tipo_Inscricao_Sacado = m_Tipo_Inscricao_Sacado
End Property

Public Property Let Tipo_Inscricao_Sacado(Valor As String)
    m_Tipo_Inscricao_Sacado = Valor
End Property


Public Property Get Valor_Abatimento_Concedito_Cancelado() As String
   Valor_Abatimento_Concedito_Cancelado = m_Valor_Abatimento_Concedito_Cancelado
End Property

Public Property Let Valor_Abatimento_Concedito_Cancelado(Valor As String)
    m_Valor_Abatimento_Concedito_Cancelado = Valor
End Property


Public Property Get Codigo_Banco_A_Ser_Na_Camera_Compensacao() As String
   Codigo_Banco_A_Ser_Na_Camera_Compensacao = m_Codigo_Banco_A_Ser_Na_Camera_Compensacao
End Property

Public Property Let Codigo_Banco_A_Ser_Na_Camera_Compensacao(Valor As String)
    m_Codigo_Banco_A_Ser_Na_Camera_Compensacao = Valor
End Property



Public Property Get Agencia_Depositaria() As String
   Agencia_Depositaria = m_Agencia_Depositaria
End Property

Public Property Let Agencia_Depositaria(Valor As String)
    m_Agencia_Depositaria = Valor
End Property

Public Property Get Banco_Encarregado_Cobranca() As String
   Banco_Encarregado_Cobranca = m_Banco_Encarregado_Cobranca
End Property

Public Property Let Banco_Encarregado_Cobranca(Valor As String)
    m_Banco_Encarregado_Cobranca = Valor
End Property

Public Property Get Valor_Titulo() As String
   Valor_Titulo = m_Valor_Titulo
End Property

Public Property Let Valor_Titulo(Valor As String)
    m_Valor_Titulo = Format(Nvl(TiraTudo(Valor), 0), String(13, "0"))
End Property

Public Property Get Data_Vencimento_Titulo() As String
   Data_Vencimento_Titulo = m_Data_Vencimento_Titulo
End Property

Public Property Let Data_Vencimento_Titulo(Valor As String)
    m_Data_Vencimento_Titulo = Format(Day(Valor), "00") & Format(Month(Valor), "00") & Right(Year(Valor), 2)
End Property

Public Property Get Numero_Documento() As String
   Numero_Documento = m_Numero_Documento
End Property

Public Property Let Numero_Documento(Valor As String)
    m_Numero_Documento = Format(Nvl(Valor, 0), String(10, "0"))
End Property

Public Property Get Identificacao_Ocorrencia() As String
   Identificacao_Ocorrencia = m_Identificacao_Ocorrencia
End Property

Public Property Let Identificacao_Ocorrencia(Valor As String)
    m_Identificacao_Ocorrencia = Format(Valor, String(2, "0"))
End Property

Public Property Get Branco_107_108() As String
   Branco_107_108 = m_Branco_107_108
End Property

Public Property Let Branco_107_108(Valor As String)
    m_Branco_107_108 = Valor
End Property

Public Property Get Enderecamento_Aviso_Debito_Automatico_Conta_Corrente() As String
   Enderecamento_Aviso_Debito_Automatico_Conta_Corrente = m_Enderecamento_Aviso_Debito_Automatico_Conta_Corrente
End Property

Public Property Let Enderecamento_Aviso_Debito_Automatico_Conta_Corrente(Valor As String)
    m_Enderecamento_Aviso_Debito_Automatico_Conta_Corrente = Valor
End Property

Public Property Get Indicador_Rateio_Credito() As String
   Indicador_Rateio_Credito = m_Indicador_Rateio_Credito
End Property

Public Property Let Indicador_Rateio_Credito(Valor As String)
    m_Indicador_Rateio_Credito = Valor
End Property

Public Property Get Identificacao_Operacao_Banco() As String
   Identificacao_Operacao_Banco = m_Identificacao_Operacao_Banco
End Property

Public Property Let Identificacao_Operacao_Banco(Valor As String)
    m_Identificacao_Operacao_Banco = Valor
End Property

Public Property Get Emitir_Papeleta_Debito_Automatico() As String
   Emitir_Papeleta_Debito_Automatico = m_Emitir_Papeleta_Debito_Automatico
End Property

Public Property Let Emitir_Papeleta_Debito_Automatico(Valor As String)
    m_Emitir_Papeleta_Debito_Automatico = Valor
End Property

Public Property Get Numero_Controle_Participante() As String
    Numero_Controle_Participante = m_Numero_Controle_Participante
End Property

Public Property Let Numero_Controle_Participante(Valor As String)
    m_Numero_Controle_Participante = Valor & String(25 - Len(Valor), " ")
End Property


Public Property Get Identificacao_Empresa_Banco() As String
    Identificacao_Empresa_Banco = "0" & Format(m_CARTEIRA, "000") & 0 & Left(m_AGENCIA, 4) & m_CONTA_CORRENTE & m_DV_CONTA_CORRENTE
End Property

Public Property Get Digito_Conta_Corrente_Sacado() As String
    Digito_Conta_Corrente_Sacado = m_Digito_Conta_Corrente_Sacado
End Property

Public Property Let Digito_Conta_Corrente_Sacado(Valor As String)
    m_Digito_Conta_Corrente_Sacado = Valor
End Property



Public Property Get Conta_Corrente_Sacado() As String
    Conta_Corrente_Sacado = m_Conta_Corrente_Sacado
End Property

Public Property Let Conta_Corrente_Sacado(Valor As String)
    m_Conta_Corrente_Sacado = Valor
End Property



Public Property Get RAzao_Conta_Corrente_Sacado() As String
    RAzao_Conta_Corrente_Sacado = m_RAzao_Conta_Corrente_Sacado
End Property

Public Property Let RAzao_Conta_Corrente_Sacado(Valor As String)
    m_RAzao_Conta_Corrente_Sacado = Valor
End Property

Public Property Get Digito_Agencia_Debito_Sacado() As String
    Digito_Agencia_Debito_Sacado = m_Digito_Agencia_Debito_Sacado
End Property

Public Property Let Digito_Agencia_Debito_Sacado(Valor As String)
    m_Digito_Agencia_Debito_Sacado = Valor
End Property

Public Property Get Agencia_Debito_Sacado() As String
    Agencia_Debito_Sacado = m_Agencia_Debito_Sacado
End Property

Public Property Let Agencia_Debito_Sacado(Valor As String)
    m_Agencia_Debito_Sacado = Valor
End Property

Public Property Get Identificacao_Registro() As String
    Identificacao_Registro = m_Identificacao_Registro
End Property

Public Property Let Identificacao_Registro(Valor As String)
    m_Identificacao_Registro = Valor
End Property

Public Property Get Mensagem_2() As String
    Mensagem_2 = m_2_MENSAGEM
End Property

Public Property Let Mensagem_2(Valor As String)
Dim Nome As String
    Nome = Trim(Left(Valor, 60))
    If Len(Nome) < 60 Then
        m_2_MENSAGEM = Nome & String(60 - Len(Nome), " ")
    Else
        m_2_MENSAGEM = Nome
    End If
End Property

Public Property Get Mensagem_1() As String
    Mensagem_1 = m_1_MENSAGEM
End Property

Public Property Let Mensagem_1(Valor As String)
 Dim Nome As String
    Nome = Trim(Left(Valor, 12))
    If Len(Nome) < 12 Then
        m_1_MENSAGEM = Nome & String(12 - Len(Nome), " ")
    Else
        m_1_MENSAGEM = Nome
    End If
End Property

Public Property Get Valor_Cobrado_Dia_Atraso() As String
    Valor_Cobrado_Dia_Atraso = m_VALOR_COBRADO_DIA_ATRASO
End Property

Public Property Let Valor_Cobrado_Dia_Atraso(Valor As String)
    m_VALOR_COBRADO_DIA_ATRASO = Format(Nvl(TiraTudo(Valor), 0), String(13, "0"))
End Property

Public Property Get Instrucao_2() As String
    Instrucao_2 = m_2_INSTRUCAO
End Property

Public Property Let Instrucao_2(Valor As String)
    If Valor = 99 Then
        m_2_INSTRUCAO = "00"
    Else
        m_2_INSTRUCAO = Format(Valor, "00")
    End If
End Property


Public Property Get Instrucao_1() As String
    Instrucao_1 = m_1_INSTRUCAO
End Property

Public Property Let Instrucao_1(Valor As String)
    m_1_INSTRUCAO = Format(Valor, "00")
End Property

Public Property Get Identificacao_Aceite() As String
    Identificacao_Aceite = m_IDENTIFICACAO_ACEITE
End Property

Public Property Let Identificacao_Aceite(Valor As String)
    m_IDENTIFICACAO_ACEITE = Valor
End Property

Public Property Get Especie_Titulo() As String
    Especie_Titulo = m_Especie_Titulo
End Property

Public Property Let Especie_Titulo(Valor As String)
    m_Especie_Titulo = Valor
End Property


Public Property Get Condicao_Emissao_Papeleta_Cobranca() As String
    Condicao_Emissao_Papeleta_Cobranca = m_CONDICAO_EMISSAO_PAPELETA_COBRANCA
End Property

Public Property Let Condicao_Emissao_Papeleta_Cobranca(Valor As String)
    m_CONDICAO_EMISSAO_PAPELETA_COBRANCA = Valor
End Property

Public Property Get Valor_Desconto_Bonificacao_Dia() As String
    Valor_Desconto_Bonificacao_Dia = m_VALOR_DESCONTO_BONIFICACAO_DIA
End Property

Public Property Let Valor_Desconto_Bonificacao_Dia(Valor As String)
    m_VALOR_DESCONTO_BONIFICACAO_DIA = Format(TiraTudo(Valor), String(10, "0"))
End Property


Public Property Get DV_Conta_Corrente() As String
    DV_Conta_Corrente = m_DV_CONTA_CORRENTE
End Property

Public Property Let DV_Conta_Corrente(Valor As String)
    m_DV_CONTA_CORRENTE = Valor
End Property

Public Property Get Conta_Corrente() As String
    Conta_Corrente = m_CONTA_CORRENTE
End Property

Public Property Let Conta_Corrente(Valor As String)
    m_CONTA_CORRENTE = Format(Valor, String(7, "0"))
End Property


Public Property Get Agencia() As String
    Agencia = m_AGENCIA
End Property

Public Property Let Agencia(Valor As String)
    m_AGENCIA = Format(Valor, String(5, "0"))
End Property


Public Property Get Carteira() As String
    Carteira = m_CARTEIRA
End Property

Public Property Let Carteira(Valor As String)
    m_CARTEIRA = Format(Valor, "00")
End Property




Public Property Get Identificador_Sistema() As String
    Identificador_Sistema = m_Identificador_Sistema
End Property

Public Property Let Identificador_Sistema(Valor As String)
    m_Identificador_Sistema = UCase(Valor)
End Property


Public Property Get Numero_Bradesco() As String
    Numero_Bradesco = m_Numero_Bradesco
End Property

Public Property Let Numero_Bradesco(Valor As String)
    m_Numero_Bradesco = Valor
End Property


Public Property Get Codigo_Empresa() As String
    Codigo_Empresa = m_Codigo_Empresa
End Property

Public Property Let Codigo_Empresa(Valor As String)
    m_Codigo_Empresa = Format(Valor, String(20, "0"))
End Property


Public Function Sequencial_Cabecalho() As String
    Sequencial_Cabecalho = Right(GeraCorrelativo(ecSequencialCabecalhoBradesco), 7)
End Function

Public Function Gera_TRAILER(Arquivo As String)
    Dim NumeroArquivo         As Integer
    Dim Dados                      As String
    NumeroArquivo = FreeFile

    'Abro o arquivo como leitura para pegar o último registro...
    Open PegaConfiguracaoEscola(TEC_BRADESCO_LOCAL_SALVAR_ARQUIVO) & "\" & Arquivo For Input As #NumeroArquivo
    Do Until EOF(Numero)
            Line Input #Numero, Dados
    Loop
    Close NumeroArquivo
    'Gero o Trailer
    Open PegaConfiguracaoEscola(TEC_BRADESCO_LOCAL_SALVAR_ARQUIVO) & "\" & Arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, 9 & String(393, " ") & Format(Right(Dados, 6) + 1, String(6, "0"))
    Close NumeroArquivo
End Function
Public Function Gera_DETALHE(Arquivo As String)
    Dim NumeroArquivo         As Integer
    NumeroArquivo = FreeFile
    Open PegaConfiguracaoEscola(TEC_BRADESCO_LOCAL_SALVAR_ARQUIVO) & "\" & Arquivo For Append Access Write As #NumeroArquivo
    Print #NumeroArquivo, Me.Identificacao_Registro & Me.Agencia_Debito_Sacado & Me.Digito_Agencia_Debito_Sacado & _
                          Me.RAzao_Conta_Corrente_Sacado & Me.Conta_Corrente_Sacado & Me.Digito_Conta_Corrente_Sacado & _
                          Me.Identificacao_Empresa_Banco & Me.Numero_Controle_Participante & _
                          Me.Codigo_Banco_A_Ser_Na_Camera_Compensacao & Me.Identificador_Do_Titulo_Banco_Nosso_Numero & _
                          Me.Valor_Desconto_Bonificacao_Dia & Me.Condicao_Emissao_Papeleta_Cobranca & _
                          Me.Emitir_Papeleta_Debito_Automatico & Me.Identificacao_Operacao_Banco & _
                          Me.Indicador_Rateio_Credito & Me.Enderecamento_Aviso_Debito_Automatico_Conta_Corrente & _
                          Me.Branco_107_108 & Me.Identificacao_Ocorrencia & _
                          Me.Numero_Documento & Me.Data_Vencimento_Titulo & _
                          Me.Valor_Titulo & Me.Banco_Encarregado_Cobranca & _
                          Me.Agencia_Depositaria & Me.Especie_Titulo & _
                          Me.Identificacao_Aceite & Me.Data_Emissao_Titulo & Me.Instrucao_1 & Me.Instrucao_2 & _
                          Me.Valor_Cobrado_Dia_Atraso & Me.Data_LImite_Concessao_Desconto & Me.Valor_Desconto & _
                          Me.Valor_IOF & Me.Valor_Abatimento_Concedito_Cancelado & Me.Tipo_Inscricao_Sacado & _
                          Me.Numero_Inscricao_Sacado & Me.Nome_Sacado & Me.Endereco_Sacado & Me.Mensagem_1 & Me.Cep_Sacado & Me.Mensagem_2 & Me.Sequencia_Registro_Detalhe
    Close NumeroArquivo
End Function
Public Function Gera_HEADER() As String
    On Error GoTo TRATA
    Dim Arquivo         As Integer
    Dim NomeArquio   As String
    Dim Sequencia As String
    Dim NomeEscola As String
    
    NomeEscola = Left(RetornaDadosEscola(PegaConfiguracaoEscola(Escola), TES_NOME), 30)
    NomeEscola = Trim(NomeEscola) & String(30 - Len(NomeEscola), " ")
    Sequencia = Sequencial_Cabecalho
    Arquivo = FreeFile
    NomeArquio = "CB" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Val(Right(Sequencia, 2)), "00") & ".REM"
    Open PegaConfiguracaoEscola(TEC_BRADESCO_LOCAL_SALVAR_ARQUIVO) & NomeArquio For Output Access Write As #Arquivo
    Print #Arquivo, "0" & "1" & "REMESSA" & "01" & "COBRANCA       " & m_Codigo_Empresa & NomeEscola & m_Numero_Bradesco & "BRADESCO       " & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Right(Year(Date), 2), "00") & String(8, " ") & m_Identificador_Sistema & Sequencia & String(277, " ") & "000001"
    Close Arquivo
    Gera_HEADER = NomeArquio
    
TRATA:
    If Err.Number = 76 Then
        Avisa "Erro ao gerar arquivo, verifique se o caminho definido na configuração existe:" & PegaConfiguracaoEscola(TEC_BRADESCO_LOCAL_SALVAR_ARQUIVO)
    End If
End Function
Private Sub Class_Initialize()
    Set CodBarra = New CodigoDeBarra
    Me.Identificacao_Registro = 1
    Me.Agencia_Debito_Sacado = String(5, "0")
    Me.Digito_Agencia_Debito_Sacado = String(1, "0")
    Me.RAzao_Conta_Corrente_Sacado = String(5, "0")
    Me.Conta_Corrente_Sacado = String(7, "0")
    Me.Digito_Conta_Corrente_Sacado = String(1, "0")
    Me.Codigo_Banco_A_Ser_Na_Camera_Compensacao = String(8, "0") '+ 5 zeros pos:66 a 70
    Me.Identificacao_Operacao_Banco = String(10, " ")
    Me.Indicador_Rateio_Credito = String(1, " ")
    Me.Enderecamento_Aviso_Debito_Automatico_Conta_Corrente = 0
    Me.Branco_107_108 = String(2, " ")
    Me.Numero_Documento = String(10, "0")
    Me.Banco_Encarregado_Cobranca = String(3, "0")
    Me.Agencia_Depositaria = String(5, "0")
    Me.Valor_Abatimento_Concedito_Cancelado = String(13, "0")
    Me.Tipo_Inscricao_Sacado = "01" ' Tipo = CPF
    Me.Emitir_Papeleta_Debito_Automatico = "0"
End Sub
Private Sub Class_Terminate()
    Set CodBarra = Nothing
End Sub


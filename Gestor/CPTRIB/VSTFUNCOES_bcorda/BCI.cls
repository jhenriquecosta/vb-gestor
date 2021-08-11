VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private ValorEdificacao As Double
Private ValorTerreno As Double
Private valorMercadoEdificacao As Double
Private valorMercadoTerreno As Double

'Informacoes do Contribuinte
Option Explicit
Private ContribuinteInscricao As String
Private ContribuinteInscricaoAuxiliar As String
Private ContribuinteNome As String
Private ContribuinteCPFcgc As String
Private ContribuinteEnderecoCodLogradouro As String
Private ContribuinteEnderecoTipoLogradouro As String
Private ContribuinteEnderecoLogradouro As String
Private ContribuinteEnderecoNumero As String
Private ContribuinteEnderecoCompl As String
Private ContribuinteEnderecoCodBairro As String
Private ContribuinteEnderecoBairro As String
Private ContribuinteCEP As String
Private ContribuinteCidade As String
Private ContribuinteUF As String

'Informacoes do Imovel
'-OBRIGATORIAS-
Private ImovelInscricao As String
Private ImovelInscricaoAnterior As String
Private ImovelUnidade As String
Private ImovelSubUnidade As String
Private ImovelInscricaoCondominio As String
Private ImovelInscricaoAnteriorBC As String
Private ImovelCodLogr As String
Private ImovelFracaoIdeal As String
Private ImovelCodTipoLogr As String
Private ImovelCodBairro As String
Private ImovelEnderecoNumero As String
Private ImovelEnderecoCompl As String
Private ImovelLote As String
Private ImovelDataCadastro As String
Private ImovelQuadra As String
Private ImovelSecao As String
Private ImovelLoteamento As String
Private ImovelEdificio As String
Private ImovelTipoLote As TipoBoletim
Private ImovelTipoBoletim As String
Private ImovelOcupanteNome As String
Private ImovelOcupanteCPF As String
Private ImovelBloco             As String
Private InscricaoReduzida    As String

'-OPCIONAIS-
Private ImovelCodMensagem As String
Private ImovelZona As String
Private ImovelAforamentoNumero As String
Private ImovelAforamentoFicha As String
Private ImovelAforamentoLivro As String
Private ImovelAforamentoFolha As String
Private ImovelAforamentoData As String
Private ImovelAforamentoRegistro As String
Private ImovelAforamentoRegistroData As String
Private AnoAquisicao As Integer
Private Rg As String
Private ImovelSalaLoja As String
Private ImovelApart As String


Dim cadastro As New VSImposto

Public Type EnderecoBC
    Complemento As String
    SalaLoja As String
    Apto As String
    Bloco As String
End Type

Public Type ProprietarioBC
    Inscricao As String
    Nome As String
    Logradouro As String
    NomeLogradouro As String
    Bairro As String
    Numero As String
    Complemento As String
    Cep As String
    Cidade As String
    UF As String
    Cpf As String
    Rg As String
End Type

Public Type BoletimCondominio
    Endereco As EnderecoBC
    Proprietario As ProprietarioBC
End Type

Public Type Edificacao
    Inscricao As String
    Unidade As Integer
    Subjetivo(1 To 200) As Boolean
    Componente(1 To 200) As String
    Deletado As Boolean
    Bc As BoletimCondominio
End Type

Public Sub CarregaDadosContribuinte(sInscricao As String, sNome As String, sCpf As String, sCodLogr As String, sTipoLogr As String, _
                                    sLogr As String, sNumero As String, sCompl As String, sCodBairro As String, sBairro As String, _
                                    sCEP As String, sCidade As String, sUF As String, Optional sRg As String)
                    
    ContribuinteInscricao = sInscricao
    ContribuinteNome = sNome
    ContribuinteCPFcgc = sCpf
    ContribuinteEnderecoCodLogradouro = sCodLogr
    ContribuinteEnderecoTipoLogradouro = sTipoLogr
    ContribuinteEnderecoLogradouro = sLogr
    ContribuinteEnderecoNumero = sNumero
    ContribuinteEnderecoCompl = sCompl
    ContribuinteEnderecoCodBairro = sCodBairro
    ContribuinteEnderecoBairro = sBairro
    ContribuinteCEP = sCEP
    ContribuinteCidade = sCidade
    ContribuinteUF = sUF
    Rg = sRg
End Sub

Public Sub CarregaDadosImovel(sInscricao As String, sInscricaoAnterior As String, sUnidade As String, sSubUnidade As String, sInscricaoCondominio As String, _
                                sInscricaoAnteriorBC As String, sCodLogr As String, sCodBairro As String, sNumero As String, _
                                sCompl As String, sLote As String, sQuadra As String, sLoteamento As String, sTipoLote As TipoBoletim, _
                                sOcupanteNome As String, sOcupanteCPF As String, Optional sCodMensagem As String, _
                                Optional sZona As String, Optional sAforamentoNumero As String, Optional sAforamentoFicha As String, _
                                Optional sAforamentoLivro As String, Optional sAforamentoFolha As String, _
                                Optional sAforamentoReg As String, Optional sAforamentoData As String, Optional sAforamentoRegistroData As String, Optional sBloco As String, Optional sIncricaoReduzida As String, Optional sSecao As String, Optional sEdificio As String, Optional sDataCadastro As String, Optional Apartamento As String, Optional SalaLoja As String, Optional sAnoAquisicao As Integer, Optional sTipoLogradouro As String, Optional sFracaoIdeal As String)
     '-OBRIGATORIAS-
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If sIncricaoReduzida = "" Then
            InscricaoReduzida = sInscricao
            ImovelInscricao = sInscricao
        Else
            ImovelInscricao = sInscricao
            InscricaoReduzida = sIncricaoReduzida
        End If
     Else
        ImovelInscricao = sInscricao
     End If
    
    ImovelInscricaoAnterior = sInscricaoAnterior
    ImovelUnidade = sUnidade
    ImovelSubUnidade = sSubUnidade
    ImovelFracaoIdeal = sFracaoIdeal
    ImovelInscricaoCondominio = sInscricaoCondominio
    ImovelInscricaoAnteriorBC = sInscricaoAnteriorBC
    ImovelCodLogr = sCodLogr
    ImovelCodTipoLogr = sTipoLogradouro
    ImovelCodBairro = sCodBairro
    ImovelEnderecoNumero = sNumero
    ImovelEnderecoCompl = sCompl
    ImovelLote = sLote
    ImovelSecao = sSecao
    ImovelQuadra = sQuadra
    ImovelLoteamento = sLoteamento
    ImovelEdificio = sEdificio
    ImovelTipoLote = sTipoLote
    ImovelTipoBoletim = IIf(Trim(sInscricaoCondominio) = "", 0, 1)
    ImovelOcupanteNome = sOcupanteNome
    ImovelOcupanteCPF = sOcupanteCPF
    ImovelBloco = sBloco
    ImovelDataCadastro = sDataCadastro
    '-OPCIONAIS-
    ImovelCodMensagem = sCodMensagem
    ImovelZona = sZona
    ImovelAforamentoNumero = sAforamentoNumero
    ImovelAforamentoFicha = sAforamentoFicha
    ImovelAforamentoLivro = sAforamentoLivro
    ImovelAforamentoFolha = sAforamentoFolha
    ImovelAforamentoRegistro = sAforamentoReg
    ImovelAforamentoData = sAforamentoData
    ImovelAforamentoRegistroData = sAforamentoRegistroData
    AnoAquisicao = sAnoAquisicao
    ImovelSalaLoja = SalaLoja
    ImovelApart = Apartamento
End Sub

Sub GravaUnidadeDetalhe(IC As String, CodComponente As String, CodGrupo As String, Unidade As String, Valor As Double, SubUnidade As Integer)
    Dim Valores As String
    Dim Campos As String
    On Error Resume Next
    
    Valores = Bdados.PreparaValor(Bdados.Converte(IC, tctexto), CodComponente, Bdados.Converte(Valor, TCDuplo), CodGrupo, Unidade, SubUnidade)
    Campos = "tdi_tim_ic,tdi_tco_cod_componente,tdi_valor_item,tdi_tgc_cod_grupo,tdi_tim_ic_unidade,tdi_tim_sub_unidade"
    Call Bdados.InsereDados("Tab_Detalhe_Imovel", Valores, Campos)
End Sub

Public Sub GravaComponentes(InscCadastral As String, Formulario As Object, ComponenteInicial As Integer, ComponenteFinal As Integer, Texto As Boolean, Unidade As Integer, Optional SubUnidade As Integer = 0)
    Dim Controle As Control
    Dim Rs As VSRecordset
    Dim CodItem As String
    Dim ValorItem As String
    Dim Valores As String
    Dim Campos As String
    Dim ValorGerado As Single
    
    For Each Controle In Formulario.Controls
        If IsNumeric(Controle.Tag) Then
            If (Val(Controle.Tag) >= ComponenteInicial And Val(Controle.Tag) <= ComponenteFinal) Then
                If Not Texto Then
                   'CodItem = cadastro.BuscaCodItemAvancado(Controle.Text, Controle.Tag)
                    CodItem = ParseString(Controle.Text, "-", 1)
                    ValorItem = Controle.ListIndex
                   'novo
                Else
                    CodItem = Controle.Tag
                    ValorItem = Nvl(Controle.Text, 0)
                End If
                
                If Trim(CodItem) <> "" Then
                    Valores = Bdados.PreparaValor(Bdados.Converte(InscCadastral, tctexto), CodItem, Bdados.Converte(Nvl(ValorItem, 0), Monetario), Controle.Tag, Unidade, SubUnidade)
                    Campos = "tdi_tim_ic,tdi_tco_cod_componente, tdi_valor_item,tdi_tgc_cod_grupo,tdi_tim_ic_unidade,tdi_tim_sub_unidade"
                    Call Bdados.InsereDados("TAB_DETALHE_IMOVEL", Valores, Campos) ', _
                        "tdi_tim_ic='" & InscCadastral & "' and  tdi_tco_cod_componente =" & CodItem & _
                        " and tdi_tgc_cod_grupo=" & Controle.Tag & _
                        " and tdi_tim_ic_unidade=" & Unidade & " and tdi_tim_sub_unidade = " & SubUnidade)
                    Bdados.FechaTabela Rs
                End If
            End If
        End If
    Next
End Sub

Public Sub GravaComponentesUnidadeProp(InscCadastral As String, Formulario As Object, ComponenteInicial As Integer, ComponenteFinal As Integer, Texto As Boolean, Unidade As Integer, Optional SubUnidade As Integer = 0)
    Dim Controle As Control
    Dim Rs As VSRecordset
    Dim CodItem As String
    Dim ValorItem As String
    Dim Valores As String
    Dim Campos As String
    Dim ValorGerado As Single
    
    For Each Controle In Formulario.Controls
        If IsNumeric(Controle.Tag) Then
            If (Val(Controle.Tag) >= ComponenteInicial And Val(Controle.Tag) <= ComponenteFinal) Then
                If Not Texto Then
                   'CodItem = cadastro.BuscaCodItemAvancado(Controle.Text, Controle.Tag)
                    CodItem = ParseString(Controle.Text, "-", 1)
                    ValorItem = Controle.ListIndex
                   'novo
                Else
                    CodItem = Controle.Tag
                    ValorItem = Controle.Text
                End If
                
                If Trim(CodItem) <> "" Then
                    Valores = Bdados.PreparaValor(Bdados.Converte(InscCadastral, tctexto), CodItem, Bdados.Converte(Nvl(ValorItem, 0), TCDuplo), Controle.Tag, Unidade, SubUnidade)
                    Campos = "tdi_tim_ic,tdi_tco_cod_componente,tdi_valor_item,tdi_tgc_cod_grupo,tdi_tim_ic_unidade,tdi_tim_sub_unidade"
                    Call Bdados.InsereDados("TAB_DETALHE_IMOVEL", Valores, Campos) ', _
                        "tdi_tim_ic='" & InscCadastral & "' and  tdi_tco_cod_componente =" & CodItem & _
                        " and tdi_tgc_cod_grupo=" & Controle.Tag & _
                        " and tdi_tim_ic_unidade=" & Unidade & " and tdi_tim_sub_unidade = " & SubUnidade)
                    Bdados.FechaTabela Rs
                End If
            End If
        End If
    Next
End Sub


Public Sub GravaComponente(InscCadastral As String, Unidade As Integer, CodComponente As String, CodGrupo As Integer, Optional SubUnidade As Integer = 0)
    Dim Valores As String
    Dim Campos As String
    
    Valores = Bdados.PreparaValor(Bdados.Converte(InscCadastral, tctexto), Unidade, SubUnidade, CodGrupo, _
        Bdados.Converte(CodComponente, TCMonetario), CodGrupo)
    
    Campos = "tdi_tim_ic,tdi_tim_ic_unidade,tdi_tim_sub_unidade,tdi_tco_cod_componente,tdi_valor_item,tdi_tgc_cod_grupo"
    Call Bdados.GravaDados("TAB_DETALHE_IMOVEL", Valores, Campos, _
        "tdi_tim_ic='" & InscCadastral & "' " & _
        " and tdi_tgc_cod_grupo=" & CodGrupo & " and tdi_tim_ic_unidade=" & Unidade & _
        " and tdi_tim_sub_unidade = " & SubUnidade)
End Sub

Function InsereTerritorio() As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Rs As VSRecordset
    Dim Valor As Double
    Dim Gravou As Boolean
    
    If valorMercadoTerreno > 0 Then
        Valor = valorMercadoTerreno + valorMercadoEdificacao
    Else
        Valor = ValorTerreno + ValorEdificacao
    End If
'    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
'        AtualizaRegistroPagamentos ContribuinteInscricao, InscricaoReduzida, ImovelUnidade
'    Else
'        AtualizaRegistroPagamentos ContribuinteInscricao, ImovelInscricao, ImovelUnidade
'    End If
    InsereTerritorio = True
    'Vou cadastrar o imovel
        
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
    
        Campos = "tim_ic,TIM_BLOCO, TIM_UNIDADE, tim_tci_im , tim_tlg_cod_logradouro, " _
        & "tim_numero,  tim_complemento, tim_loteamento, tim_quadra," _
        & "tim_lote, tim_cep,tim_tipo_imovel," _
        & "tim_ocupante,tim_cgc_cpf_ocupante,tim_ic_anterior, " & _
        " TIM_VALOR_TERRENO_MERCADO, TIM_VALOR_EDIFICACAO_MERCADO," & _
        " tim_ic_auxiliar,tim_tci_im_auxiliar,tim_secao,tim_ted_cod_edificio," & _
        " tim_ano_aquis,TIM_APTO,TIM_SALA_LOJA,TIM_TTL_COD_TIP_LOGR,TIM_TSC_COD_SIT_CAD,TIM_SITUACAO_LOTE"

        Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoReduzida, tctexto), Bdados.Converte(ImovelBloco, tctexto), Bdados.Converte(ImovelUnidade, tctexto), Bdados.Converte(ContribuinteInscricao, tctexto), _
        Bdados.Converte(ImovelCodLogr, tctexto), Bdados.Converte(Trim(ImovelEnderecoNumero), tctexto), Bdados.Converte(Trim(ImovelEnderecoCompl), tctexto), Trim(ImovelLoteamento), _
        Bdados.Converte(Trim(ImovelQuadra), tctexto), Bdados.Converte(Trim(ImovelLote), tctexto), Bdados.Converte(Trim(Temp.PegaParametro(Bdados, "CEP CLIENTE") & _
        "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")), tctexto), ImovelTipoLote, Bdados.Converte(ImovelOcupanteNome, tctexto), _
        Bdados.Converte(ImovelOcupanteCPF, tctexto), Bdados.Converte(Nvl(ImovelInscricaoAnterior, 0), tctexto), _
        valorMercadoTerreno, valorMercadoEdificacao, _
        Bdados.Converte(ImovelInscricao, tctexto), ContribuinteInscricaoAuxiliar, _
        Bdados.Converte(ImovelSecao, tctexto), ImovelEdificio, AnoAquisicao, _
        Bdados.Converte(ImovelApart, tctexto), Bdados.Converte(ImovelSalaLoja, tctexto), _
        ImovelCodTipoLogr, 1, 2)

    Else
    
        Campos = "tim_ic,TIM_BLOCO, TIM_UNIDADE, tim_tci_im , tim_tlg_cod_logradouro, " _
        & "tim_numero,  tim_complemento, tim_loteamento, tim_quadra," _
        & "tim_lote, tim_cep,tim_tipo_imovel," _
        & "tim_ocupante,tim_cgc_cpf_ocupante,tim_ic_anterior, " & _
        " TIM_VALOR_TERRENO_MERCADO, TIM_VALOR_EDIFICACAO_MERCADO,tim_ano_aquis,TIM_APTO,TIM_SALA_LOJA," & _
        "tim_secao,tim_ted_cod_edificio,TIM_TTL_COD_TIP_LOGR,TIM_TSC_COD_SIT_CAD,TIM_SITUACAO_LOTE"

    
        Valores = Bdados.PreparaValor(Bdados.Converte(ImovelInscricao, tctexto), Bdados.Converte(ImovelBloco, tctexto), ImovelUnidade, Bdados.Converte(ContribuinteInscricao, tctexto), _
        ImovelCodLogr, Trim(ImovelEnderecoNumero), Trim(ImovelEnderecoCompl), Trim(ImovelLoteamento), _
        Trim(ImovelQuadra), Bdados.Converte(Trim(ImovelLote), tctexto), Bdados.Converte(Trim(Temp.PegaParametro(Bdados, "CEP CLIENTE") & _
        "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")), tctexto), ImovelTipoLote, ImovelOcupanteNome, _
        Bdados.Converte(ImovelOcupanteCPF, tctexto), Bdados.Converte(Nvl(ImovelInscricaoAnterior, 0), tctexto), _
        valorMercadoTerreno, valorMercadoEdificacao, AnoAquisicao, Bdados.Converte(ImovelApart, tctexto), _
        Bdados.Converte(ImovelSalaLoja, tctexto), ImovelSecao, ImovelEdificio, ImovelCodTipoLogr, 1, 2)
        
    End If
    
    If Trim(ImovelInscricaoCondominio) <> "" Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(ImovelInscricaoCondominio, tctexto))
        Campos = Campos & ",TIM_IC_CONDOMINIO"
    End If
    Valores = Valores & Bdados.PreparaValor(Nvl(ImovelCodMensagem, 0), ImovelCodBairro, Nvl(ImovelZona, 0))
    Campos = Campos & ",TIM_COD_MENSAGEM,TIM_TBA_COD_BAIRRO,TIM_ZONA"
    Campos = Campos & ",TIM_AFORAMENTO_NUMERO,TIM_AFORAMENTO_FICHA,TIM_AFORAMENTO_LIVRO,TIM_AFORAMENTO_FOLHA,TIM_AFORAMENTO_REGISTRO,TIM_SUB_UNIDADE,TIM_TUS_COD_USUARIO,TIM_DATA_CADASTRO,TIM_TIPO_BOLETIM"
    Valores = Valores & Bdados.PreparaValor(ImovelAforamentoNumero, ImovelAforamentoFicha, ImovelAforamentoLivro, _
    ImovelAforamentoFolha, ImovelAforamentoRegistro, ImovelSubUnidade, Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto), Bdados.Converte(ImovelDataCadastro, TCDataHora), IIf(ImovelTipoLote <> tbo_Condominio, 0, 1))
    If Trim(ImovelAforamentoRegistroData) <> "" Then
        Valores = Valores & Bdados.PreparaValor(ImovelAforamentoRegistroData)
        Campos = Campos & ",TIM_DATA_REGISTRO"
    End If
    If Trim(ImovelAforamentoData) <> "" Then
        Campos = Campos & ",TIM_AFORAMENTO_DATA"
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(ImovelAforamentoData, TCDataHora))
    End If
'    Bdados.AbreTrans
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") <> "REDUZIDA" Then
        If Trim(InscricaoReduzida) = "" Then
            Gravou = Bdados.GravaDados("TAB_IMOVEL", Valores, Campos, "tim_ic='" & ImovelInscricao & "'")
        Else
            Gravou = Bdados.GravaDados("TAB_IMOVEL", Valores, Campos, "tim_ic='" & ImovelInscricao & "' or tim_ic_auxiliar ='" & InscricaoReduzida & "'")
        End If
        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC ='" & ImovelInscricao & "'"
    Else
        Gravou = Bdados.GravaDados("TAB_IMOVEL", Valores, Campos, "tim_ic='" & InscricaoReduzida & "'")
        Bdados.DeletaDados "TAB_DETALHE_IMOVEL", "TDI_TIM_IC ='" & InscricaoReduzida & "'"
    End If
    InsereTerritorio = Gravou

    Bdados.FechaTabela Rs
End Function

Function InsereContribuinte(Optional GerarNovo As Boolean) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Campos As String
    Dim Valores As String
    Dim SitCadastral As String
    InsereContribuinte = True
    If Not cadastro.ContribuinteHabilitado(ContribuinteInscricao, SitCadastral) Then
        Call Util.Avisa("O Contribuinte está " & SitCadastral & " e não pode adquirir novos imóveis.")
        Screen.MousePointer = 0
        InsereContribuinte = False
        Exit Function
    End If
    'Verifica se gravou contribuinte
    If Trim(ContribuinteInscricao) = "" Then
        Sql = "Select tci_nome, tci_im,tci_im_auxiliar from tab_contribuinte where tci_nome ='" & Trim(ContribuinteNome) & "'"
    Else
        Sql = "Select tci_nome, tci_im,tci_im_auxiliar from tab_contribuinte where tci_im ='" & Trim(ContribuinteInscricao) & "'"
    End If
    If Bdados.AbreTabela(Sql, Rs) And Trim(ContribuinteInscricao) <> "" Then
        ContribuinteInscricao = Rs!tci_im
        ContribuinteInscricaoAuxiliar = "" & Rs!tci_im_auxiliar
    Else
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            ContribuinteInscricao = Conta.GeraCodPagamento("CADASTRO ECONOMICO")
            ContribuinteInscricaoAuxiliar = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
        Else
            ContribuinteInscricao = cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
            ContribuinteInscricaoAuxiliar = Conta.GeraCodPagamento("INSCRICAO CONTRIBUINTE")
        End If
    End If
    
    Campos = "tci_im, tci_nome ,tci_cod_logradouro, tci_logradouro, " _
        & "tci_nome_logradouro, tci_numero, tci_complemento, tci_cod_bairro," _
        & "tci_bairro, tci_cep , tci_cidade, tci_UF, tci_data_cadastro," _
        & "tci_tsc_cod_sit_cad,tci_tus_cod_usuario,tci_tipo_contribuinte,TCI_IM_AUXILIAR"
    ContribuinteCEP = Edita.TiraTudo(ContribuinteCEP)
    Valores = Bdados.PreparaValor(ContribuinteInscricao, Bdados.Converte(Trim(ContribuinteNome), tctexto), Bdados.Converte(Trim(ContribuinteEnderecoCodLogradouro), tctexto), _
        Bdados.Converte(ContribuinteEnderecoTipoLogradouro, tctexto), Bdados.Converte(Trim(ContribuinteEnderecoLogradouro), tctexto), Trim(ContribuinteEnderecoNumero), _
        Bdados.Converte(Trim(ContribuinteEnderecoCompl), tctexto), Trim(ContribuinteEnderecoCodBairro), _
        Bdados.Converte(Trim(ContribuinteEnderecoBairro), tctexto), Trim(ContribuinteCEP), Bdados.Converte(Trim(ContribuinteCidade), tctexto), Bdados.Converte(ContribuinteUF, tctexto), _
        Bdados.Converte(Date, TCDataHora), 1, Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto), 0, Bdados.Converte(ContribuinteInscricaoAuxiliar, tctexto))
    If Trim(ContribuinteCPFcgc) <> "" Then
        Campos = Campos & ",TCI_CGC_CPF"
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(ContribuinteCPFcgc, tctexto))
    End If
    If Trim(Rg) <> "" Then
        Campos = Campos & ",TCI_RG"
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Rg, tctexto))
    End If
    Call Bdados.GravaDados("Tab_Contribuinte" _
     , Valores, Campos, "tci_im='" & ContribuinteInscricao & "'")
     
    Bdados.FechaTabela Rs
End Function
Public Function LoteCadastrado(Distrito As String, Setor As String, Quadra As String) As Boolean
    Dim Sql As String
    Sql = "SELECT TCL_TUS_COD_USUARIO,TCL_DATA_ABERTURA FROM TAB_CONTROLE_LOTE WHERE " _
            & "TCL_DISTRITO = '" & Distrito _
            & "' AND TCL_SETOR='" & Setor _
            & "' AND TCL_QUADRA = '" & Quadra & "'"
     If Bdados.AbreTabela(Sql) Then LoteCadastrado = True
     
End Function

Public Function LoteFechado(Distrito As String, Setor As String, Quadra As String) As Boolean
Dim Sql As String
Dim Rs As VSRecordset
    Sql = "SELECT TCL_STATUS FROM TAB_CONTROLE_LOTE WHERE " _
            & "TCL_DISTRITO = '" & Distrito _
            & "' AND TCL_SETOR='" & Setor _
            & "' AND TCL_QUADRA = '" & Quadra & "'"
     If Bdados.AbreTabela(Sql, Rs) Then
        LoteFechado = IIf(Rs!TCL_STATUS = 1, False, True)
     Else
        LoteFechado = False
     End If
End Function

Public Function FechaLote(Distrito As String, Setor As String, Quadra As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim BTs As Double
    Dim BPs As Double
    Dim BCs As Double
    
    FechaLote = False
    
    Sql = "Select * from VIS_LOTE_DIGITACAO where TCL_DISTRITO = '" & Distrito & "' AND TCL_SETOR='" & Setor & "' AND TCL_QUADRA = '" & Quadra & "' and TCL_QTD_BT =" & BTs & " and TCL_QTD_BP =" & BPs & " and TCL_QTD_BC = " & BCs
    If Bdados.AbreTabela(Sql, Rs) Then
        FechaLote = True
        Bdados.AtualizaDados "TAB_CONTROLE_LOTE", Bdados.PreparaValor(2), "TCL_STATUS", "TCL_DISTRITO = '" & Distrito & _
            "' AND TCL_SETOR='" & Setor & "' AND TCL_QUADRA = '" & Quadra & _
            "' and TCL_QTD_BT =" & BTs & " and TCL_QTD_BP =" & BPs & " and TCL_QTD_BC = " & BCs
    End If
End Function

Public Sub AtualizaRegistroPagamentos(Contribuinte As String, Imovel As String, Unidade As String)
    If CInt(Nvl(Unidade, 0)) < 200 Then Unidade = 0
    Bdados.AtualizaDados "TAB_GERACAO_TRIBUTO", Bdados.PreparaValor(Contribuinte), "TGT_INSCRICAO", "TGT_TIM_IC ='" & Imovel & "'"
    Bdados.AtualizaDados "tab_conta_contribuinte", Bdados.PreparaValor(Contribuinte), "tcc_im", "tcc_tim_ic ='" & Imovel & "'"
    Bdados.AtualizaDados "tab_darm_recebido", Bdados.PreparaValor(Contribuinte), "TDR_INSCRICAO", "tdr_tim_ic ='" & Imovel & "'"
    Bdados.AtualizaDados "tab_parcelamento", Bdados.PreparaValor(Contribuinte), "TPA_TCI_IM", "TPA_TIM_IC ='" & Imovel & "'"
End Sub

Public Sub GravaHistorico(sInscricaoCadastral As String, sMotivoAlteracao As String)
    '*****************************************Grava Historico do imovel************************************
    Dim Conta                                 As New ContaCorrente
    Dim Sql                                   As String
    Dim CodMudanca                            As Double
    Dim Rs                                    As VSRecordset
    
    CodMudanca = Conta.GeraCodPagamento(0)
    '- TAB_IMOVEL_HISTORICO
       
    Sql = " select TIM_IC,"
    Sql = Sql & " TIM_UNIDADE,"
    Sql = Sql & " TIM_IC_AUXILIAR,"
    Sql = Sql & " TIM_TCI_IM,"
    Sql = Sql & " TIM_IC_ANTERIOR,"
    Sql = Sql & " TIM_TLG_COD_LOGRADOURO,"
    Sql = Sql & " TIM_NUMERO,"
    Sql = Sql & " TIM_COMPLEMENTO,"
    Sql = Sql & " TIM_CEP,"
    Sql = Sql & " TIM_ANO_AQUIS,"
    Sql = Sql & " TIM_VALOR,"
    Sql = Sql & " TIM_TIPO_IMOVEL,"
    Sql = Sql & " TIM_OCUPANTE,"
    Sql = Sql & " TIM_CGC_CPF_OCUPANTE,"
    Sql = Sql & " TIM_IC_CONDOMINIO,"
    Sql = Sql & " TIM_TBA_COD_BAIRRO,"
    Sql = Sql & " TIM_COD_MENSAGEM,"
    Sql = Sql & " TIM_ZONA,"
    Sql = Sql & " TIM_VALOR_TERRENO,"
    Sql = Sql & " TIM_VALOR_EDIFIC,"
    Sql = Sql & " TIM_LOTEAMENTO,"
    Sql = Sql & " TIM_SECAO,"
    Sql = Sql & " TIM_QUADRA,"
    Sql = Sql & " TIM_LOTE,"
    Sql = Sql & " TIM_TSC_COD_SIT_CAD,"
    Sql = Sql & " TIM_AFORADO,"
    Sql = Sql & " TIM_OBS,"
    Sql = Sql & " TIM_AFORAMENTO_REGISTRO,"
    Sql = Sql & " TIM_DATA_REGISTRO,"
    Sql = Sql & " TIM_AFORAMENTO_NUMERO,"
    Sql = Sql & " TIM_AFORAMENTO_FICHA,"
    Sql = Sql & " TIM_AFORAMENTO_LIVRO,"
    Sql = Sql & " TIM_AFORAMENTO_FOLHA,"
    Sql = Sql & " TIM_AFORAMENTO_DATA,"
    Sql = Sql & " TIM_MOTIVO_ALTERACAO,"
    Sql = Sql & " TIM_VALOR_TERRENO_MERCADO,"
    Sql = Sql & " TIM_VALOR_EDIFICACAO_MERCADO,"
    Sql = Sql & " TIM_SITUACAO_LOTE,"
    Sql = Sql & " TIM_SUB_UNIDADE,"
    Sql = Sql & " TIM_TUS_COD_USUARIO,"
    Sql = Sql & " TIM_DATA_CADASTRO,"
    Sql = Sql & " TIM_TIPO_BOLETIM,"
    Sql = Sql & " TIM_TED_COD_EDIFICIO,"
    Sql = Sql & " TIM_BLOCO,"
    Sql = Sql & " TIM_APTO,"
    Sql = Sql & " TIM_SALA_LOJA,"
    Sql = Sql & " TIM_TTC_SEQ_TRECHO,"
    Sql = Sql & " TIM_TTC_COD_TRECHO,"
    Sql = Sql & " TIM_TCI_IM_AUXILIAR,"
    Sql = Sql & " TIM_TTL_COD_TIP_LOGR"
    Sql = Sql & " FROM TAB_IMOVEL where tim_ic = '" & sInscricaoCadastral & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Sql = " insert into tab_imovel_historico (TIM_IC,"
        Sql = Sql & " TIM_UNIDADE,"
        Sql = Sql & " TIM_IC_AUXILIAR,"
        Sql = Sql & " TIM_TCI_IM,"
        Sql = Sql & " TIM_IC_ANTERIOR,"
        Sql = Sql & " TIM_TLG_COD_LOGRADOURO,"
        Sql = Sql & " TIM_NUMERO,"
        Sql = Sql & " TIM_COMPLEMENTO,"
        Sql = Sql & " TIM_CEP,"
        Sql = Sql & " TIM_ANO_AQUIS,"
        Sql = Sql & " TIM_VALOR,"
        Sql = Sql & " TIM_TIPO_IMOVEL,"
        Sql = Sql & " TIM_OCUPANTE,"
        Sql = Sql & " TIM_CGC_CPF_OCUPANTE,"
        Sql = Sql & " TIM_IC_CONDOMINIO,"
        Sql = Sql & " TIM_TBA_COD_BAIRRO,"
        Sql = Sql & " TIM_COD_MENSAGEM,"
        Sql = Sql & " TIM_ZONA,"
        Sql = Sql & " TIM_VALOR_TERRENO,"
        Sql = Sql & " TIM_VALOR_EDIFIC,"
        Sql = Sql & " TIM_LOTEAMENTO,"
        Sql = Sql & " TIM_SECAO,"
        Sql = Sql & " TIM_QUADRA,"
        Sql = Sql & " TIM_LOTE,"
        Sql = Sql & " TIM_TSC_COD_SIT_CAD,"
        Sql = Sql & " TIM_AFORADO,"
        Sql = Sql & " TIM_OBS,"
        Sql = Sql & " TIM_AFORAMENTO_REGISTRO,"
        Sql = Sql & " TIM_DATA_REGISTRO,"
        Sql = Sql & " TIM_AFORAMENTO_NUMERO,"
        Sql = Sql & " TIM_AFORAMENTO_FICHA,"
        Sql = Sql & " TIM_AFORAMENTO_LIVRO,"
        Sql = Sql & " TIM_AFORAMENTO_FOLHA,"
        Sql = Sql & " TIM_AFORAMENTO_DATA,"
        Sql = Sql & " TIM_MOTIVO_ALTERACAO,"
        Sql = Sql & " TIM_VALOR_TERRENO_MERCADO,"
        Sql = Sql & " TIM_VALOR_EDIFICACAO_MERCADO,"
        Sql = Sql & " TIM_SITUACAO_LOTE,"
        Sql = Sql & " TIM_SUB_UNIDADE,"
        Sql = Sql & " TIM_TUS_COD_USUARIO,"
        Sql = Sql & " TIM_DATA_CADASTRO,"
        Sql = Sql & " TIM_TIPO_BOLETIM,"
        Sql = Sql & " TIM_TED_COD_EDIFICIO,"
        Sql = Sql & " TIM_BLOCO,"
        Sql = Sql & " TIM_APTO,"
        Sql = Sql & " TIM_SALA_LOJA,"
        Sql = Sql & " TIM_TTC_SEQ_TRECHO,"
        Sql = Sql & " TIM_TTC_COD_TRECHO,"
        Sql = Sql & " TIM_TCI_IM_AUXILIAR,"
        Sql = Sql & " TIM_TTL_COD_TIP_LOGR,"
        Sql = Sql & " TIM_DATA    ,"
        Sql = Sql & " TIM_MOTIVO,"
        Sql = Sql & " TIM_TUS_USUARIO,TIM_COD_MUDANCA"
        Sql = Sql & "  ) Values ( "
        
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_ic, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_UNIDADE, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_ic_auxiliar, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_tci_im, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_IC_ANTERIOR, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_tlg_cod_logradouro, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_numero, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_complemento, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_CEP, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_ano_aquis, 0)) & "',"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!tim_valor, 0)), TCMonetario) & " ,"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_tipo_imovel, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_OCUPANTE, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_CGC_CPF_OCUPANTE, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_IC_CONDOMINIO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_tba_cod_bairro, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_COD_MENSAGEM, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_ZONA, 0)) & "',"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_VALOR_TERRENO, 0)), TCMonetario) & " ,"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_VALOR_EDIFIC, 0)), TCMonetario) & " ,"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_LOTEAMENTO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_secao, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_QUADRA, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_LOTE, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_TSC_COD_SIT_CAD, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_AFORADO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_OBS, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_AFORAMENTO_REGISTRO, 0)) & "',"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_DATA_REGISTRO, Date)), TCDataHora) & ","
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_AFORAMENTO_NUMERO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_AFORAMENTO_FICHA, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_AFORAMENTO_LIVRO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_AFORAMENTO_FOLHA, 0)) & "',"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_AFORAMENTO_DATA, Date)), TCDataHora) & " ,"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_MOTIVO_ALTERACAO, 0)) & "',"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_VALOR_TERRENO_MERCADO, 0)), TCMonetario) & " ,"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_VALOR_EDIFICACAO_MERCADO, 0)), TCMonetario) & " ,"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_SITUACAO_LOTE, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_sub_unidade, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_TUS_COD_USUARIO, 0)) & "',"
        Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs!TIM_DATA_CADASTRO, Date)), TCDataHora) & ","
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_TIPO_BOLETIM, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!tim_ted_cod_edificio, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_BLOCO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_APTO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_SALA_LOJA, 0)) & "',"
        Sql = Sql & Bdados.Converte(Date, TCDataHora) & ","
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_TTC_COD_TRECHO, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_TCI_IM_AUXILIAR, 0)) & "',"
        Sql = Sql & "'" & Trim(Nvl("" & Rs!TIM_TTL_COD_TIP_LOGR, 0)) & "',"
        Sql = Sql & Bdados.Converte(Date, TCDataHora) & ","
        Sql = Sql & "'" & sMotivoAlteracao & "',"
        Sql = Sql & "'" & AplicacoesVTFuncoes.Usuario & "',"
        Sql = Sql & "'" & CodMudanca & "')"
        
        Bdados.Executa Sql
    End If
'
'    Sql = "Select tab_imovel.* , "
'    Sql = Sql & Bdados.Converte(Date, TCDataHora) & ","
'    Sql = Sql & Bdados.Converte(Time, tctexto) & ","
'    Sql = Sql & Bdados.Converte(Trim(sMotivoAlteracao), tctexto) & ","
'    Sql = Sql & Bdados.Converte(AplicacoesVTFuncoes.Usuario, tctexto) & ","
'    Sql = Sql & Bdados.Converte(CodMudanca, tctexto)
'    Sql = Sql & " from tab_imovel where tim_ic = '" & sInscricaoCadastral & "'"
'    Bdados.Executa "Insert into tab_imovel_historico " & Sql
'
'    '- TAB_DETALHE_IMOVEL_HISTORICO
    Rem Sql = "Select tab_detalhe_imovel.* ," & CodMudanca & " from tab_detalhe_imovel where TDI_tim_ic = '" & sInscricaoCadastral & "'"
    Rem Bdados.Executa "Insert into tab_detalhe_imovel_historico  " & Sql
    '*****************************************FIM DO HISTORICO*********************************************
    Sql = " select TDI_TIM_IC,TDI_TIM_IC_UNIDADE,TDI_TGC_COD_GRUPO,"
    Sql = Sql & " TDI_TCO_COD_COMPONENTE, TDI_VALOR_CALC,TDI_VALOR_ITEM,"
    Sql = Sql & " TDI_TIM_SUB_UNIDADE"
    Sql = Sql & " from tab_detalhe_imovel"
    Sql = Sql & " WHERE TDI_tim_ic = '" & sInscricaoCadastral & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Do Until Rs.EOF
            Sql = " Insert into tab_detalhe_imovel_historico(  TDI_TIM_IC,TDI_TIM_IC_UNIDADE,"
            Sql = Sql & " TDI_TGC_COD_GRUPO,"
            Sql = Sql & " TDI_TCO_COD_COMPONENTE, TDI_VALOR_CALC,TDI_VALOR_ITEM,"
            Sql = Sql & " TDI_TIM_SUB_UNIDADE,TDI_TIM_COD_MUDANCA ) Values ( "
            Sql = Sql & "'" & Rs.Fields("TDI_TIM_IC") & "',"
            Sql = Sql & "'" & Rs.Fields("TDI_TIM_IC_UNIDADE") & "',"
            Sql = Sql & "'" & Rs.Fields("TDI_TGC_COD_GRUPO") & "',"
            Sql = Sql & "'" & Rs.Fields("TDI_TCO_COD_COMPONENTE") & "',"
            Sql = Sql & Bdados.Converte(Trim(Nvl("" & Rs.Fields("TDI_VALOR_CALC"), 0)), TCMonetario) & " ,"
            Sql = Sql & Bdados.Converte(Rs.Fields("TDI_VALOR_ITEM"), TCMonetario) & " ,"
            Sql = Sql & "'" & Rs.Fields("TDI_TIM_SUB_UNIDADE") & "',"
            Sql = Sql & "'" & CodMudanca & "')"
            Bdados.Executa Sql
           Rs.MoveNext
        Loop
    End If
    
End Sub
Public Function ApagaDetalheImovel(sInscricaoCadastral As String) As Boolean
    ApagaDetalheImovel = Bdados.DeletaDados("TAB_DETALHE_IMOVEL", "TDI_TIM_IC = '" & sInscricaoCadastral & "'")
End Function

Public Function ApagaImovel(sInscricaoCadastral As String) As Boolean
    Dim Rs As VSRecordset
    
    If Bdados.AbreTabela("SELECT TIM_VALOR_TERRENO, TIM_VALOR_EDIFIC, TIM_VALOR_TERRENO_MERCADO, TIM_VALOR_EDIFICACAO_MERCADO FROM TAB_IMOVEL WHERE TIM_IC='" & sInscricaoCadastral & "'", Rs) Then
        ValorTerreno = Nvl("" & Rs!TIM_VALOR_TERRENO, 0)
        ValorEdificacao = Nvl("" & Rs!TIM_VALOR_EDIFIC, 0)
        valorMercadoTerreno = Nvl("" & Rs!TIM_VALOR_TERRENO_MERCADO, 0)
        valorMercadoEdificacao = Nvl("" & Rs!TIM_VALOR_EDIFICACAO_MERCADO, 0)
    End If
    Bdados.FechaTabela Rs
    ApagaImovel = Bdados.DeletaDados("TAB_IMOVEL", "tim_ic = '" & sInscricaoCadastral & "'")
End Function

Public Function VerificaFechamentoAreas(Lista As Object) As Boolean
    Dim AreaConstruidaTotal As Double
    Dim AreaConstruidaSomatorio As Double
    AreaConstruidaSomatorio = 0
    Dim j As Integer
    VerificaFechamentoAreas = True
    If Lista.ListItems.Count > 0 Then 'Se definiu alguma edificacao
        AreaConstruidaTotal = Util.Nvl(Lista.ListItems(1).SubItems(12), 0)
        For j = 1 To Lista.ListItems.Count 'Para cada edificacao
            Lista.ListItems(j).Selected = True
            If CStr(CDbl(Nvl(Lista.SelectedItem.SubItems(12), 0))) <> CStr(CDbl(AreaConstruidaTotal)) Then
                Avisa "Foram informados diferentes valores de área total contruída."
                VerificaFechamentoAreas = False
                Screen.MousePointer = 0
                Exit Function
            Else
                AreaConstruidaSomatorio = AreaConstruidaSomatorio + Nvl(Lista.SelectedItem.SubItems(11), 0)
            End If
        Next
        If AplicacoesVTFuncoes.Municipio <> "PETROLINA" Then
            If CStr(CDbl(AreaConstruidaSomatorio)) <> CStr(CDbl(AreaConstruidaTotal)) Then
                Avisa "Somatório das áreas construídas das unidades não batem com a área construída total informada."
                Screen.MousePointer = 0
                VerificaFechamentoAreas = False
                Exit Function
            End If
        End If
    End If
End Function

Public Function VerificaDigitacaoBP(Lista As Object, CodMensagem As String, ObjTab As Object) As Boolean
    If Lista.ListItems.Count = 0 Then 'Nao definiu nenhum predio
        If CDbl(Nvl(CodMensagem, 0)) <> 98 And CDbl(Nvl(CodMensagem, 0)) <> 99 Then
            Avisa "Falta definição do Boletim Predial."
            ObjTab.Tab = 2
            Screen.MousePointer = 0
            VerificaDigitacaoBP = False
            Exit Function
        End If
    End If
    VerificaDigitacaoBP = True
End Function
    
Public Function AtualizaValoresMercado(InscricaoCadastral As String, ValorTerreno As String, ValorEdificacao As String) As Boolean
    Dim Valores As String
    Dim Campos As String
    Campos = "TIM_VALOR_TERRENO_MERCADO,TIM_VALOR_EDIFICACAO_MERCADO"
    Valores = Bdados.PreparaValor(Bdados.Converte(ValorTerreno, TCMonetario), Bdados.Converte(ValorEdificacao, TCMonetario))
    AtualizaValoresMercado = Bdados.GravaDados("TAB_IMOVEL", Valores, Campos, "tim_ic='" & InscricaoCadastral & "'")
End Function

Public Function BuscaLogradouro(sCodLogr As Object, sTipoLogr As Object, sNomeLogr As Object, Optional sMunici As Object, Optional sCEP As Object, Optional sUF As Object) As Boolean
    Dim Query As String
    Dim Rs As VSRecordset
    On Error Resume Next
    If Trim(sCodLogr) <> "" Then
        Query = "SELECT TAB_TIPO_LOGR.TTL_NOME, TAB_LOGRADOURO.tlg_nome, " & _
            " TAB_BAIRRO.TBA_NOME FROM TAB_LOGRADOURO, TAB_BAIRRO,TAB_TIPO_LOGR  " & _
            " where TAB_LOGRADOURO.tlg_tba_cod_bairro = TAB_BAIRRO.TBA_COD_BAIRRO and " & _
             " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr = TAB_TIPO_LOGR.TTL_COD_TIP_LOGR and TLG_COD_LOGRADOURO ='" & sCodLogr.Text & "' and tlg_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio
        If Bdados.AbreTabela(Query, Rs) Then
            sTipoLogr = Rs(0)
            sNomeLogr = Rs(1)
            If Not sMunici Is Nothing Then sMunici = AplicacoesVTFuncoes.Municipio
            BuscaLogradouro = True
        End If
        Bdados.FechaTabela Rs
        
        If Not sCEP Is Nothing Then sCEP = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
        If Not sUF Is Nothing Then sUF = Temp.PegaParametro(Bdados, "ESTADO CLIENTE")
    End If
End Function

Public Function BuscaBairro(sCodBairro As Object, sNomBairro As Object) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    If Trim(sCodBairro) <> "" Then
        Sql = " select TBA_NOME from TAB_BAIRRO where tba_cod_bairro=" & sCodBairro.Text & " and tba_tmu_cod_municipio=" & AplicacoesVTFuncoes.Codigo_Municipio
        If Bdados.AbreTabela(Sql, Rs) Then
            sNomBairro = Rs(0)
        Else
            Avisa "Bairro inexistente."
            sCodBairro.SetFocus
            Exit Function
        End If
    End If
    Bdados.FechaTabela Rs
End Function

Public Function ValorVenal(Inscricao As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select tim_valor from tab_imovel where tim_ic ='" & Inscricao & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        ValorVenal = Nvl("" & Rs!tim_valor, 0)
    End If
End Function

Public Sub GravaBoletimTerritorial(TreeView As Object, Inscricao As String, Unidade As String)
    Dim I                 As Integer
    Dim CodigoComponente  As String
    Dim Campos            As String
    Dim Valores           As String
    Dim Base              As Integer
    
    Campos = "tdi_tim_ic,tdi_tim_ic_unidade,tdi_tgc_cod_grupo,tdi_tco_cod_componente,tdi_valor_calc,tdi_valor_item,tdi_tim_sub_unidade"
'    Bdados.AbreTrans
    For I = 1 To TreeView.NodesCollection.Count
        If IsNumeric(Left(TreeView.NodesCollection(I).Key, 3)) Then
            If TreeView.Value(I) = 1 Then
                Valores = Bdados.PreparaValor(Bdados.Converte(Inscricao, tctexto), Unidade, CInt(Mid(TreeView.NodesCollection(I).Key, 4, 3)), _
                    Left(TreeView.NodesCollection(I).Key, 3), 0, Left(TreeView.NodesCollection(I).Key, 3), 0)
                If Bdados.InsereDados("TAB_DETALHE_IMOVEL", Valores, Campos) Then
                    Base = Base + 1
                End If
                ', "tdi_tim_ic ='" & Inscricao & _
                    "'AND tdi_tim_ic_unidade =" & Unidade & " AND tdi_tgc_cod_grupo=" & _
                    CInt(Mid(TreeView.NodesCollection(I).Key, 4, 3)) & " AND tdi_tco_cod_componente =" & _
                    Left(TreeView.NodesCollection(I).Key, 3) then
                    
            End If
        End If
    Next
'    If Base = TreeView.NodesCollection.Count Then
'        Bdados.GravaTrans
'    Else
'        Bdados.GravaTrans
'    End If
End Sub

Public Sub GravaBoletimPredial(Edificacoes() As Edificacao, Inscricao As String)
    Dim unidades As Integer
    Dim grupos As Integer
    Dim Condicao As String
    Dim Campos As String
    Dim Valores As String
    Campos = "tdi_tim_ic,tdi_tim_ic_unidade,tdi_tgc_cod_grupo,tdi_tco_cod_componente,tdi_valor_calc," & _
                "tdi_valor_item,tdi_tim_sub_unidade"
    For unidades = LBound(Edificacoes) To UBound(Edificacoes) 'Para Cada unidade edificada
        If Not Edificacoes(unidades).Deletado Then
            For grupos = 1 To UBound(Edificacoes(unidades).Componente) 'Para Cada Grupo de Componente da Edificacao(Estado,Padrao,etc..)
                If Trim(Edificacoes(unidades).Componente(grupos)) <> "" Then
                    If Edificacoes(unidades).Subjetivo(grupos) Then 'Valores digitados, não escolhiveis
                        Valores = Bdados.PreparaValor(Bdados.Converte(Inscricao, tctexto), unidades, grupos, grupos, _
                                0, Edificacoes(unidades).Componente(grupos), 0)
                         Condicao = "tdi_tim_ic ='" & Inscricao & "'AND tdi_tim_ic_unidade =" & unidades & _
                            " AND tdi_tgc_cod_grupo=" & grupos
                        
                    Else
                        Valores = Bdados.PreparaValor(Bdados.Converte(Inscricao, tctexto), unidades, grupos, Edificacoes(unidades).Componente(grupos), _
                                 0, Edificacoes(unidades).Componente(grupos), 0)
                        Condicao = "tdi_tim_ic ='" & Inscricao & "'AND tdi_tim_ic_unidade =" & unidades & _
                                " AND tdi_tgc_cod_grupo=" & grupos & " AND tdi_tco_cod_componente =" & _
                                Bdados.Converte(Edificacoes(unidades).Componente(grupos), TCMonetario)
                    End If
                    Bdados.GravaDados "TAB_DETALHE_IMOVEL", Valores, Campos, Condicao
                End If
            Next
        End If
    Next
End Sub

Public Sub GravaBP(Lista As Object, CodMensagem As String, InscricaoSemUnidade As String, Unidade)
    Dim InscricaoCadastral As String
    Dim j As Integer
    Dim I As Integer
    Dim SubUnidade As String
    
    If Lista.ListItems.Count = 0 Then Exit Sub
    If CInt(Nvl(CStr(Unidade), 0)) = 200 Or CInt(Nvl(CStr(Unidade), 0)) = 600 Then
        If Lista.ListItems.Count > 0 Then
            For j = 1 To Lista.ListItems.Count 'Para cada edificacao
                Lista.ListItems(j).Selected = True
                 ApagaDetalheImovel InscricaoSemUnidade & Format(Lista.SelectedItem, "000")
            Next
        End If
    End If
    If CDbl(Nvl(CodMensagem, 0)) <> 98 And CDbl(Nvl(CodMensagem, 0)) <> 99 Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            InscricaoCadastral = InscricaoSemUnidade & Format(Unidade, "000")
        Else
            InscricaoCadastral = InscricaoSemUnidade
        End If
        If Lista.ListItems.Count > 0 Then 'Se definiu alguma edificacao
            For j = 1 To Lista.ListItems.Count 'Para cada edificacao
                Lista.ListItems(j).Selected = True
                For I = 1 To Lista.ColumnHeaders.Count 'Para cada item da edificacao
                    If Trim(Lista.ColumnHeaders(I).Tag) <> "" Then
                        If CInt(Nvl(CStr(Unidade), 0)) >= 200 Then InscricaoCadastral = InscricaoSemUnidade & Format(Lista.SelectedItem, "000")
                        If Lista.ColumnHeaders.Count = 15 Then
                            SubUnidade = Lista.SelectedItem.ListSubItems(14).Text
                        Else
                            SubUnidade = 0
                        End If
                        If CInt(Lista.ColumnHeaders(I).Tag) < 110 Then
                            GravaUnidadeDetalhe InscricaoCadastral, Lista.SelectedItem.ListSubItems(I - 1), Lista.ColumnHeaders(I).Tag, Lista.SelectedItem, Nvl(Lista.SelectedItem.ListSubItems(I - 1), 0), CInt(SubUnidade)
                        Else
                            GravaUnidadeDetalhe InscricaoCadastral, Lista.ColumnHeaders(I).Tag, Lista.ColumnHeaders(I).Tag, Lista.SelectedItem, Nvl(Lista.SelectedItem.ListSubItems(I - 1).Text, 0), CInt(SubUnidade)
                        End If
                        DoEvents
                    End If
                Next
            Next
        End If
    End If
End Sub


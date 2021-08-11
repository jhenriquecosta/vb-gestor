VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=============================================================='
'      VsTIss.cRetencao: Usado em cRetencao,cSubstituicao,TRET401,TRET101,TSUB101 '
'                    Feito por Djeison  -  23/04/2003                                                                            '
'=============================================================='
Option Explicit
Private m_IM_CPF As String
Private m_Nome_Empresa As String
Private m_Endereco As cEndereco
Private m_Atividade As String
Private m_Nota_fiscal As String
Private m_Data_Emissao As String
Private m_Periodo_Ref As String
Private m_Total_Nota As Double
Private m_Valor_Material_ICMS As Double
Private m_Base_Calculo As Double
Private m_ISS_Devido As Double
Private m_Usuario As String
Private m_NumIM As String
Private m_NumCGC As String
Private m_Aliquota As Double
Private m_Cod_Imposto As String
Private m_Nome_Imposto As String
Private m_Cod_Pagamento As String
Private m_Data_Recolhimento As String
Private m_Data_Venc As String
Private m_Multa As String
Private m_Juros As String
Private m_Arrecadacao As String

Public Property Get Arrecadacao() As String
    Arrecadacao = m_Arrecadacao
End Property

Public Property Let Arrecadacao(ByVal Value As String)
    m_Arrecadacao = Value
End Property


Public Property Get Juros() As String
    Juros = m_Juros
End Property

Public Property Let Juros(ByVal Value As String)
    m_Juros = Value
End Property

Public Property Get Multa() As String
    Multa = m_Multa
End Property

Public Property Let Multa(ByVal Value As String)
    m_Multa = Value
End Property

Public Property Get Data_Venc() As String
    Data_Venc = m_Data_Venc
End Property

Public Property Let Data_Venc(ByVal Value As String)
    m_Data_Venc = Value
End Property

Public Property Get Data_Recolhimento() As String
    Data_Recolhimento = m_Data_Recolhimento
End Property

Public Property Let Data_Recolhimento(ByVal Value As String)
    m_Data_Recolhimento = Value
End Property

Public Property Get Cod_Pagamento() As String
    Cod_Pagamento = m_Cod_Pagamento
End Property

Public Property Let Cod_Pagamento(ByVal Value As String)
    m_Cod_Pagamento = Value
End Property

Public Property Get Nome_Imposto() As String
    Nome_Imposto = m_Nome_Imposto
End Property

Public Property Let Nome_Imposto(ByVal Value As String)
    m_Nome_Imposto = Value
End Property

Public Property Get Cod_Imposto() As String
    Cod_Imposto = m_Cod_Imposto
End Property

Public Property Let Cod_Imposto(ByVal Value As String)
    m_Cod_Imposto = Value
End Property

Public Property Get Aliquota() As Double
    Aliquota = m_Aliquota
End Property

Public Property Let Aliquota(ByVal Value As Double)
    m_Aliquota = Value
End Property

Public Property Let NumIM(ByVal Value As String)
    m_NumIM = Value
End Property

Public Property Get NumIM() As String
    NumIM = m_NumIM
End Property

Public Property Let NumCGC(ByVal Value As String)
    m_NumCGC = Value
End Property

Public Property Get NumCGC() As String
    NumCGC = m_NumCGC
End Property

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal Value As String)
    m_Usuario = Value
End Property

Public Property Get ISS_Devido() As Double
    ISS_Devido = m_ISS_Devido
End Property

Public Property Let ISS_Devido(ByVal Value As Double)
    m_ISS_Devido = Value
End Property

Public Property Get Base_Calculo() As Double
    Base_Calculo = m_Base_Calculo
End Property

Public Property Let Base_Calculo(ByVal Value As Double)
    m_Base_Calculo = Value
End Property

Public Property Get Valor_Material_ICMS() As Double
    Valor_Material_ICMS = m_Valor_Material_ICMS
End Property

Public Property Let Valor_Material_ICMS(ByVal Value As Double)
    m_Valor_Material_ICMS = Value
End Property

Public Property Get Total_Nota() As Double
    Total_Nota = m_Total_Nota
End Property

Public Property Let Total_Nota(ByVal Value As Double)
    m_Total_Nota = Value
End Property

Public Property Get Periodo_Ref() As String
    Periodo_Ref = m_Periodo_Ref
End Property

Public Property Let Periodo_Ref(ByVal Value As String)
    m_Periodo_Ref = Value
End Property

Public Property Get Data_emissao() As String
    Data_emissao = m_Data_Emissao
End Property

Public Property Let Data_emissao(ByVal Value As String)
    m_Data_Emissao = Value
End Property

Public Property Get Nota_fiscal() As String
    Nota_fiscal = m_Nota_fiscal
End Property

Public Property Let Nota_fiscal(ByVal Value As String)
    m_Nota_fiscal = Value
End Property

Public Property Get Atividade() As String
    Atividade = m_Atividade
End Property

Public Property Let Atividade(ByVal Value As String)
    m_Atividade = Value
End Property

Public Property Get Endereco() As cEndereco
    Set Endereco = m_Endereco
End Property

Public Property Set Endereco(ByVal Value As cEndereco)
    Set m_Endereco = Value
End Property

Public Property Get Nome_Empresa() As String
    Nome_Empresa = m_Nome_Empresa
End Property

Public Property Let Nome_Empresa(ByVal Value As String)
    m_Nome_Empresa = Value
End Property

Public Property Get IM_CPF() As String
    IM_CPF = m_IM_CPF
End Property

Public Property Let IM_CPF(ByVal Value As String)
    m_IM_CPF = Value
End Property

Private Sub Class_Initialize()
    Set m_Endereco = New cEndereco
End Sub

Private Sub Class_Terminate()
    Set m_Endereco = Nothing
End Sub

Public Sub GravaNovoRemetente()
    Dim Valores As String
    Dim Campos As String
    Valores = Bdados.PreparaValor(m_IM_CPF, m_Nome_Empresa, m_Endereco.Endereco, m_Endereco.Bairro, m_Endereco.CEP, m_Endereco.Municipio, m_Endereco.UF, m_Usuario)
    Campos = "tca_identidade,tca_nome,tca_endereco,tca_bairro,tca_cep,tca_cidade,tca_uf,tca_tus_cod_usuario"
    Bdados.InsereDados "Tab_Contribuinte_Avulso", Valores, Campos
End Sub

Public Function Buscar(TXT As Object, Im As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Imposto As New VSImposto

    If Trim(Im) = "" Then Exit Function
    Sql = "Select * FROM tab_Contribuinte where tci_im='" & Im & "' or tci_cgc_cpf='" & Im & "' and tci_tsc_cod_sit_cad =1"
    
    Buscar = False
    If Bdados.AbreTabela(Sql, rs) Then
        
        If IsNull(rs!tci_tae_cae) Then
            Avisa "Contribuinte sem obrigação de recolhimento de ISSQN."
            Screen.MousePointer = 0
'            TXT.Enabled = True
'            TXT.SetFocus
'            Exit Function
        End If
        m_Nome_Empresa = "" & rs!tci_nome
        m_Endereco.Endereco = "" & rs!tci_logradouro & " " & rs!tci_nome_logradouro & ", " & rs!tci_NUMERO & IIf(Trim(rs!tci_COMPLEMENTO) = "", "", ", " & rs!tci_COMPLEMENTO)
        m_Endereco.Bairro = "" & rs!tci_BAIRRO
        m_Endereco.CEP = "" & rs!tci_cep
        m_Endereco.UF = "" & rs!tci_uf
        m_Endereco.Municipio = "" & rs!tci_cidade
        m_NumCGC = "" & rs!TCI_CGC_CPF
        m_NumIM = "" & rs!tci_im
        m_Atividade = "" & rs!tci_tae_cae
       
        Buscar = True
        
    Else '===============================================================

        Sql = "Select * FROM tab_Contribuinte_Avulso where tca_identidade='" & Im & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            m_Nome_Empresa = "" & rs!tca_nome
            m_Endereco.Endereco = "" & rs!tca_endereco
            m_Endereco.Bairro = "" & rs!tca_bairro
            m_Endereco.CEP = "" & rs!tca_cep
            m_Endereco.UF = "" & rs!tca_uf
            m_Endereco.Municipio = "" & rs!tca_cidade
            m_Atividade = "" & rs!tci_tae_cae
            m_NumCGC = Im
            Buscar = True
        End If
        
    End If
    Bdados.FechaTabela rs
End Function

Public Function PreencherGrid(Grid As Object, Optional P_Inicial As String, Optional P_Final As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    
    Sql = "Select  tna_cod_pagamento AS Cod_Pagamento, "
    Sql = Sql & " tna_tca_identidade_remetente AS Emitente, "
    Sql = Sql & " tna_numero_nota as Num_Nota, "
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = Sql & " (right(tna_periodo,2) " & Bdados.Concatena & " '/' " & Bdados.Concatena & "  left(tna_periodo,4)) as Periodo, "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = Sql & " (RPAD(tna_periodo,2) " & Bdados.Concatena & " '/' " & Bdados.Concatena & "  LPAD(tna_periodo,4)) as Periodo, "
    End If
    Sql = Sql & " tna_data_emissao as Dt_Emissão, "
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = Sql & Bdados.Converte("tna_valor_nota", TCDuplo) & " as Vl_Nota, "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = Sql & " TO_NUMBER(tna_valor_nota,'999999.99')  as Vl_Nota ,"
    End If
    Sql = Sql & " tna_valor_imposto AS Vl_Imposto "
    Sql = Sql & " from tab_nota_avulsa "
    
    Condicao = ""
    If Trim(P_Inicial) <> "" Then Condicao = " and tna_periodo >=" & Right(P_Inicial, 4) & Left(P_Inicial, 2)
    If Trim(P_Final) <> "" Then Condicao = " and tna_periodo <=" & Right(P_Final, 4) & Left(P_Final, 2)
    
    If Condicao <> "" Then Sql = Sql & " where " & Right(Condicao, Len(Condicao) - 4)
    If Bdados.AbreTabela(Sql, rs) Then
        Grid.Preencher Bdados, Sql
        If Condicao <> "" Then
            Condicao = "where " & Right(Condicao, Len(Condicao) - 4)
        End If
        Sql = "SELECT SUM(tna_valor_imposto),COUNT(*) from tab_nota_avulsa " & Condicao
        If Bdados.AbreTabela(Sql, rs) Then
            m_Arrecadacao = "Valor Arrecadado: R$" & Format(rs(0), Const_Monetario)
        Else
            m_Arrecadacao = "Valor Arrecadado: R$0,00"
        End If
        PreencherGrid = True
    Else
        PreencherGrid = False
    End If
    Bdados.FechaTabela rs
End Function

Public Sub PreencherCboAtividade(Combo As Object)
    Dim Sql As String
    Sql = "SELECT  tae_nome, tae_cae FROM TAB_ATIVIDADE_ECONOMICA order by tae_nome"
    Combo.Preencher Bdados, Sql
End Sub

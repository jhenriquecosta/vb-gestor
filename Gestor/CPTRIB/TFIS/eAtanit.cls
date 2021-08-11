VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eAtanit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public procedimento As New eProcedimentosFiscais
Private a_CodPerfilNaoIncidencia As String

Public Property Get CodPerfilNaoIncidencia() As String
    CodPerfilNaoIncidencia = a_CodPerfilNaoIncidencia
End Property

Public Property Let CodPerfilNaoIncidencia(ByVal Value As String)
    a_CodPerfilNaoIncidencia = Value
End Property

Public Function CarregaAtasNit()
    Dim sql As String
    Dim Condicao As String
    
    sql = "SELECT * FROM TAB_ATA_NIT"
    
    If Atributos.procedimento.Atributos.Contribuinte <> "" Then
        If Condicao = "" Then
            Condicao = " tan_TCI_IM = '" & Atributos.procedimento.Atributos.Contribuinte & "'"
        Else
             Condicao = " AND tan_TCI_IM = '" & Atributos.procedimento.Atributos.Contribuinte & "'"
        End If
    End If
    If Atributos.procedimento.Atributos.NumeroProcedimento <> "" Then
        If Condicao = "" Then
            Condicao = " tan_tpf_numero_procedimento = " & Atributos.procedimento.Atributos.NumeroProcedimento
        Else
             Condicao = " AND tan_tpf_numero_procedimento = " & Atributos.procedimento.Atributos.NumeroProcedimento
        End If
    End If
    If Atributos.procedimento.Atributos.NumeroFiscalizacao <> "" Then
        If Condicao = "" Then
            Condicao = " tan_tfi_numero_fiscalizacao = " & Atributos.procedimento.Atributos.NumeroFiscalizacao
        Else
             Condicao = " AND tan_tfi_numero_fiscalizacao = " & Atributos.procedimento.Atributos.NumeroFiscalizacao
        End If
    End If
    If Condicao <> "" Then sql = sql & " where " & Condicao
End Function

 
Public Function Grava() As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Condicao As String
    
    Grava = False
    If Atributos.procedimento.Grava Then
        Campos = " tan_tpf_numero_procedimento, tan_tfi_numero_fiscalizacao, "
        Campos = Campos & " tan_tci_im, tan_tpn_cod_perfil_nao_incid "
        
        Valores = Bdados.PreparaValor(Atributos.procedimento.Atributos.NumeroProcedimento, _
        Atributos.procedimento.Atributos.NumeroFiscalizacao, Atributos.procedimento.Atributos.Contribuinte, Atributos.CodPerfilNaoIncidencia)
        
        Condicao = " tan_tpf_numero_procedimento = " & Atributos.procedimento.Atributos.NumeroProcedimento
        
        If Bdados.GravaDados("TAB_ATA_NIT", Valores, Campos, Condicao) Then
            Grava = True
        End If
    End If
End Function

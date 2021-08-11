VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eAtarif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public procedimento As New eProcedimentosFiscais
Private a_CodPerfilConstitucional As String
Private a_CumprimentoExigencias As String

Public Property Get CodPerfilConstitucional() As String
    CodPerfilConstitucional = a_CodPerfilConstitucional
End Property

Public Property Let CodPerfilConstitucional(ByVal Value As String)
    a_CodPerfilConstitucional = Value
End Property

Public Property Get CumprimentoExigencias() As String
    CumprimentoExigencias = a_CumprimentoExigencias
End Property

Public Property Let CumprimentoExigencias(ByVal Value As String)
    a_CumprimentoExigencias = Value
End Property

Public Function CarregaAtasRif()
    Dim sql As String
    Dim Condicao As String
    
    sql = "SELECT * FROM TAB_ATA_RIF"
    
    If Atributos.procedimento.Atributos.Contribuinte <> "" Then
        If Condicao = "" Then
            Condicao = " TAR_TCI_IM = '" & Atributos.procedimento.Atributos.Contribuinte & "'"
        Else
             Condicao = " AND TAR_TCI_IM = '" & Atributos.procedimento.Atributos.Contribuinte & "'"
        End If
    End If
    If Atributos.procedimento.Atributos.NumeroProcedimento <> "" Then
        If Condicao = "" Then
            Condicao = " tar_tpf_numero_procedimento = " & Atributos.procedimento.Atributos.NumeroProcedimento
        Else
             Condicao = " AND tar_tpf_numero_procedimento = " & Atributos.procedimento.Atributos.NumeroProcedimento
        End If
    End If
    If Atributos.procedimento.Atributos.NumeroFiscalizacao <> "" Then
        If Condicao = "" Then
            Condicao = " tar_tfi_numero_fiscalizacao = " & Atributos.procedimento.Atributos.NumeroFiscalizacao
        Else
             Condicao = " AND tar_tfi_numero_fiscalizacao = " & Atributos.procedimento.Atributos.NumeroFiscalizacao
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
        Campos = " tar_tpf_numero_procedimento, tar_tfi_numero_fiscalizacao, "
        Campos = Campos & " tar_tci_im, tar_perfil_isencao_fiscal,tar_cumprimento_exigencias"
        
        Valores = Bdados.PreparaValor(Atributos.procedimento.Atributos.NumeroProcedimento, _
        Atributos.procedimento.Atributos.NumeroFiscalizacao, Atributos.procedimento.Atributos.Contribuinte, _
        Atributos.PerfilIsencaoFiscal, Atributos.CumprimentoExigencias)
        
        Condicao = " tar_tpf_numero_procedimento = " & Atributos.procedimento.Atributos.NumeroProcedimento
        
        If Bdados.GravaDados("TAB_ATA_RIF", Valores, Campos, Condicao) Then
            Grava = True
        End If
    End If
End Function

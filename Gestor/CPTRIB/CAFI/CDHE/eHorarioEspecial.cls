VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eHorarioEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_Icod As String
Private m_Horario1 As String
Private m_Horario2 As String
Private m_Horario3 As String
Private m_Motivo As String
Private m_Im As String
Public cadastro As New eCadastroFiscal
Dim Codigo As New ContaCorrente
    
Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property


Public Property Let Icod(ByVal Value As String)
    m_Icod = Value
End Property

Public Property Get Icod() As String
    Icod = m_Icod
End Property

Public Property Let Horario1(ByVal Value As String)
    m_Horario1 = Value
End Property

Public Property Get Horario1() As String
    Horario1 = m_Horario1
End Property

Public Property Let Horario2(ByVal Value As String)
    m_Horario2 = Value
End Property

Public Property Get Horario2() As String
    Horario2 = m_Horario2
End Property

Public Property Let Horario3(ByVal Value As String)
    m_Horario3 = Value
End Property

Public Property Get Horario3() As String
    Horario3 = m_Horario3
End Property

Public Property Let Motivo(ByVal Value As String)
    m_Motivo = Value
End Property

Public Property Get Motivo() As String
    Motivo = m_Motivo
End Property
    
    
Public Function Salvar() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    Dim codCadastro As String
    Set Codigo = New ContaCorrente
    Set cadastro = New eCadastroFiscal
       
    If (m_Icod = "") Then
        codCadastro = Codigo.GeraCodPagamento(44)
        With cadastro
            .icad = codCadastro
            .Im = m_Im
            .Tipo_Cadastro = ecHorarioEspecial
            .Status = ecAtivo
            .Data_Cadastro = Date
        End With
        If Not cadastro.Salvar Then Exit Function
    Else
        codCadastro = m_Icod
    End If
     
        campos = "THE_COD_CADASTRO,THE_HORARIO1, THE_HORARIO2, THE_HORARIO3,THE_MOTIVO"
        valores = Bdados.PreparaValor(codCadastro, m_Horario1, m_Horario2, m_Horario3, m_Motivo)
        condicao = "THE_COD_CADASTRO =" & codCadastro
        Bdados.AbreTrans
        If Bdados.GravaDados("TAB_CAD_HORARIO_ESPECIAL", valores, campos, condicao) Then
        Salvar = True

            Bdados.GravaTrans
        Else
            Bdados.CancelaTrans
        End If
    
End Function

Public Function PreencherGrd(Grid As Object, Im As String) As Boolean
    Dim Sql As String
    Sql = " SELECT THE_COD_CADASTRO AS Código, "
    Sql = Sql & " HORARIO_1 AS Horário_1,"
    Sql = Sql & " HORARIO_2 AS Horário_2,"
    Sql = Sql & " HORARIO_3 AS Horário_3,"
    Sql = Sql & " THE_MOTIVO AS Motivo,"
    Sql = Sql & " THE_HORARIO1,"
    Sql = Sql & " THE_HORARIO2,"
    Sql = Sql & " THE_HORARIO3,"
    Sql = Sql & " TCF_TCI_IM"
    Sql = Sql & " From VIS_CAD_HORARIO_ESPECIAL"
    Sql = Sql & " WHERE TCF_STATUS= " & ecAtivo & " "
    If Im <> "" Then Sql = Sql & "and TCF_TCI_IM = '" & Im & "'"
    
    If Not Grid.Preencher(Bdados, Sql, 1300, 2000, 2000, 2000, 4000, 0, 0, 0, 0) Then
        Util.Avisa "Contribuinte não possui Cadastro de Horário Especial."
    End If
End Function


Public Function PreencherGrdConsulta(Grid As Object, Im As String, cod As String, StatusCad As String) As Boolean
    Dim Sql As String
    Dim condicao As String

    If (cod <> "") Then condicao = condicao & " AND THE_COD_CADASTRO  = " & cod
    If (StatusCad <> "") Then condicao = condicao & " AND TCF_STATUS = " & StatusCad

    Sql = " SELECT THE_COD_CADASTRO AS Código, "
    Sql = Sql & " tci_nome AS Contribuinte,"
    Sql = Sql & " TCF_DATA_CADASTRO AS Data_Cadastro,"
    Sql = Sql & " STATUS AS Status,"
    Sql = Sql & " TCF_DATA_CANCELAMENTO AS Data_Baixa,"
    Sql = Sql & " HORARIO_1 AS Horário_1,"
    Sql = Sql & " HORARIO_2 AS Horário_2,"
    Sql = Sql & " HORARIO_3 AS Horário_3,"
    Sql = Sql & " THE_MOTIVO As Motivo"
    Sql = Sql & " From VIS_CAD_HORARIO_ESPECIAL where 1=1"
    If Im <> "" Then Sql = Sql & " and TCF_TCI_IM = '" & Im & "'"
    Sql = Sql & condicao
    If Not Grid.Preencher(Bdados, Sql, 1100, 5000, 1100, 1200, 1150, 2000, 2000, 2000, 4000) Then
     Util.Avisa "Contribuinte não possui Cadastro de Horário Especial."
    End If
End Function

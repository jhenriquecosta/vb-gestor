VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Util As New VSClass.VSUtil
Private Edita As New VSTexto
Private Relatorio As New CRAXDRT.Report
Attribute Relatorio.VB_VarHelpID = -1
Private Aplica As New CRAXDRT.Application
Private Arq As String
Private Diretorio As String

Private SubRel As String
Public Titulo As String
Public Arvore As Boolean
Public Detalhamento As Boolean
Public Exportacao As Boolean

Public Enum TipoRpt
    Vertical = 0
    Horizontal = 1
End Enum

Public Enum AlinhamentoTexto
    Centralizado = 0
    Esquerdo = 1
    'Direito = 2
    'Justificado = 3
End Enum

Public Enum TipoFormulas
    Normal
    Especial
End Enum

Public Property Let CopiasDetalhes(Num As Long)
    If bRegistrado Then Relatorio.Areas.item("D").CopiesToPrint = Num
End Property

Public Property Get Arquivo() As String
    If bRegistrado Then Arquivo = Arq
End Property

Public Function DefinirArquivo(Bdados As Object, ArquivoRPT As String) As Boolean
    On Error GoTo Trata
    If bRegistrado Then
        Arq = ArquivoRPT
        
        Dim i As Integer
        i = 1
        i = InStr(i, ArquivoRPT, "/")
        Diretorio = ""
        Do
            Diretorio = Mid(ArquivoRPT, 1, i)
            i = InStr(i + 1, ArquivoRPT, "\")
        Loop Until i = 0
        If Dir(Arq) <> "" Then
            Set Relatorio = Aplica.OpenReport(Arq, 1)
            AtualizaBanco Bdados, Relatorio
            AtualizaSubRelatorios Bdados
            LimparFormulas
            DefinirArquivo = True
        Else
'            Util.Informa "Arquivo " & Mid(ArquivoRPT, InStrRev(ArquivoRPT, "\") + 1) & " não encontrado."
            Util.Informa "Arquivo " & Arq & " não encontrado."
            DefinirArquivo = False
        End If
    End If
Trata:
    If Err.Number <> 0 Then
        Util.Erro Err.Number & " - " & Err.Description & "."
    End If
End Function

Public Sub LimparFormulas(Optional Todas As Boolean = False)
    If bRegistrado Then
        Dim i As Integer
        For i = 1 To Relatorio.FormulaFields.Count
            If Todas Or UCase(Mid$(CStr(Relatorio.FormulaFields.item(i).Name), 3, 2)) <> "VT" Then
                Relatorio.FormulaFields.item(i).Text = ""
            End If
        Next
    End If
End Sub

Public Sub Formulas(Formula As String, Valor As String, Optional TipoFormula As TipoFormulas = Normal)
    If bRegistrado Then Relatorio.FormulaFields.GetItemByName(Trim(Formula)).Text = IIf(TipoFormula = Normal, "'" & Edita.TiraPic(Valor, "'") & "'", Valor)
End Sub

Public Sub Parametros(Parametro As String, Valor)
    If bRegistrado Then
        Dim objParametros As CRAXDRT.ParameterFieldDefinitions
        Dim objParametro As CRAXDRT.ParameterFieldDefinition
    
        Set objParametros = Relatorio.ParameterFields
        For Each objParametro In objParametros
            With objParametro
                If .ParameterFieldName = Parametro Then
                    .SetCurrentValue Valor
                    Exit For
                End If
            End With
        Next objParametro
    
        Relatorio.EnableParameterPrompting = False
    End If
End Sub

Public Function Imprimir(Optional Confirmacao As Boolean = False, Optional Copias As Long = 1, Optional Agrupar As Boolean = False, Optional InicioPag As Long = 1, Optional FimPag As Long = 1) As Boolean
    On Error GoTo Trata
    
    If bRegistrado Then
        'Preparar
        Relatorio.PrintOut Confirmacao, Copias, Agrupar, InicioPag, FimPag
        
        Imprimir = True
    End If
Trata:
    If Err.Number <> 0 Then
        Util.Erro Err.Number & " - " & Err.Description & "."
    End If

End Function

Public Function Visualizar() As Boolean
    On Error GoTo Trata
    
    If bRegistrado Then
        If Not VSVisualizador.Visible Then
            Preparar
            VSVisualizador.Rpt.ReportSource = Relatorio
            VSVisualizador.Rpt.DisplayGroupTree = Arvore
            VSVisualizador.Rpt.EnableDrillDown = Detalhamento
            VSVisualizador.Rpt.EnableExportButton = Exportacao
            VSVisualizador.Caption = Titulo
            
            VSVisualizador.Imprimir
            
            Visualizar = True
        Else
            Util.Avisa "Já existe um relatório sendo visualizado."
        End If
    End If
Trata:
    If Err.Number <> 0 Then
        Util.Erro Err.Number & " - " & Err.Description & "."
    End If

End Function

Private Sub Preparar()
    If bRegistrado Then
        Relatorio.DiscardSavedData
        Relatorio.VerifyOnEveryPrint = True
    End If
End Sub

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
    If bRegistrado Then
        Arvore = False
        Detalhamento = True
        Exportacao = True
        Titulo = "Impressão de Relatório"
    End If
End Sub

Public Property Get SubRelatorio() As String
    If bRegistrado Then SubRelatorio = SubRel
End Property

Public Property Let SubRelatorio(ByVal novaSub As String)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Val As Byte
        
        Val = 0
        
        SubRel = novaSub
        If SubRel <> "" Then
            Set Relatorio = Relatorio.OpenSubreport(SubRel)
pos:
            LimparFormulas
        Else
            If Not Relatorio.Parent Is Nothing Then Set Relatorio = Relatorio.Parent
        End If
        
Trata:
        If Err.Number = -2147190528 Then
            If Val = 2 Then
                Set Relatorio = Relatorio.Parent
                Resume
            ElseIf Val = 1 Then
                Val = 2
                Set Relatorio = Relatorio.OpenSubreport(UCase(SubRel))
                Resume pos
            ElseIf Val = 0 Then
                Val = 1
                Set Relatorio = Relatorio.OpenSubreport(LCase(SubRel))
                Resume pos
            End If
        ElseIf Err.Number <> 0 Then
            Util.Erro Err.Number & " - " & Err.Description & "."
        End If
    End If
End Property

Private Sub AtualizaBanco(Bdados As Object, R As CRAXDRT.Report)
    On Error GoTo Trata
    
    If bRegistrado Then
        Dim X As CRAXDRT.DatabaseTable
        For Each X In R.Database.Tables
            X.SetLogOnInfo Bdados.Conexao.Dsn, IIf(Bdados.Conexao.FormatoBanco <> SQLServer, "", Bdados.Conexao.Catalog), Bdados.Conexao.User, Bdados.Conexao.Password
        Next
        Select Case Bdados.Conexao.FormatoBanco
            Case Access
                R.Database.LogOnServerEx "p2soledb.dll", Bdados.Conexao.Dsn, "", "", "", "OLE DB", Bdados.Conexao.ConnectionString
            Case SQLServer, oracle, interbase
                R.Database.LogOnServerEx "p2soledb.dll", Bdados.Conexao.Dsn, Bdados.Conexao.Catalog, Bdados.Conexao.User, Bdados.Conexao.Password, "OLE DB", Bdados.Conexao.ConnectionString
        End Select
    End If
Trata:
    If Err.Number <> 0 Then
        Util.Erro Err.Number & " - " & Err.Description & "."
    End If

End Sub

Private Sub AtualizaSubRelatorios(Bdados As Object)
    On Error GoTo Trata
    
    If bRegistrado Then
        Dim SubRel As CRAXDRT.Report
        Dim Objeto As Object
        Dim Secao As Section
        
        For Each Secao In Relatorio.Sections
            For Each Objeto In Secao.ReportObjects
                Select Case Objeto.Kind
                    Case crSubreportObject
                        Set SubRel = Relatorio.OpenSubreport(Objeto.SubreportName)
                        AtualizaBanco Bdados, SubRel
                End Select
            Next Objeto
        Next Secao
    End If
Trata:
    If Err.Number <> 0 Then
        Util.Erro Err.Number & " - " & Err.Description & "."
    End If

End Sub

Public Sub Cabecalho(Estado As String, Cliente As String, Secretaria As String, Departamento As String, Optional Alinhamento As AlinhamentoTexto = Centralizado)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Val As Byte
        
        Val = 0
        
        
        Val = 0
        If Alinhamento = Esquerdo Then
            Set Relatorio = Relatorio.OpenSubreport("VSCabEsq.rpt")
        Else
            Set Relatorio = Relatorio.OpenSubreport("VSCab.rpt")
        End If
                
pos:
    
        Formulas "Dep_1", "'" & Estado & "'"
        Formulas "Dep_2", "'" & Cliente & "'"
        Formulas "Dep_3", "'" & Secretaria & "'"
        Formulas "Dep_4", "'" & Departamento & "'"
        Set Relatorio = Relatorio.Parent
        
        Exit Sub
Trata:
        If Err.Number = -2147190528 Then
            If Val = 2 Then
                Util.Erro "Rodapé não existente."
            ElseIf Val = 0 Then
                Val = 1
                Set Relatorio = Relatorio.OpenSubreport("VSCAB.RPT")
                Resume pos
            ElseIf Val = 1 Then
                Val = 2
                Set Relatorio = Relatorio.OpenSubreport("vscab.rpt")
                Resume pos
            End If
        ElseIf Err.Number <> 0 Then
            Util.Erro Err.Number & " - " & Err.Description & "."
        End If
    End If
End Sub

Public Sub Rodape(Administrador As String, Cliente As String, Endereco As String, Cod_Relatorio As String, Cod_Usuario As String, Optional Modo As TipoRpt = Vertical)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Val As Byte
        
        Val = 0
        If Modo = Horizontal Then
            Set Relatorio = Relatorio.OpenSubreport("VSRodHor.rpt")
        Else
            Set Relatorio = Relatorio.OpenSubreport("VSRod.rpt")
        End If
        
pos:
        
        Formulas "Administrador", "'" & Administrador & "'"
        Formulas "Cliente", "'" & Cliente & "'"
        Formulas "CodUsuario", "'" & Cod_Usuario & "'"
        Formulas "Endereco", "'" & Endereco & "'"
        Formulas "CodRelatorio", "'" & Cod_Relatorio & "'"
        Set Relatorio = Relatorio.Parent
        
        Exit Sub
Trata:
        If Err.Number = -2147190528 Then
            If Val = 2 Then
                Util.Erro "Rodapé não existente."
            ElseIf Val = 0 Then
                Val = 1
                Set Relatorio = Relatorio.OpenSubreport("VSROD.RPT")
                Resume pos
            ElseIf Val = 1 Then
                Val = 2
                Set Relatorio = Relatorio.OpenSubreport("vsrod.rpt")
                Resume pos
            End If
        ElseIf Err.Number <> 0 Then
            Util.Erro Err.Number & " - " & Err.Description & "."
        End If
    End If
End Sub

Public Property Let Selecao(ByVal TextoSelecao As String)
    If bRegistrado Then
        Relatorio.DiscardSavedData
        Relatorio.RecordSelectionFormula = TextoSelecao
        Relatorio.VerifyOnEveryPrint = True
    End If
End Property

Public Function PreparaSelecaoData(Campo As String, DataInicial As String, Optional DataFinal As String) As String
    PreparaSelecaoData = Campo & IIf(Trim(DataFinal) = "", " = ", " in ")
    PreparaSelecaoData = PreparaSelecaoData & "Date (" & Year(DataInicial) & "," & Month(DataInicial) & "," & Day(DataInicial) & ") "
    If Trim(DataFinal) <> "" Then
        PreparaSelecaoData = PreparaSelecaoData & " to Date (" & Year(DataFinal) & "," & Month(DataFinal) & "," & Day(DataFinal) & ")"
    End If
End Function

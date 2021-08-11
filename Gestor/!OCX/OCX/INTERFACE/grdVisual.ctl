VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl grdVISUAL 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   2535
   ScaleWidth      =   3885
   ToolboxBitmap   =   "grdVisual.ctx":0000
   Begin MSComctlLib.ListView grdGrid 
      Height          =   1905
      Left            =   30
      TabIndex        =   0
      Top             =   285
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   3360
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16119285
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2280
      Width           =   210
   End
   Begin VB.Shape shpBordaExterna 
      BorderColor     =   &H8000000F&
      Height          =   2205
      Left            =   15
      Top             =   0
      Width           =   3885
   End
   Begin VB.Image cmdImprimir 
      Height          =   240
      Left            =   3555
      Picture         =   "grdVisual.ctx":0312
      Top             =   2250
      Width           =   240
   End
   Begin VB.Label lblQtd 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00A8DCDD&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3735
      TabIndex        =   2
      Top             =   30
      Width           =   90
   End
   Begin VB.Label lblRotulo 
      AutoSize        =   -1  'True
      BackColor       =   &H00F5F5F5&
      BackStyle       =   0  'Transparent
      Caption         =   " Resultado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   900
   End
   Begin VB.Shape shpBorda 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   285
      Left            =   15
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "grdVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum enuTipoCampo
    tipTexto = 5
    tipData = 1
    tipInteiro = 2
    tipMoeda = 3
End Enum

Public Enum enuAlinhamentoCampo
    aliEsquerda = 0
    aliCentro = 1
    aliDireita = 2
End Enum
Private Const cteEspacamentoColunas As Integer = 2
Private m_RegistrosPorPagina As Integer
Private m_LarguraPagina As Integer
Public Enum eTipoPapel
    A4 = 0
    Matricial = 1
End Enum
Public Enum eOrientacaoPapel
    Vertical = 1
    Horizontal = 2
End Enum
Private m_TipoPapel As eTipoPapel
Private m_PagInicial As Integer
Private m_PagFinal As Integer
Private m_TamFonte As Integer
Private m_CabecalhoEstado As String
Private m_CabecalhoCliente As String
Private m_CabecalhoSecretaria As String
Private m_CabecalhoDepartamento As String
Private m_CabecalhoTitulo As String
Private m_RodapeUsuario As String
Private m_LarguraCorpo As Integer
Private m_TamQtdRegistros As Integer
Private m_CorFundo As OLE_COLOR
Private m_Caption As String
Private m_CorTitulo As OLE_COLOR
Private m_CorCaption As OLE_COLOR
Private m_CorDica As OLE_COLOR
Private m_OcultarRodape As Boolean
Private m_CheckBox As Boolean

Public Event Click()
Public Event DblClick()
Public Event ItemCheck(ByVal Item As MSComctlLib.ListItem)
Public Event ItemClick(ByVal Item As MSComctlLib.ListItem)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private m_Colunas As clsColunas
Private m_Mensagem As String
Private m_OrientacaoPapel As eOrientacaoPapel
Private m_MarcaUnico As Boolean
Private m_Ordenavel As Boolean

Public Property Get Ordenavel() As Boolean
    Ordenavel = m_Ordenavel
End Property

Public Property Let Ordenavel(ByVal Value As Boolean)
    m_Ordenavel = Value
    PropertyChanged "Ordenavel"
End Property

Public Property Get MarcaUnico() As Boolean
    MarcaUnico = m_MarcaUnico
End Property

Public Property Let MarcaUnico(ByVal Value As Boolean)
    m_MarcaUnico = Value
End Property

Public Property Get MultiSelect() As Boolean
    If bRegistrado Then MultiSelect = grdGrid.MultiSelect
End Property

Public Property Let MultiSelect(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.MultiSelect = Value
End Property

Public Property Get HideSelection() As Boolean
    If bRegistrado Then HideSelection = grdGrid.HideSelection
End Property

Public Property Let HideSelection(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.HideSelection = Value
End Property

Public Property Get HideColumnHeaders() As Boolean
    If bRegistrado Then HideColumnHeaders = grdGrid.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.HideColumnHeaders = Value
End Property

Public Property Get GridLines() As Boolean
    If bRegistrado Then GridLines = grdGrid.GridLines
End Property

Public Property Let GridLines(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.GridLines = Value
End Property

Public Property Get GetFirstVisible() As ListItem
    If bRegistrado Then Set GetFirstVisible = grdGrid.GetFirstVisible
End Property

Public Property Get FullRowSelect() As Boolean
    If bRegistrado Then FullRowSelect = grdGrid.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.FullRowSelect = Value
End Property

Public Property Get FlatScrollBar() As Boolean
    If bRegistrado Then FlatScrollBar = grdGrid.FlatScrollBar
End Property

Public Property Let FlatScrollBar(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.FlatScrollBar = Value
End Property

Public Property Get AllowColumnReorder() As Boolean
    If bRegistrado Then AllowColumnReorder = grdGrid.AllowColumnReorder
End Property

Public Property Let AllowColumnReorder(ByVal Value As Boolean)
    If bRegistrado Then grdGrid.AllowColumnReorder = Value
End Property

Public Property Get OrientacaoPapel() As eOrientacaoPapel
    If bRegistrado Then OrientacaoPapel = m_OrientacaoPapel
End Property

Public Property Let OrientacaoPapel(ByVal Value As eOrientacaoPapel)
    If bRegistrado Then
        m_OrientacaoPapel = Value
        PropertyChanged "OrientacaoPapel"
    End If
End Property

Public Property Get Mensagem() As String
    If bRegistrado Then Mensagem = m_Mensagem
End Property

Public Property Let Mensagem(ByVal Value As String)
    If bRegistrado Then
        m_Mensagem = Value
        lblMensagem = m_Mensagem
        PropertyChanged "Mensagem"
    End If
End Property

Public Function Colunas(Indice As Integer) As clsColuna
    If bRegistrado Then
        If Not m_Colunas(Indice) Is Nothing Then
            Set Colunas = m_Colunas(Indice)
        End If
    End If
End Function

Public Property Get SelectedItem() As ListItem
    If bRegistrado Then Set SelectedItem = grdGrid.SelectedItem
End Property

Public Property Get ListItems() As ListItems
Attribute ListItems.VB_UserMemId = 0
    If bRegistrado Then Set ListItems = grdGrid.ListItems
End Property

Public Property Get CheckBox() As Boolean
    If bRegistrado Then CheckBox = m_CheckBox
End Property

Public Property Let CheckBox(ByVal Value As Boolean)
    If bRegistrado Then
        m_CheckBox = Value
        grdGrid.Checkboxes = m_CheckBox
        PropertyChanged "CheckBox"
    End If
End Property

Public Property Get OcultarRodape() As Boolean
    If bRegistrado Then OcultarRodape = m_OcultarRodape
End Property

Public Property Let OcultarRodape(ByVal Value As Boolean)
    If bRegistrado Then
        m_OcultarRodape = Value
        
        If Value Then
            shpBordaExterna.Height = shpBordaExterna.Height + 280
            grdGrid.Height = grdGrid.Height + 280
            cmdImprimir.Visible = False
        Else
            grdGrid.Height = grdGrid.Height - 280
            shpBordaExterna.Height = shpBordaExterna.Height - 280
            cmdImprimir.Visible = True
        End If
        PropertyChanged "OcultarRodape"
    End If
End Property


Public Property Get CorDica() As OLE_COLOR
    If bRegistrado Then CorDica = m_CorDica
End Property

Public Property Let CorDica(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        m_CorDica = Value
        lblMensagem.ForeColor = Value
        PropertyChanged "CorDica"
    End If
End Property

Public Property Get CorCaption() As OLE_COLOR
    If bRegistrado Then
        CorCaption = m_CorCaption
    End If
End Property

Public Property Let CorCaption(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        m_CorCaption = Value
        lblRotulo.ForeColor = Value
        lblQtd.ForeColor = Value
        PropertyChanged "CorCaption"
    End If
End Property

Public Property Get CorTitulo() As OLE_COLOR
    If bRegistrado Then CorTitulo = m_CorTitulo
End Property

Public Property Let CorTitulo(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        m_CorTitulo = Value
        shpBorda.BackColor = Value
        PropertyChanged "CorTitulo"
    End If
End Property

Public Property Get Caption() As String
    If bRegistrado Then Caption = m_Caption
End Property

Public Property Let Caption(ByVal Value As String)
    If bRegistrado Then
        m_Caption = Value
        lblRotulo = Value
        PropertyChanged "Caption"
    End If
End Property

Public Property Get corFundo() As OLE_COLOR
    If bRegistrado Then corFundo = m_CorFundo
End Property

Public Property Let corFundo(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        m_CorFundo = Value
        BackColor = m_CorFundo
        'cmdImprimir.CorBorda = m_CorFundo
        'cmdImprimir.CorFundo = m_CorFundo
        'cmdImprimir.CorFoco = m_CorFundo
        PropertyChanged "CorFundo"
    End If
End Property

Private Property Get TamQtdRegistros() As Integer
' VTOcx.grdVISUAL.Property TamQtdRegistros
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Informa o tamanho da coluna usada para informar o n° do registro
'
' Ex: TamQtdRegistros = 3, se houver entre 10 e 999 registros
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_TamQtdRegistros = Len(CStr(grdGrid.ListItems.Count))
        TamQtdRegistros = m_TamQtdRegistros
    End If

End Property

Private Property Get LarguraCorpo() As Integer
' VTOcx.grdVISUAL.Property LarguraCorpo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Calcula, considerando os espacos necessarios para o n° do registro
'               e para os espacamentos entre as colunas, o quanto as colunas estao
'               ocupando
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim Coluna As clsColuna, EspacoColunas As Integer
        
        m_LarguraCorpo = 0
    
        If m_Colunas Is Nothing Then Exit Property
        For Each Coluna In m_Colunas
    
            m_LarguraCorpo = m_LarguraCorpo + Coluna.Tamanho
    
        Next
        
        EspacoColunas = m_Colunas.Count * cteEspacamentoColunas
        m_LarguraCorpo = m_LarguraCorpo + TamQtdRegistros + EspacoColunas
        
        LarguraCorpo = m_LarguraCorpo
    End If

End Property

Public Property Get RodapeUsuario() As String

    If bRegistrado Then RodapeUsuario = m_RodapeUsuario

End Property

Public Property Let RodapeUsuario(ByVal Value As String)
' VTOcx.grdVISUAL.Property RodapeUsuario
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Nome do usuario que aparecera no rodape da impressao
'
' Parametros : Value as String - Codigo do usuario
'
' Ex: .RodapeUsuario = "SERGIO"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_RodapeUsuario = Value
        PropertyChanged "RodapeUsuario"
    End If

End Property

Public Property Get CabecalhoTitulo() As String

    If bRegistrado Then CabecalhoTitulo = m_CabecalhoTitulo

End Property

Public Property Let CabecalhoTitulo(ByVal Value As String)
' VTOcx.grdVISUAL.Property CabecalhoTitulo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Titulo do relatorio
'
' Parametros : Value as String - Titulo
'
' Ex: CabecalhoTitulo = "Relacao dos meus fornecedores"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_CabecalhoTitulo = Value
        PropertyChanged "CabecalhoTitulo"
    End If

End Property

Public Property Get CabecalhoDepartamento() As String

    If bRegistrado Then CabecalhoDepartamento = m_CabecalhoDepartamento

End Property

Public Property Let CabecalhoDepartamento(ByVal Value As String)
' VTOcx.grdVISUAL.Property CabecalhoDepartamento
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Tendo como base o VSCab.rpt, CabecalhoDepartamento é o nome do Departamento que está
'               emitindo o relatorio
'
' Parametros : Value as String - Departamento
'
' Ex: CabecalhoDepartamento = "Departamento de Estatisticas"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_CabecalhoDepartamento = Value
        PropertyChanged "CabecalhoDepartamento"
    End If

End Property

Public Property Get CabecalhoSecretaria() As String

    If bRegistrado Then CabecalhoSecretaria = m_CabecalhoSecretaria

End Property

Public Property Let CabecalhoSecretaria(ByVal Value As String)
' VTOcx.grdVISUAL.Property CabecalhoSecretaria
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Tendo como base o VSCab.rpt, CabecalhoSecretaria é o nome da secretaria que está
'               emitindo o relatorio
'
' Parametros : Value as String - CabecalhoSecretaria
'
' Ex: CabecalhoSecretaria = "Secretaria Municipal de Relatorios"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_CabecalhoSecretaria = Value
        PropertyChanged "CabecalhoSecretaria"
    End If

End Property

Public Property Get CabecalhoCliente() As String

    If bRegistrado Then CabecalhoCliente = m_CabecalhoCliente

End Property

Public Property Let CabecalhoCliente(ByVal Value As String)
' VTOcx.grdVISUAL.Property CabecalhoCliente
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Nome do cliente
'
' Parametros : Value (String) - CabecalhoCliente
'
' Ex: CabecalhoCliente = "Prefeitura Municipal de Sao Luis"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_CabecalhoCliente = Value
        PropertyChanged "CabecalhoCliente"
    End If

End Property

Public Property Get CabecalhoEstado() As String

    If bRegistrado Then CabecalhoEstado = m_CabecalhoEstado

End Property

Public Property Let CabecalhoEstado(ByVal Value As String)
' VTOcx.grdVISUAL.Property CabecalhoEstado
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Uf do cliente
'
' Parametros : Value (String) - CabecalhoEstado
'
' Ex: CabecalhoEstado = "MA"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_CabecalhoEstado = Value
        PropertyChanged "CabecalhoEstado"
    End If

End Property

Public Property Get TamFonte() As Integer

    If bRegistrado Then TamFonte = m_TamFonte

End Property

Public Property Let TamFonte(ByVal Value As Integer)
' VTOcx.grdVISUAL.Property TamFonte
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Tamanho da fonte Courier New que sera usado na impressao
'
' Parametros : Value (Integer)
'
' Ex: TamFonte = 8
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_TamFonte = Value
        PropertyChanged "TamFonte"
    End If

End Property

Public Property Get TotalPag() As Integer
' VTOcx.grdVISUAL.Property TotalPag
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Calcula quantas paginas o relatorio vai usar
'
' Ex: TotalPag = 20, se Papel = A4Horizontal e QtdRegistros = 709
'--------------------------------------------------------------------------------

    If bRegistrado Then TotalPag = CalcularTotalPaginas(grdGrid.ListItems.Count, RegistrosPorPagina)

End Property

Public Property Get PagFinal() As Integer

    If bRegistrado Then PagFinal = m_PagFinal

End Property

Public Property Let PagFinal(ByVal Value As Integer)
' VTOcx.grdVISUAL.Property PagFinal
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:23
'
' Descricao  : Pagina onde terminará a impressao
'
' Parametros : Value (Integer)
'
' Ex: PagFinal = 5, o relatório será impresso até a pag. 5
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_PagFinal = Value
        PropertyChanged "PagFinal"
    End If

End Property

Public Property Get PagInicial() As Integer

    If bRegistrado Then PagInicial = m_PagInicial

End Property

Public Property Let PagInicial(ByVal Value As Integer)
' VTOcx.grdVISUAL.Property PagInicial
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Pagina onde inicia a impressao
'
' Parametros : Value (Integer)
'
' Ex: PagInicial = 10, o relatorio sera impresso a partir da pag. 10
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_PagInicial = Value
        PropertyChanged "PagInicial"
    End If

End Property

Public Property Get TipoPapel() As eTipoPapel

    If bRegistrado Then TipoPapel = m_TipoPapel

End Property

Public Property Let TipoPapel(ByVal Value As eTipoPapel)
' VTOcx.grdVISUAL.Property TipoPapel
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Tipo do papel usado na impressora
'
' Parametros : Value (eTipoPapel)
'
' Ex: Papel = A4Vertical
'--------------------------------------------------------------------------------

    If bRegistrado Then m_TipoPapel = Value

End Property

Private Property Get RegistrosPorPagina() As Integer
' VTOcx.grdVISUAL.Property RegistrosPorPagina
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Calcula quantos registros a pagina comporta. Depende do tipo de papel
'               escolhido
'
' Ex: RegistrosPorPagina = 37, se TipoPapel = A4Horizontal
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Select Case TipoPapel
            Case A4
                Select Case OrientacaoPapel
                    Case eOrientacaoPapel.Vertical
                        Select Case TamFonte
                            Case 6: m_RegistrosPorPagina = 114 - 13 '13 = 9 linh. cab. + 4 linh. rod.
                            Case 8: m_RegistrosPorPagina = 91 - 13
                            Case 10: m_RegistrosPorPagina = 73 - 13
                            Case Else: m_RegistrosPorPagina = 60 - 13
                        End Select
                    Case eOrientacaoPapel.Horizontal
                        Select Case TamFonte
                            Case 6: m_RegistrosPorPagina = 82 - 13
                            Case 8: m_RegistrosPorPagina = 66 - 13
                            Case 10: m_RegistrosPorPagina = 53 - 13
                            Case Else: m_RegistrosPorPagina = 40 - 13
                        End Select
                End Select
            Case Matricial: m_RegistrosPorPagina = 46
        End Select
    
        RegistrosPorPagina = m_RegistrosPorPagina
    End If
End Property

Private Property Get LarguraPagina() As Integer
' VTOcx.grdVISUAL.Property LarguraPagina
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Calcula a quantidade de colunas que a pagina comporta
'
' Ex: LarguraPagina = 129, se TipoPapel = A4Horizontal
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Select Case TipoPapel
            Case A4
                Select Case OrientacaoPapel
                    Case eOrientacaoPapel.Vertical
                        Select Case TamFonte
                            Case 6: m_LarguraPagina = 160
                            Case 8: m_LarguraPagina = 120
                            Case 10: m_LarguraPagina = 96
                            Case Else: m_LarguraPagina = 70
                        End Select
                    Case eOrientacaoPapel.Horizontal
                        Select Case TamFonte
                            Case 6: m_LarguraPagina = 221
                            Case 8: m_LarguraPagina = 165
                            Case 10: m_LarguraPagina = 132
                            Case Else: m_LarguraPagina = 100
                        End Select
                End Select
    
            Case Matricial: m_LarguraPagina = 158
        End Select
    
        LarguraPagina = m_LarguraPagina
    End If
End Property

Private Sub cmdImprimir_Click()
' VTOcx.grdVISUAL.Sub cmdImprimir_Click
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Envia o conteudo do grid para a impressora
'
'--------------------------------------------------------------------------------

    On Error GoTo Trata
    If bRegistrado Then
        Dim intPagina As Integer
        Dim intRegistroInicial As Integer, intRegistroFinal As Integer
        Dim i As Integer, j As Integer
        Dim PrimeiroReg As Integer
        Dim Util As New VSUtil
        
        If grdGrid.ListItems.Count = 0 Then
    
            lblMensagem = Nvl(Mensagem, "Não há informação para imprimir.")
    
        Else
            frmImprimir.Tag = CabecalhoTitulo & "||" & TipoPapel & "||" & PagInicial & "||" & TotalPag & "||" & TamFonte & "||" & OrientacaoPapel
            frmImprimir.Show vbModal
            If OpcoesImpressao = "" Then Exit Sub
            CabecalhoTitulo = Util.ParseString(OpcoesImpressao, "||", 1)
            TipoPapel = Util.ParseString(OpcoesImpressao, "||", 2)
            PagInicial = Util.ParseString(OpcoesImpressao, "||", 3)
            PagFinal = Util.ParseString(OpcoesImpressao, "||", 4)
            TamFonte = Util.ParseString(OpcoesImpressao, "||", 5)
            OrientacaoPapel = Util.ParseString(OpcoesImpressao, "||", 6) + 1
            
            lblMensagem = Nvl(Mensagem, "Enviando informações para impressora...")
            Screen.MousePointer = 11
            intRegistroInicial = 0
    
            PrimeiroReg = 0
    
            Printer.Orientation = OrientacaoPapel
    
            For intPagina = PagInicial To PagFinal
    
                intRegistroInicial = (RegistrosPorPagina * (intPagina - 1)) + 1
    
                If PrimeiroReg = 0 Then PrimeiroReg = intRegistroInicial
    
                intRegistroFinal = intRegistroInicial + RegistrosPorPagina - 1
                ImprimirCabecalho
                ImprimirCorpo intRegistroInicial, intRegistroFinal
                If Mensagem <> "" And intPagina = PagFinal Then
                    Printer.Print
                    Printer.Print Mensagem
                    intRegistroFinal = intRegistroFinal + 2
                End If

                i = 0
                j = RegistrosPorPagina * intPagina
    
                For i = intRegistroFinal + 1 To j
    
                    Printer.Print ""
    
                Next
    
                ImprimirRodape intPagina
                Printer.NewPage
    
            Next
    
            Printer.EndDoc
            Screen.MousePointer = 0
            lblMensagem = Nvl(Mensagem, "Fim da impressão.")
    
        End If
    End If
    Exit Sub
    
Trata:
    Screen.MousePointer = 0
    lblMensagem = Nvl(Mensagem, Err.Description)
    Printer.KillDoc

End Sub

Private Sub cmdImprimir_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
' VTOcx.grdVISUAL.Sub cmdImprimir_MouseEnter
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Exibe mensagem da funcionalidade do botao
'
' Parametros : Button (Integer)
'              Shift (Integer)
'              X (Single)
'              Y (Single)
'
'--------------------------------------------------------------------------------

    If bRegistrado Then lblMensagem = Nvl(Mensagem, "Imprimir o conteúdo")

End Sub

Private Sub cmdImprimir_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
' VTOcx.grdVISUAL.Sub cmdImprimir_MouseExit
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Apaga a dica do botao
'
' Parametros : Button (Integer)
'              Shift (Integer)
'              X (Single)
'              Y (Single)
'
'--------------------------------------------------------------------------------

    If bRegistrado Then lblMensagem = Nvl(Mensagem, "")

End Sub

Private Sub grdGrid_Click()
    If bRegistrado Then RaiseEvent Click
End Sub

Private Sub grdGrid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' VTOcx.grdVISUAL.Sub grdGrid_ColumnClick
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Ordena o grid pela coluna clicada
'
' Parametros : ColumnHeader (MSComctlLib.ColumnHeader) - Coluna clicada
'
'--------------------------------------------------------------------------------

    'toggle the sort order for use in the CompareXX routines

    If bRegistrado Then
        If Ordenavel Then
            If grdGrid.ListItems.Count = 0 Then Exit Sub
            If m_Colunas Is Nothing Then Exit Sub
            If grdGrid.SortKey <> ColumnHeader.Index - 1 Then
        
                sOrder = True
        
            Else
        
                sOrder = Not sOrder
        
            End If
           
            grdGrid.SortKey = ColumnHeader.Index - 1
            lngSubItem = grdGrid.SortKey
           
            Select Case m_Colunas.Item(ColumnHeader.Index).Tipo
        
                Case tipTexto
                    grdGrid.Sorted = True
                    'Use sort routine to sort by String
                    'grdGrid.Sorted = False
                    'SendMessage grdGrid.hwnd, _
                       LVM_SORTITEMS, _
                       grdGrid.hwnd, _
                       ByVal FARPROC(AddressOf CompareText)
              
                Case tipData
                    'Use sort routine to sort by date
                    grdGrid.Sorted = False
                    SendMessage grdGrid.hWnd, _
                       LVM_SORTITEMS, _
                       grdGrid.hWnd, _
                       ByVal FARPROC(AddressOf CompareDates)
        
                Case tipInteiro, tipMoeda
                    'Use sort routine to sort by value
                    grdGrid.Sorted = False
                    SendMessage grdGrid.hWnd, _
                       LVM_SORTITEMS, _
                       grdGrid.hWnd, _
                       ByVal FARPROC(AddressOf CompareValues)
                       
            End Select
        End If
    End If
End Sub

Private Sub grdGrid_DblClick()
    If bRegistrado Then RaiseEvent DblClick
End Sub

Private Sub grdGrid_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If bRegistrado Then
        If m_MarcaUnico Then
            Dim Indice As Double
            Dim ValorIndice As Boolean
            
            Indice = Item.Index
            ValorIndice = Item.Checked
            
            MarcarTodos False
            grdGrid.ListItems(Indice).Selected = True
            Item.Checked = ValorIndice
        End If
        RaiseEvent ItemCheck(Item)
    End If
End Sub

Private Sub grdGrid_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If bRegistrado Then
        RaiseEvent ItemClick(Item)
    End If
End Sub

Private Sub grdGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If bRegistrado Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub grdGrid_KeyPress(KeyAscii As Integer)
    If bRegistrado Then RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub grdGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    If bRegistrado Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub grdGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bRegistrado Then RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub grdGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bRegistrado Then RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub grdGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bRegistrado Then RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Initialize()
' VTOcx.grdVISUAL.Sub UserControl_Initialize
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Extrai as figuras dos botoes
'
'--------------------------------------------------------------------------------

    ValidaComponente "INTERFACE"
    m_TamQtdRegistros = 0
'    If bRegistrado Then
        Set Util = New VSUtil
'    End If

End Sub

Private Sub UserControl_Paint()
' VTOcx.grdVISUAL.Sub UserControl_Paint
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Torna o cabecalho do grid flat
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        If Tag = "" Then
            FlatColumn UserControl.ContainerHwnd, grdGrid
            Tag = "flat"
        End If
    End If
End Sub

Private Sub UserControl_Resize()
' VTOcx.grdVISUAL.Sub UserControl_Resize
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Redesenha os objetos que compoem o controle
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        shpBordaExterna.Width = Width - 30
        shpBorda.Width = shpBordaExterna.Width
        lblQtd.Left = shpBorda.Width - 150
        shpBordaExterna.Height = IIf(Height - cmdImprimir.Height - 50 > 0, Height - cmdImprimir.Height - 50, 100)
        grdGrid.Width = shpBordaExterna.Width - 70
        grdGrid.Height = IIf(shpBordaExterna.Height - shpBorda.Height - 30 > 0, shpBordaExterna.Height - shpBorda.Height - 30, 100)
        cmdImprimir.Top = grdGrid.Top + grdGrid.Height + 60
        lblMensagem.Top = cmdImprimir.Top
        cmdImprimir.Left = grdGrid.Width - cmdImprimir.Width + 100
'        PrbBarra.Left = grdGrid.Left
'        PrbBarra.Top = cmdImprimir.Top
'        PrbBarra.Width = grdGrid.Width - cmdImprimir.Width
    End If
End Sub

Public Property Get CorBorda() As OLE_COLOR

    If bRegistrado Then CorBorda = shpBorda.BorderColor

End Property

Public Property Let CorBorda(ByVal vnewvalue As OLE_COLOR)
' VTOcx.grdVISUAL.Property CorBorda
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Define a cor da borda do controle
'
' Parametros : vnewvalue (OLE_COLOR)
'
' Ex: CorBorda = vbBlue
'--------------------------------------------------------------------------------

    If bRegistrado Then
        shpBorda.BorderColor = vnewvalue
        shpBordaExterna.BorderColor = vnewvalue
    End If
End Property

Private Sub UserControl_InitProperties()
' VTOcx.grdVISUAL.Sub UserControl_InitProperties
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Inicializa os valores padroes das propriedades
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        CorBorda = UserControl.Ambient.ForeColor
        TipoPapel = A4
        m_PagInicial = 1
        m_PagFinal = 1
        m_TamFonte = 8
        m_CabecalhoEstado = "Estado"
        m_CabecalhoCliente = "Prefeitura"
        m_CabecalhoSecretaria = "Secretaria"
        m_CabecalhoDepartamento = "Depto"
        m_CabecalhoTitulo = "Titulo"
        m_RodapeUsuario = "Usuario"
        m_LarguraCorpo = 0
    
        corFundo = UserControl.Ambient.BackColor
        
        Caption = "Resultado"
        CorTitulo = UserControl.Ambient.BackColor
        CorCaption = UserControl.Ambient.ForeColor
        CorDica = UserControl.Ambient.ForeColor
        CheckBox = False
        MarcaUnico = False
        Enabled = True
        Mensagem = ""
        OrientacaoPapel = Vertical
        Ordenavel = True
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' VTOcx.grdVISUAL.Sub UserControl_ReadProperties
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Guarda o valor das propriedades nas variaveis locais
'
' Parametros : PropBag (PropertyBag) - Propriedades do controle
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        CorBorda = PropBag.ReadProperty("CorBorda", UserControl.Ambient.ForeColor)
        TipoPapel = PropBag.ReadProperty("TipoPapel", A4)
        PagInicial = PropBag.ReadProperty("PagInicial", 1)
        PagFinal = PropBag.ReadProperty("PagFinal", 1)
        TamFonte = PropBag.ReadProperty("TamFonte", 8)
        CabecalhoEstado = PropBag.ReadProperty("CabecalhoEstado", "Estado")
        CabecalhoCliente = PropBag.ReadProperty("CabecalhoCliente", "Prefeitura")
        CabecalhoSecretaria = PropBag.ReadProperty("CabecalhoSecretaria", "Secretaria")
        CabecalhoDepartamento = PropBag.ReadProperty("CabecalhoDepartamento", "Depto")
        CabecalhoTitulo = PropBag.ReadProperty("CabecalhoTitulo", "Titulo")
        RodapeUsuario = PropBag.ReadProperty("RodapeUsuario", "Usuario")
        m_LarguraCorpo = PropBag.ReadProperty("LarguraCorpo", 0)
        corFundo = PropBag.ReadProperty("CorFundo", UserControl.Ambient.BackColor)
        Caption = PropBag.ReadProperty("Caption", "Resultado")
        CorTitulo = PropBag.ReadProperty("CorTitulo", UserControl.Ambient.BackColor)
        CorCaption = PropBag.ReadProperty("CorCaption", UserControl.Ambient.ForeColor)
        CorDica = PropBag.ReadProperty("CorDica", UserControl.Ambient.ForeColor)
'        OcultarCabecalho = PropBag.ReadProperty("OcultarCabecalho", False)
        OcultarRodape = PropBag.ReadProperty("OcultarRodape", False)
        CheckBox = PropBag.ReadProperty("CheckBox", False)
        Enabled = PropBag.ReadProperty("Enabled", True)
        Mensagem = PropBag.ReadProperty("Mensagem", "")
        OrientacaoPapel = PropBag.ReadProperty("OrientacaoPapel", eOrientacaoPapel.Vertical)
        MarcaUnico = PropBag.ReadProperty("MarcaUnico", False)
        Ordenavel = PropBag.ReadProperty("Ordenavel", True)
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' VTOcx.grdVISUAL.Sub UserControl_WriteProperties
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Seta as propriedades com o valor das variaveis
'
' Parametros : PropBag (PropertyBag) - Propriedades do controle
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Call PropBag.WriteProperty("CorBorda", shpBorda.BorderColor, UserControl.Ambient.ForeColor)
        Call PropBag.WriteProperty("TipoPapel", m_TipoPapel, A4)
        Call PropBag.WriteProperty("PagInicial", m_PagInicial, 1)
        Call PropBag.WriteProperty("PagFinal", m_PagFinal, 1)
        Call PropBag.WriteProperty("TamFonte", m_TamFonte, 8)
        Call PropBag.WriteProperty("CabecalhoEstado", m_CabecalhoEstado, "Estado")
        Call PropBag.WriteProperty("CabecalhoCliente", m_CabecalhoCliente, "Prefeitura")
        Call PropBag.WriteProperty("CabecalhoSecretaria", m_CabecalhoSecretaria, "Secretaria")
        Call PropBag.WriteProperty("CabecalhoDepartamento", m_CabecalhoDepartamento, "Depto")
        Call PropBag.WriteProperty("CabecalhoTitulo", m_CabecalhoTitulo, "Titulo")
        Call PropBag.WriteProperty("RodapeUsuario", m_RodapeUsuario, "Usuario")
        Call PropBag.WriteProperty("LarguraCorpo", m_LarguraCorpo, 0)
    
        Call PropBag.WriteProperty("CorFundo", m_CorFundo, UserControl.Ambient.BackColor)
        Call PropBag.WriteProperty("Caption", m_Caption, "Resultado")
        Call PropBag.WriteProperty("CorTitulo", m_CorTitulo, UserControl.Ambient.BackColor)
        Call PropBag.WriteProperty("CorCaption", m_CorCaption, UserControl.Ambient.ForeColor)
        Call PropBag.WriteProperty("CorDica", m_CorDica, UserControl.Ambient.ForeColor)
        Call PropBag.WriteProperty("OcultarRodape", m_OcultarRodape, False)
        Call PropBag.WriteProperty("CheckBox", m_CheckBox, False)
        Call PropBag.WriteProperty("Enabled", grdGrid.Enabled, True)
        Call PropBag.WriteProperty("Mensagem", m_Mensagem, "")
        Call PropBag.WriteProperty("OrientacaoPapel", m_OrientacaoPapel, eOrientacaoPapel.Vertical)
        Call PropBag.WriteProperty("MarcaUnico", m_MarcaUnico, False)
        Call PropBag.WriteProperty("Ordenavel", m_Ordenavel, True)
    End If
End Sub

Public Function Preencher(BDados As Object, Sql As String, ParamArray Tamanho_Colunas()) As Boolean
  ' VTOcx.grdVISUAL.Function Preencher
  '================================================================================
  ' Queiroz em VTDES_01
  ' 01/06/2002-14:03:22
  '
  ' Descricao  : Preenche o grid com o resultado de um comando sql
  '
  ' Parametros : BDados (Object) - Objeto do tipo VSDados. Conexao com o banco
  '              Sql (String) - Comando sql
  '              Tamanho_Colunas() (Variant) - Tamanho individual de cada coluna
  '
  ' Ex: Preencher(BDados, "SELECT * FROM Tabela",1000,2000)
  '       Preencher (BDados, "SELECT * FROM Tabela") - As colunas terao o tamanho do maior valor
  '--------------------------------------------------------------------------------
    On Error GoTo Trata

    Dim RS As VSRecordset
    Dim i As Integer
    Dim ItmX As Object

    
    Set m_Colunas = New clsColunas
  
    lblQtd = "0"
    lblQtd.Visible = False
    lblMensagem = ""
    If Trim$(Sql) = "" Then
        Call SendMessage(grdGrid.hWnd, LVM_DELETEALLITEMS, 0, ByVal 0&)
    Else
        Call SendMessage(grdGrid.hWnd, LVM_DELETEALLITEMS, 0, ByVal 0&)
        If BDados.AbreTabela(Sql, RS) Then
            preencherCabecalho RS, Tamanho_Colunas
            preencherCorpo RS
            
            lblQtd.Visible = True
            Preencher = True
        End If
        BDados.FechaTabela RS
    End If

    For i = 1 To m_Colunas.Count
        If m_Colunas(i).Tipo = tipInteiro Or m_Colunas(i).Tipo = tipMoeda Then m_Colunas(i).Media = m_Colunas(i).Soma / lblQtd
        If m_Colunas(i).Tipo = tipInteiro And i > 1 Then grdGrid.ColumnHeaders(m_Colunas(i).Nome).Alignment = lvwColumnRight
        If m_Colunas(i).Tipo = tipMoeda And i > 1 Then
            grdGrid.ColumnHeaders(m_Colunas(i).Nome).Alignment = lvwColumnRight
            For Each ItmX In grdGrid.ListItems
                ItmX.SubItems(i - 1) = Format$(ItmX.SubItems(i - 1), "#,##0.00")
                m_Colunas(i).Tamanho = Len(ItmX.SubItems(i - 1))
            Next
        End If
    Next
    If LBound(Tamanho_Colunas) > UBound(Tamanho_Colunas) Then DimensionarColunas
    Exit Function
    
Trata:
    lblMensagem = "Não há registros para a consulta."
End Function

Private Sub preencherCabecalho(RS As VSRecordset, ParamArray Tamanho_Colunas())
    Dim i As Integer
    Dim Coluna As clsColuna
    
    ReDim TipoColunas(1 To RS.Fields.Count) As enuTipoCampo 'ANDRE(03/08/2002)
    
    grdGrid.Arrange = 2 'lvwAutoTop
    grdGrid.GridLines = True
    grdGrid.LabelEdit = 1 'lvwManual
    grdGrid.View = 3 'lvwReport
    grdGrid.FullRowSelect = True
    grdGrid.HotTracking = True
    grdGrid.FlatScrollBar = False
    grdGrid.HideSelection = False
    grdGrid.LabelWrap = True
    grdGrid.ListItems.Clear
    grdGrid.ColumnHeaders.Clear
    
    For i = 0 To RS.Fields.Count - 1
        Set Coluna = New clsColuna
        Coluna.Nome = RS.Fields(i).Name
        Coluna.Tamanho = Len(Coluna.Nome)
        m_Colunas.Add Coluna
        If i <= UBound(Tamanho_Colunas()(0)) And UBound(Tamanho_Colunas()(0)) > 0 Then
            If Tamanho_Colunas()(0)(i) <= 100 Then
                grdGrid.ColumnHeaders.Add , Coluna.Nome, Coluna.Nome, (Tamanho_Colunas()(0)(i) / 100) * grdGrid.Width
            Else
                grdGrid.ColumnHeaders.Add , Coluna.Nome, Coluna.Nome, Tamanho_Colunas()(0)(i)
            End If
        Else
            grdGrid.ColumnHeaders.Add , Coluna.Nome, Coluna.Nome, (grdGrid.Width / RS.Fields.Count)
        End If
        Coluna.Width = grdGrid.ColumnHeaders(i + 1).Width
    Next
End Sub

Private Function PegaMax(RS As VSRecordset) As Integer
    Dim i As Integer
    RS.MoveFirst
    Do Until RS.EOF
        i = i + 1
        RS.MoveNext
    Loop
    PegaMax = i
    RS.MoveFirst
End Function

Private Sub preencherCorpo(RS As VSRecordset)
    Dim i As Integer
    Dim Valor As String
    Dim ItmX As Object
    Dim ValorBarra As Integer
'    PrbBarra.Visible = True
    
'    PrbBarra.Max = PegaMax(RS)
'    PrbBarra.FloodShowPct = True
    Do Until RS.EOF
        ValorBarra = ValorBarra + 1
'        PrbBarra.FloodPercent = ValorBarra
'        DoEvents
        'Coluna 0
        Valor = "" & RS(0)
        Set ItmX = grdGrid.ListItems.Add(, , Valor)
        m_Colunas.Item(1).Tamanho = Len(Nvl(Valor, 0))
        lblQtd = lblQtd + 1
        If IsDate(RS(0)) Then
            m_Colunas.Item(1).Tipo = tipData
        ElseIf IsNumeric(RS(0)) And m_Colunas.Item(1).Tipo <> tipTexto Then
            i = InStr(1, RS(0), ",")
            m_Colunas.Item(1).Min = RS(0)
            m_Colunas.Item(1).Max = RS(0)
            m_Colunas.Item(1).Soma = RS(0)
            m_Colunas.Item(1).Tipo = IIf(i > 0, tipMoeda, tipInteiro)
        Else
            m_Colunas.Item(1).Tipo = tipTexto
        End If
        'Colunas 1..x
        For i = 1 To RS.Fields.Count - 1
            Valor = "" & RS(i)
            ItmX.SubItems(i) = Valor
            m_Colunas.Item(i + 1).Tamanho = Len(Valor)
            If IsDate(CStr("" & RS(i))) And InStr(1, CStr("" & RS(i)), "/") > 0 Then
                m_Colunas.Item(i + 1).Tipo = tipData
            ElseIf IsNumeric(CStr("" & RS(i))) And m_Colunas.Item(i + 1).Tipo <> tipTexto Then
                m_Colunas.Item(i + 1).Min = RS(i)
                m_Colunas.Item(i + 1).Max = RS(i)
                m_Colunas.Item(i + 1).Soma = RS(i)
                m_Colunas.Item(i + 1).Tipo = IIf((InStr(1, RS(i), ",") > 0), tipMoeda, tipInteiro)
            Else
                m_Colunas.Item(i + 1).Tipo = tipTexto
            End If
        Next
        RS.MoveNext
    Loop
'    PrbBarra.Visible = False
End Sub
Private Sub DimensionarColunas()
' VTOcx.grdVISUAL.Sub DimensionarColunas
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Aplica o tamanho do maior valor em cada coluna
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim col2adjust As Long, j As Long
        j = grdGrid.ColumnHeaders.Count - 1
    
        For col2adjust = 0 To j
       
            Call SendMessage(grdGrid.hWnd, _
               LVM_SETCOLUMNWIDTH, _
               col2adjust, _
               ByVal LVSCW_AUTOSIZE_USEHEADER)
    
        Next
    End If

End Sub

Private Sub ImprimirCabecalho()
' VTOcx.grdVISUAL.Sub ImprimirCabecalho
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Imprime o cabecalho do relatorio
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = TamFonte
        
        '1 - Estado
        Printer.Print AlinharCampo(CabecalhoEstado, aliCentro, LarguraPagina)
        
        '2 - Prefeitura
        Printer.Print AlinharCampo(CabecalhoCliente, aliCentro, LarguraPagina)
        
        '3 - Secretaria
        Printer.Print AlinharCampo(CabecalhoSecretaria, aliCentro, LarguraPagina)
        
        '4 - Departamento
        Printer.Print AlinharCampo(CabecalhoDepartamento, aliCentro, LarguraPagina)
    
        '5 - Branco
        Printer.Print
        '6 - Titulo
        Printer.Print AlinharCampo(CabecalhoTitulo, aliCentro, LarguraPagina)
        '7 - Branco
        Printer.Print
        
        Dim Coluna As clsColuna
        Dim strColunas As String, strLinha As String
        
        If Not m_Colunas Is Nothing Then
            For Each Coluna In m_Colunas
        
                strColunas = strColunas & IIf(Len(strColunas) = 0, Space$(TamQtdRegistros), "")
                strLinha = strLinha & IIf(Len(strLinha) = 0, Space$(TamQtdRegistros), "")
                
                If Coluna.Width > 0 Then
                    strColunas = strColunas & Space$(cteEspacamentoColunas) & PreencherCampo(UCase$(Coluna.Nome), Coluna.Tamanho, Coluna.Tipo)
                    strLinha = strLinha & Space$(cteEspacamentoColunas) & String$(Coluna.Tamanho, "-")
                End If
                 
        
                '<Removed by: Project Administrator at: 25/07/2002-17:59:04 on machine: VTDES01>
                '        strColunas = strColunas & IIf(Len(strColunas) = 0, Space$(TamQtdRegistros), "") & Space$(cteEspacamentoColunas) & PreencherCampo(UCase$(Coluna.Nome), Coluna.Tamanho, Coluna.Tipo)
                '        strLinha = strLinha & IIf(Len(strLinha) = 0, Space$(TamQtdRegistros), "") & Space$(cteEspacamentoColunas) & String$(Coluna.Tamanho, "-")
                '</Removed by: Project Administrator at: 25/07/2002-17:59:04 on machine: VTDES01>
            Next
        End If
    
        '8 - Nome das colunas
        Printer.Print strColunas
        
        '9 - Linhas
        Printer.Print strLinha
    End If
End Sub

Private Sub ImprimirCorpo(LinhaInicial As Integer, LinhaFinal As Integer)
' VTOcx.grdVISUAL.Sub ImprimirCorpo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Imprime o corpo do cabecalho
'
' Parametros : LinhaInicial (Integer) - Registro onde tera inicio
'              LinhaFinal (Integer) - Registro final  da impressao
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        On Error GoTo Trata
        Dim intLinha As Integer, intColuna As Integer
        Dim strLinha As String
        
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = TamFonte
        
        If LinhaFinal > grdGrid.ListItems.Count Then LinhaFinal = grdGrid.ListItems.Count
    
        For intLinha = LinhaInicial To LinhaFinal
    
            '1 - Numero de ordem da linha
            strLinha = PreencherCampo(CStr(intLinha), TamQtdRegistros, tipInteiro)
            
            '2 - Primeiro campo
            If grdGrid.ColumnHeaders(1).Width > 0 Then
                If Not Colunas(1) Is Nothing Then
                    strLinha = strLinha & Space$(cteEspacamentoColunas) & PreencherCampo(grdGrid.ListItems(intLinha), Colunas(1).Tamanho, Colunas(1).Tipo)
                End If
            End If
            
            '3 - Demais campos
    
            If Not m_Colunas Is Nothing Then
                For intColuna = 1 To m_Colunas.Count - 1
        
                    If grdGrid.ColumnHeaders(intColuna + 1).Width > 0 Then
                        strLinha = strLinha & Space$(cteEspacamentoColunas) & _
                           PreencherCampo(grdGrid.ListItems(intLinha).SubItems(intColuna), _
                           Colunas(intColuna + 1).Tamanho, _
                           Colunas(intColuna + 1).Tipo)
                    End If
                Next
            End If
            
            '4 - Imprime a linha
            Printer.Print strLinha
    
        Next
    End If
    Exit Sub
Trata:
    lblMensagem = Nvl(Mensagem, Err.Description)

End Sub

Private Function CalcularTotalPaginas(QuantidadeRegistros As Integer, LinhasPorPagina As Integer) As Integer
' VTOcx.grdVISUAL.Function CalcularTotalPaginas
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Calcula o numero de paginas que o relatorio ira necessitar
'
' Parametros : QuantidadeRegistros (Integer) - Qtos registros?
'              LinhasPorPagina (Integer) - Qto cabe na pagina?
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim dblQuantidadePaginas As Double
        
        dblQuantidadePaginas = QuantidadeRegistros / LinhasPorPagina
        CalcularTotalPaginas = CInt(dblQuantidadePaginas + 0.5) ' by Silmar Bosing
    End If
End Function

Private Sub ImprimirRodape(NumeroPagina As Integer)
' VTOcx.grdVISUAL.Sub ImprimirRodape
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Imprime o rodape do relatorio
'
' Parametros : NumeroPagina (Integer) - Numero da pagina a ser impressa
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim strLinha As String
    
        Printer.Font.Name = "Courier New"
        Printer.Font.Size = TamFonte
    
        '1 - Usuario / pagina
        Printer.Print
        strLinha = Me.RodapeUsuario & Space$(8) & "Página " & NumeroPagina & " de " & TotalPag
        Printer.Print AlinharCampo(strLinha, aliDireita, LarguraPagina)
        strLinha = String$(Len(strLinha), "-")
        Printer.Print AlinharCampo(strLinha, aliDireita, LarguraPagina)
        '2 - Data / hora
        strLinha = Format$(Now, "dd/mm/yyyy") & "      " & Format$(Now, "hh:mm:ss")
        Printer.Print AlinharCampo(strLinha, aliDireita, LarguraPagina)
    End If
End Sub

Private Function AlinharCampo(Campo As String, Posicao As enuAlinhamentoCampo, Tamanho As Integer) As String
' VTOcx.grdVISUAL.Function AlinharCampo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Posicionamento horizontal da informacao
'
' Parametros : Campo (String) - Informacao
'              Posicao (enuAlinhamentoCampo) - Localizacao horizontal
'              Tamanho (Integer) - Tamanho da pagina
'
' Ex: AlinharCampo("abc", aliEsquerda, 5) = "abc  "
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim intQuantidadeEspacos As Integer
    
        Select Case Posicao
    
            Case aliEsquerda
                intQuantidadeEspacos = 0
                
            Case aliCentro
                intQuantidadeEspacos = IIf(Len(Campo) > Tamanho, 0, (Tamanho - Len(Campo)) / 2)
            
            Case aliDireita
                intQuantidadeEspacos = Tamanho - Len(Campo)
    
        End Select
        
        AlinharCampo = String$(intQuantidadeEspacos, " ") & Campo
    End If
End Function

Private Function PreencherCampo(Campo As String, Tamanho As Integer, Tipo As enuTipoCampo) As String
' VTOcx.grdVISUAL.Function PreencherCampo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Torna as informacoes homogeneas, preenchendo quando necessario para atingir
'               o tamanho desejado
'
' Parametros : Campo (String) - Informacao
'              Tamanho (Integer) - Tamanho desejado
'              Tipo (enuTipoCampo) - Tipo do dado. Determina o caracter de preenchimento
'
' Ex: PreencherCampo("3", 5, tipInteiro) = "00003"
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim caractere As Long
        caractere = 32
       
        '1 - Consistencia do parametro Campo
        Dim str As String
        str = Campo
    
        If Tipo = tipMoeda Then
    
            str = Format$(Campo, "Standard")
    
        End If
        
        '2 - Trunca o tamanho do Campo
    
        If Tipo <> tipTexto Then
    
            str = Right$(str, Tamanho)
    
        Else
    
            str = Left$(str, Tamanho)
    
        End If
        
        '3 - Quantidade de posicoes que faltam preencher
        Dim i As Integer
        i = Tamanho - Len(str)
    
        If i < 0 Then i = 0
        
        '4 - Preenchimento das posicoes
        Dim preenche As String
        preenche = String$(i, caractere)
        
        If Tipo = tipTexto Then
    
            '5 - Alinha texto à esquerda
            str = UCase$(str) & preenche
    
        Else
    
            '6 - Alinha numericos à direita
            str = preenche & str
    
        End If
    
        Campo = str
        PreencherCampo = str
    End If
End Function


Private Sub FlatColumn(hwndTela As Long, Grid As Object)
End Sub

Public Property Get ColumnHeaders() As ColumnHeaders
    If bRegistrado Then Set ColumnHeaders = grdGrid.ColumnHeaders
End Property

Public Function FindItem(sz As String, Optional Where, Optional Index, Optional fPartial) As ListItem
    If bRegistrado Then Set FindItem = grdGrid.FindItem(sz, Where, Index, fPartial)
End Function

Public Sub MarcarTodos(Valor As Boolean)
    If bRegistrado Then
        Dim i As Integer
        
        For i = 1 To grdGrid.ListItems.Count
            grdGrid.ListItems(i).Checked = Valor
        Next
    End If
End Sub

Public Property Get Enabled() As Boolean

    If bRegistrado Then Enabled = grdGrid.Enabled

End Property

Public Property Let Enabled(ByVal vnewvalue As Boolean)

    If bRegistrado Then
        grdGrid.Enabled = vnewvalue
        cmdImprimir.Enabled = vnewvalue
        lblRotulo.Enabled = vnewvalue
        lblQtd.Enabled = vnewvalue
    End If
End Property

Public Function AtualizarQtd() As Long

    If bRegistrado Then
        lblQtd = grdGrid.ListItems.Count
        AtualizarQtd = grdGrid.ListItems.Count
    End If
End Function


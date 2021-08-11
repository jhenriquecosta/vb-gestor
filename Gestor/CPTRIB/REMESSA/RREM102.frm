VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{467EEF11-5281-4102-AFD3-AD54F754C329}#1.1#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.3#0"; "Cabecalho.ocx"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Begin VB.Form RREM102 
   BackColor       =   &H00FBEDE8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RREM102"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   4575
      Left            =   3690
      Pattern         =   "*.RET"
      TabIndex        =   7
      Top             =   660
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   75
      TabIndex        =   6
      Top             =   1095
      Width           =   3600
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   675
      Width           =   3615
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   1138
      Icone           =   "RREM102.frx":0000
      ImagemFundo     =   "RREM102.frx":031A
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   5805
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   847
      CorFundo        =   -2147483633
      ImagemFundo     =   "RREM102.frx":14274
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   60
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   1
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7140
         TabIndex        =   3
         Top             =   60
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
      End
      Begin VTOcx.cmdVISUAL cmdGravar 
         Height          =   375
         Left            =   5145
         TabIndex        =   2
         Top             =   60
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
         Caption         =   "Receber"
         Acao            =   3
      End
   End
   Begin Threed.SSPanel ssBarraAluno 
      Height          =   240
      Left            =   105
      TabIndex        =   10
      Top             =   5535
      Visible         =   0   'False
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   423
      _Version        =   196610
      ForeColor       =   16777215
      Windowless      =   -1  'True
      Caption         =   "SSPanel1"
      FloodType       =   1
      FloodFillStyle  =   1
      RoundedCorners  =   0   'False
   End
   Begin ActiveResizer.SSResizer SSResizer2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196610
      MinFontSize     =   1
      MaxFontSize     =   100
      DesignWidth     =   7980
      DesignHeight    =   6285
   End
   Begin VB.Label LblCaminho 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   9
      Top             =   5325
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CAMINHO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   5310
      Width           =   930
   End
End
Attribute VB_Name = "RREM102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravar_Click()
    On Error Resume Next
    Dim A  As Integer
    Dim Dados As String
    Dim varDOCUMENTO As String
    Dim varDT_OPERACAO  As String
    Dim varDT_CREDITO As String
    Dim varVL_DOCUMENTO  As String
    Dim varVL_ACRESCIMO  As String
    Dim SQL As String
    Dim Rs As VSRecordset
    Dim Campos As String
    Dim Valores As String
    Dim Linha As Integer
    'Pos.Inicial Pos.Final
    'TP.OPERACAO    014                014
    'DOCUMENTO      044                048         TP.OPERACAO = T
    'VL.DOCUMENTO   090                096         TP.OPERACAO = T
    'VL.ACRESCIMO   027                032         TP.OPERACAO = U
    'DT.OPERACAO    138                145         TP.OPERACAO = U
    'DT.CREDITO     146                153         TP.OPERACAO = U
    
    Const DOCUMENTO_INICIO = 44
    Const DOCUMENTO_fim = 48
    Const DATA_Operacao_INICIO = 138
    Const DATA_Operacao_FIM = 145
    Const TIPO_Operacao_INICIO = 14
    Const TIPO_Operacao_FIM = 14
    Const DT_CREDITO_INICIO = 146
    Const DT_CREDITO_FIM = 153
    Const VL_DOCUMENTO_INICIO = 85
    Const VL_DOCUMENTO_fim = 92
    Const VL_ACRESCIMO_INICIO = 27
    Const VL_ACRESCIMO_FIM = 32
    'Dim Rs As VSRecordset
    If LblCaminho.Caption = "" Then Exit Sub
    ssBarraAluno.Visible = True
    A = FreeFile
    Open LblCaminho For Input As #A
    Do Until EOF(A)
        
        Line Input #A, Dados
        If Mid(Dados, TIPO_Operacao_INICIO, 1) = "T" Then
            Linha = Linha + 1
            ssBarraAluno.FloodPercent = Linha
            varDOCUMENTO = 12000000 + Mid(Dados, DOCUMENTO_INICIO, DOCUMENTO_fim - DOCUMENTO_INICIO + 1)
        ElseIf Mid(Dados, TIPO_Operacao_INICIO, 1) = "U" Then
            varVL_DOCUMENTO = Abs(Mid(Dados, VL_DOCUMENTO_INICIO, VL_DOCUMENTO_fim - VL_DOCUMENTO_INICIO + 1))
            If varVL_DOCUMENTO <> "0" Then
                varVL_DOCUMENTO = Format(Left(varVL_DOCUMENTO, Len(varVL_DOCUMENTO) - 2) & "," & Right(varVL_DOCUMENTO, 2), Const_Monetario)
            End If
            varVL_ACRESCIMO = Abs(Mid(Dados, VL_ACRESCIMO_INICIO, VL_ACRESCIMO_FIM - VL_ACRESCIMO_INICIO + 1))
            If Val(varVL_ACRESCIMO) <> 0 Then
                varVL_ACRESCIMO = Format(Left(varVL_ACRESCIMO, Len(varVL_ACRESCIMO) - 2) & "," & Right(varVL_ACRESCIMO, 2), Const_Monetario)
            End If
            varDT_OPERACAO = Mid(Dados, DATA_Operacao_INICIO, DATA_Operacao_FIM - DATA_Operacao_INICIO + 1)
            varDT_OPERACAO = Format(Left(varDT_OPERACAO, 2), "00") & "/" & Mid(varDT_OPERACAO, 3, 2) & "/" & Right(varDT_OPERACAO, 4)
            varDT_CREDITO = Mid(Dados, DT_CREDITO_INICIO, DT_CREDITO_FIM - DT_CREDITO_INICIO + 1)
            varDT_CREDITO = Format(Left(varDT_CREDITO, 2), "00") & "/" & Mid(varDT_CREDITO, 3, 2) & "/" & Right(varDT_CREDITO, 4)
            Campos = "TCR_VALOR_PAGO,TCR_SALDO_DEVEDOR,TCR_MULTA,TCR_STATUS"
            Valores = Bdados.PreparaValor((CCur(varVL_DOCUMENTO)), 0, varVL_ACRESCIMO, esrQuitado)
            If PegaStatusRecebimento(varDOCUMENTO) = esrAberto And varDT_CREDITO <> "00/00/0000" Then
                If Bdados.GravaDados("TAB_CONTA_RECEBER", Valores, Campos, "TCR_COD_CONTA = " & varDOCUMENTO) Then
                    If Bdados.AbreTabela("Select tcr_valor,tcr_vencimento from tab_conta_receber where tcr_cod_conta = " & varDOCUMENTO, Rs) Then
                        If CDate(varDT_OPERACAO) > CDate(Rs.Fields("tcr_vencimento")) Then
                            Campos = "TCR_DESCONTO,tcr_valor_apagar"
                            Valores = Bdados.PreparaValor(0, Rs.Fields("tcr_valor"))
                            Call Bdados.GravaDados("TAB_CONTA_RECEBER", Valores, Campos, "TCR_COD_CONTA = " & varDOCUMENTO)
                        Else
                            Campos = "TCR_DESCONTO,tcr_valor_apagar"
                            Valores = Bdados.PreparaValor((Rs.Fields("tcr_valor") - CCur(varVL_DOCUMENTO)), CCur(varVL_DOCUMENTO))
                            Call Bdados.GravaDados("TAB_CONTA_RECEBER", Valores, Campos, "TCR_COD_CONTA = " & varDOCUMENTO)
                        End If
                    End If
                    Bdados.DeletaDados "TAB_BAIXA_RECEBIMENTO", "TBR_TCR_CODIGO = " & varDOCUMENTO
                    Campos = "TBR_TCR_CODIGO,"
                    Campos = Campos & "TBR_ORDEM,"
                    Campos = Campos & "TBR_OPERACAO,"
                    Campos = Campos & "TBR_VALOR_PAGO,"
                    Campos = Campos & "TBR_MULTA,"
                    Campos = Campos & "TBR_JUROS,"
                    Campos = Campos & "TBR_DESCONTO,"
                    Campos = Campos & "TBR_SUB_TOTAL,"
                    Campos = Campos & "TBR_FORMA_PAGAMENTO,"
                    Campos = Campos & "TBR_DATA_PAGAMENTO,"
                    Campos = Campos & "TBR_TCB_CONTA,"
                    Campos = Campos & "TBR_USUARIO,"
                    Campos = Campos & "TBR_TIPO_BAIXA"
                    Valores = Bdados.PreparaValor(varDOCUMENTO, 1, 1, CCur(varVL_DOCUMENTO) - CCur(varVL_ACRESCIMO), varVL_ACRESCIMO, 0, 0, (CCur(varVL_DOCUMENTO)), efpBoletoBancario - 1, varDT_CREDITO, 1, Aplica.Usuario, 2)
                    Bdados.InsereDados "TAB_BAIXA_RECEBIMENTO", Valores, Campos
                End If
            End If
        End If
    Loop
    Close A
    ssBarraAluno.FloodPercent = (ssBarraAluno.FloodPercent + (100 - ssBarraAluno.FloodPercent))
    If Linha = 0 Then
        Avisa "Arquivo sem movimentação."
    End If
    If Linha > 0 Then
        Avisa "Recepção concluída com sucesso"
    End If
    ssBarraAluno.FloodPercent = 0
    ssBarraAluno.Visible = False
End Sub

Private Sub cmdNovo_Click()
    LimpaCampos Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    
    File1.Path = Dir1.Path
    PegaCaminho
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    PegaCaminho
End Sub

Private Sub PegaCaminho()
    LblCaminho = File1.Path & "\" & File1.List(File1.ListIndex)
End Sub

Private Sub File1_Click()
    PegaCaminho
End Sub

Private Sub Form_Load()
    ssBarraAluno.FloodShowPct = True
    ssBarraAluno.FloodType = ssLeftToRight
    ssBarraAluno.FloodFillStyle = 0
End Sub

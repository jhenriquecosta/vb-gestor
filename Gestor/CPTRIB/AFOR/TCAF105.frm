VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCAF105 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCAF105.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1138
      Icone           =   "TCAF105.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   10
      Top             =   3195
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6300
         TabIndex        =   8
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5205
         TabIndex        =   7
         Top             =   120
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1005
      Left            =   30
      TabIndex        =   12
      Top             =   1680
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   1773
      Altura          =   1905
      Caption         =   " Informacões Cadastrais"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483626
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtData 
         Height          =   510
         Left            =   5760
         TabIndex        =   6
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         Caption         =   "Data"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   15
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtFolha 
         Height          =   510
         Left            =   4320
         TabIndex        =   5
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         Caption         =   "Folha"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   15
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtxOrdem 
         Height          =   510
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   900
         Caption         =   "No. de Ordem"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   15
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtFicha 
         Height          =   510
         Left            =   1515
         TabIndex        =   3
         Top             =   390
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         Caption         =   "Ficha"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   15
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtLivro 
         Height          =   510
         Left            =   2910
         TabIndex        =   4
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   900
         Caption         =   "Livro"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   15
         Mascara         =   "0000"
      End
   End
   Begin VTOcx.txtVISUAL txtConsultaIC 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   847
      Caption         =   "IC"
      Text            =   ""
      Restricao       =   2
      AlinhamentoRotulo=   1
      CorFundo        =   -2147483633
      MaxLen          =   15
   End
   Begin VTOcx.txtVISUAL txtConsultaLogr 
      Height          =   300
      Left            =   2520
      TabIndex        =   13
      Top             =   840
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      CorFundo        =   -2147483633
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaICConsulta 
      Height          =   315
      Left            =   2145
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtProp 
      Height          =   480
      Left            =   120
      TabIndex        =   15
      Top             =   1140
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   847
      Caption         =   "Proprietário"
      Text            =   ""
      Enabled         =   0   'False
      AlinhamentoRotulo=   1
      CorFundo        =   -2147483633
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtZona 
      Height          =   510
      Left            =   6480
      TabIndex        =   1
      Top             =   1110
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   900
      Caption         =   "Zona"
      Text            =   ""
      Restricao       =   2
      AlinhamentoRotulo=   1
      MaxLen          =   15
      Mascara         =   "0000"
   End
End
Attribute VB_Name = "TCAF105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AforManu As New cAforManu
Dim Aforamento As New cAforamento

Private Sub cmdPesquisaICConsulta_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtConsultaIC
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valida As Boolean
    Dim Campos As String
    Dim Valores As String
    
    Valida = False
    If Edita.CriticaCampos(Me) Then
        Campos = "TIM_AFORAMENTO_NUMERO,TIM_AFORAMENTO_FICHA,TIM_AFORAMENTO_LIVRO,TIM_AFORAMENTO_FOLHA,TIM_AFORAMENTO_DATA,TIM_ZONA"
        Valores = Bdados.PreparaValor(txtxOrdem, txtFicha, txtLivro, txtFolha, Bdados.Converte(txtData, TCDataHora), txtZona)
        If Bdados.GravaDados("TAB_IMOVEL", Valores, Campos, "TIM_IC ='" & txtConsultaIC & "'") Then
            Avisa "Registro gravdo com sucesso."
        Else
            Avisa "Dados não foram salvos."
        End If
        Edita.LimpaCampos Me
        txtConsultaIC.SetFocus
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AforManu = Nothing
    Set Aforamento = Nothing
End Sub

Private Sub txtConsultaIC_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    
    If txtConsultaIC = "" Then Exit Sub
    If Trim(txtConsultaIC) <> "" Then
        txtConsultaIC = BuscaContribuinte(txtConsultaIC, txtProp, txtConsultaLogr, , etiImovel)
        If Trim(txtConsultaIC) = "" Then
            Avisa "Inscricão não encontrada"
            txtConsultaIC.SetFocus
        End If
    End If
    Sql = "Select TIM_AFORAMENTO_NUMERO,TIM_AFORAMENTO_FICHA,TIM_AFORAMENTO_LIVRO,TIM_AFORAMENTO_FOLHA," & _
            " TIM_AFORAMENTO_DATA,TIM_ZONA from tab_imovel where tim_ic = '" & txtConsultaIC & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        txtxOrdem = "" & rs!TIM_AFORAMENTO_NUMERO
        txtFicha = "" & rs!TIM_AFORAMENTO_FICHA
        txtLivro = "" & rs!TIM_AFORAMENTO_LIVRO
        txtFolha = "" & rs!TIM_AFORAMENTO_FOLHA
        txtData = "" & rs!TIM_AFORAMENTO_DATA
        txtZona = "" & rs!TIM_ZONA
    Else
        Avisa "Imóvel não encontrado."
        
    End If
End Sub

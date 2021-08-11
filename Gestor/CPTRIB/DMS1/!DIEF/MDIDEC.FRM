VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIdec 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "SIAT - Sistema Integrado de Administração Tributária - SEFAZ - MA"
   ClientHeight    =   4425
   ClientLeft      =   1035
   ClientTop       =   2310
   ClientWidth     =   6765
   Icon            =   "MDIDEC.frx":0000
   LinkTopic       =   "MDIdec"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6765
      TabIndex        =   8
      Top             =   525
      Width           =   6765
      Begin VB.Label lbl_sair 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sair"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4620
         TabIndex        =   14
         Top             =   15
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbl_ajuda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tópicos de Ajuda"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3795
         TabIndex        =   13
         Top             =   15
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lbl_reparar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reparar Base"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2670
         TabIndex        =   12
         Top             =   15
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lbl_declaracao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Declaração"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   870
         TabIndex        =   11
         Top             =   15
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lbl_Diretorio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Arquivo de Declarações"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1125
         TabIndex        =   10
         Top             =   15
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lbl_cadastro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cadastro"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox BarraFerr 
      Align           =   1  'Align Top
      ForeColor       =   &H00C0C0C0&
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   6705
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      Begin VB.PictureBox SSPanel1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   0
         Left            =   6480
         ScaleHeight     =   405
         ScaleWidth      =   4905
         TabIndex        =   15
         Top             =   30
         Width           =   4965
      End
      Begin VB.PictureBox img_sair 
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   2760
         ScaleHeight     =   405
         ScaleWidth      =   495
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   30
         Width           =   555
      End
      Begin VB.PictureBox img_help 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   4050
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.PictureBox img_calculadora 
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   2220
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   30
         Width           =   525
      End
      Begin VB.PictureBox img_manutencao 
         AutoSize        =   -1  'True
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   1680
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   525
      End
      Begin VB.PictureBox img_diretorio 
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   1140
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   525
      End
      Begin VB.PictureBox img_declaracao 
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   600
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   525
      End
      Begin VB.PictureBox img_cadastro 
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   6000
         Picture         =   "MDIDEC.frx":0442
         Stretch         =   -1  'True
         Top             =   60
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog caminho 
      Left            =   45
      Top             =   1755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuGIM 
      Caption         =   "&DIEF"
      Begin VB.Menu MenCad 
         Caption         =   "&Cadastro de Contrbuinte e Contador"
      End
      Begin VB.Menu MenDir 
         Caption         =   "&Arquivo de Declarações"
      End
      Begin VB.Menu separador2 
         Caption         =   "-"
      End
      Begin VB.Menu MenDec 
         Caption         =   "&Declaração"
      End
      Begin VB.Menu Separador1 
         Caption         =   "-"
      End
      Begin VB.Menu Salir 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu MenMan 
      Caption         =   "&Manutenção"
      Index           =   4
      Begin VB.Menu MnuImport 
         Caption         =   "&Importar Arquivo"
      End
      Begin VB.Menu Copi 
         Caption         =   "&Cópia de Segurança"
         Begin VB.Menu mnuGerarCS 
            Caption         =   "&Gerar"
         End
         Begin VB.Menu mnuRestCS 
            Caption         =   "Re&staurar"
         End
      End
      Begin VB.Menu separador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReparar 
         Caption         =   "&Reparar Base"
      End
   End
   Begin VB.Menu Ajuda 
      Caption         =   "&Ajuda"
      Index           =   5
      Begin VB.Menu mnuTopAju 
         Caption         =   "&Tópicos de Ajuda"
      End
      Begin VB.Menu Separador4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre a DIEF - Eletrônica..."
      End
   End
End
Attribute VB_Name = "MDIdec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Ajuda_Click(Index As Integer)
'Form2.Show
End Sub




Private Sub BarraFerr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_cadastro.Visible = False
lbl_declaracao.Visible = False
lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = False
End Sub


Private Sub BarraFerr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub

Private Sub img_cadastro_Click()
limpa_baloes
MenCad_Click
End Sub

Private Sub img_cadastro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbl_cadastro.Visible = True

lbl_declaracao.Visible = False

lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = False



End Sub


Private Sub img_cadastro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 lbl_cadastro.Visible = False
 
End Sub

Private Sub img_calculadora_Click()
Dim MyAppID

On Error GoTo errocalc:
    
If rota = True Then
    lbl_reparar.Visible = False
    MyAppID = Shell("C:\WINDOWS\CALC.EXE", 1) ' Run Calculator.
    limpa_baloes
    rota = False
Else
    
    AppActivate "Calculadora"

End If

Exit Sub


errocalc:
If Err = 5 Then
    lbl_reparar.Visible = False
    MyAppID = Shell("C:\WINDOWS\CALC.EXE", 1) ' Run Calculator.
    limpa_baloes
End If

End Sub

Private Sub img_calculadora_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbl_cadastro.Visible = False

lbl_declaracao.Visible = False

lbl_Diretorio.Visible = False

lbl_reparar.Left = 1725
lbl_reparar.Caption = "Calculadora"
lbl_reparar.Visible = True
lbl_ajuda.Visible = False
lbl_sair.Visible = False


End Sub


Private Sub img_calculadora_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub

Private Sub img_declaracao_Click()
limpa_baloes
gformaAnte = 0
MenDec_Click
End Sub

Private Sub img_declaracao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbl_cadastro.Visible = False

lbl_declaracao.Visible = True
lbl_declaracao.Left = 540
lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = False

End Sub


Private Sub img_declaracao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_declaracao.Visible = False
End Sub

Private Sub img_diretorio_Click()
   limpa_baloes
   gformaAnte = 1
   MenDir_Click
   
End Sub

Private Sub img_diretorio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_cadastro.Visible = False

lbl_declaracao.Visible = False

lbl_Diretorio.Visible = True
'lbl_Diretorio = "Diretório"
lbl_Diretorio.Left = 1155
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = False

End Sub


Private Sub img_diretorio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lbl_Diretorio.Visible = False
End Sub


Private Sub img_help_Click()
'limpa_baloes
'mnuTopAju_Click
End Sub

Private Sub img_help_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbl_cadastro.Visible = False
lbl_declaracao.Visible = False
lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = True
lbl_ajuda.Left = 2370
lbl_sair.Visible = False

End Sub


Private Sub img_help_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_ajuda.Visible = False
End Sub


Private Sub img_manutencao_Click()
'If mdkidec.formscount - 1 > 1 Then

    limpa_baloes
    mnuReparar_Click
    
End Sub

Private Sub img_manutencao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_cadastro.Visible = False

lbl_declaracao.Visible = False

lbl_Diretorio.Visible = False
lbl_reparar.Left = 1725

lbl_reparar.Caption = "Reparar Base"

lbl_reparar.Visible = True
lbl_ajuda.Visible = False
lbl_sair.Visible = False

End Sub


Private Sub img_manutencao_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_reparar.Visible = False
End Sub


Private Sub img_sair_Click()
   limpa_baloes
   salir_Click
   
End Sub

Private Sub img_sair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbl_cadastro.Visible = False
lbl_declaracao.Visible = False
lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = True
lbl_sair.Left = 2700


End Sub


Private Sub img_sair_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl_sair.Visible = False

End Sub


Private Sub MDIForm_Activate()
Dim mensagem As String
If primeira_vez = True Then
    rota = True
    primeira_vez = False
    sobre_frm.Show 1
End If

End Sub

Private Sub MDIForm_Load()
 
 Versao = "V00"
 abre_base_gim
 primeira_vez = True
 gformulario = 0
 App.HelpFile = App.Path & "\helpdief.HLP"
 '  TesteDBgeral
 'Inicializa Tab UF
 
    TabUf(1) = "AC"
    TabUf(2) = "AL"
    TabUf(3) = "AM"
    TabUf(4) = "AP"
    TabUf(5) = "BA"
    TabUf(6) = "CE"
    TabUf(7) = "DF"
    TabUf(8) = "ES"
    TabUf(9) = "GO"
    TabUf(10) = "MA"
    TabUf(11) = "MG"
    TabUf(12) = "MS"
    TabUf(13) = "MT"
    TabUf(14) = "PA"
    TabUf(15) = "PB"
    TabUf(16) = "PE"
    TabUf(17) = "PI"
    TabUf(18) = "PR"
    TabUf(19) = "RJ"
    TabUf(20) = "RN"
    TabUf(21) = "RO"
    TabUf(22) = "RR"
    TabUf(23) = "RS"
    TabUf(24) = "SC"
    TabUf(25) = "SE"
    TabUf(26) = "SP"
    TabUf(27) = "TO"

End Sub


Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
limpa_baloes
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
'ReturnValue = Shell("Calc.exe", 1)  ' Run Calculator.
'AppActivate ReturnValue     ' Activate the Calculator.

'For i = 1 To 100    ' Set up counting loop.''

'SendKeys i & "{+}", True    ' Send keystrokes to Calculator

'Next i  ' to add each value of I.
'SendKeys "=", True  ' Get grand total.

SendKeys "%{F4}", True  ' Send ALT+F4 to close Calculator.
   Close
   End
   
End Sub


Private Sub MenCad_Click()
On Error GoTo erro:


    lbl_cadastro.Visible = False

dec001.Show 1

Exit Sub
erro:
If Err > 0 Then
    dec001.Show
End If
End Sub

Private Sub MenDec_Click()
teve_alteracao_declaracao = False
 teve_alteracao = False
lbl_declaracao.Visible = False
g_condicao = 0
 g_num_inscricao = ""
g_periodo_inicial = ""
g_tipo_decla = 0
   quem_me_chamou = ""
 dec002.Show  ' 1
End Sub

Private Sub MenDir_Click()
teve_alteracao = False
teve_alteracao_contador = False
teve_alteracao_declaracao = False
gformulario = 3
    lbl_Diretorio.Visible = False
    g_condicao = 0
    g_num_inscricao = ""
    g_periodo_inicial = ""
    g_tipo_decla = 0
    'lbl_Diretorio = ""
    dec003.Show 1

End Sub

Private Sub mnuGerarCS_Click()

 On Error Resume Next


    Dim copia
    Dim SourceFile
    Dim DestinationFile
    Me.caminho.DialogTitle = "Informe o caminho onde deseja guardar a Cópia de Segurança"
    Me.caminho.Filter = "Arquivos de Dados do GIM |Dief.mdb|"
    Me.caminho.InitDir = App.Path
    Me.caminho.Action = 2
    Me.caminho.FileName = "Dief.bak"
    On Error GoTo Erro_Copia:

    SourceFile = App.Path + "\Dief.mdb"
    DestinationFile = caminho.FileName
    dbgim.Close
    FileCopy SourceFile, DestinationFile
    abre_base_gim
    
    Exit Sub

Erro_Copia:
MsgBox "Error não geração da Cópia, tentar novamente.", 33, NOME_SISTEMA_DEFAULT
Resume Next

End Sub

Private Sub MnuImport_Click()
lbl_cadastro.Visible = False
lbl_declaracao.Visible = False
lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = False
ImpEnvio.Show 1
End Sub

Private Sub mnuReparar_Click()
Dim errLoop As Error

    If MsgBox("Corrigir o banco de dados Dief.MDB?", _
            vbYesNo) = vbYes Then
        On Error GoTo Err_Repair
        dbgim.Close
        RepairDatabase App.Path & "\Dief.MDB"
        On Error GoTo 0
        MsgBox "Fim do procedimento de correção!", vbInformation, NOME_SISTEMA_DEFAULT
        abre_base_gim
    End If

    Exit Sub

Err_Repair:

For Each errLoop In DBEngine.Errors
 MsgBox "Correção mal-sucedida!" & vbCr & _
    "Número do erro: " & errLoop.Number & _
    vbCr & errLoop.Description, 48, NOME_SISTEMA_DEFAULT

Next errLoop
End Sub

Private Sub mnuRestCS_Click()
 On Error Resume Next
 Dim copia
    Dim SourceFile
    Dim DestinationFile
    Me.caminho.DialogTitle = "Restaurar Cópia de Segurança [Dief.bak]"
    Me.caminho.Filter = "Arquivos de backup do GIM |Dief.bak|"
    Me.caminho.InitDir = App.Path
    Me.caminho.Action = 1
    Me.caminho.FileName = "Dief.bak"
    SourceFile = App.Path + "\Dief.mdb"
     On Error GoTo Erro_Copia
    DestinationFile = caminho.FileName
    dbgim.Close
    FileCopy DestinationFile, SourceFile
    abre_base_gim
    
    Exit Sub

Erro_Copia:
MsgBox "Error na recuperação da Cópia de segurança, tentar novamente.", 33, NOME_SISTEMA_DEFAULT
Resume Next


End Sub

Private Sub mnuSobre_Click()
   sobre_frm.Show 1
End Sub

Private Sub mnuTopAju_Click()
lbl_ajuda.Visible = False
SendKeys "{F1}"
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lbl_cadastro.Visible = False
lbl_declaracao.Visible = False
lbl_Diretorio.Visible = False
lbl_reparar.Visible = False
lbl_ajuda.Visible = False
lbl_sair.Visible = False

End Sub


Private Sub salir_Click()
Dim result As Integer
result = MsgBox("Deseja Realmente Sair do Sistema", 33, "Aviso do Sistema")

If result = 1 Then
   End
End If
   
End Sub





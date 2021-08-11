VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form TINT101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Tag             =   "TDEC104"
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Height          =   555
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   5910
         TabIndex        =   6
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         Caption         =   "&Sair"
         Acao            =   7
      End
      Begin VB.PictureBox picProgresso 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   2640
         ScaleHeight     =   465
         ScaleWidth      =   2475
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   2475
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   180
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgresso 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2385
         End
      End
      Begin VB.TextBox txtInformacoes 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2700
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   2295
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   1138
      Icone           =   "TINT101.frx":0000
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   3240
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5715
      Altura          =   1905
      Caption         =   " Enviar Arquivo"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdEnviar 
         Height          =   315
         Left            =   5730
         TabIndex        =   7
         Top             =   360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "&Enviar"
         Acao            =   3
      End
      Begin VTOcx.cmdVISUAL cmdLocalizarArquivo 
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   390
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtArquivo 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   390
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   503
         Caption         =   "Arquivo"
         Text            =   ""
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.grdVISUAL grdArquivos 
         Height          =   2595
         Left            =   90
         TabIndex        =   4
         Top             =   780
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   4577
         CorBorda        =   8421504
         Caption         =   "Arquivos enviados"
         CorTitulo       =   12632256
         OcultarRodape   =   -1  'True
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6360
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RemotePort      =   21
      URL             =   "http://"
      RequestTimeout  =   1800
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   4950
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Localizar Arquivo"
      Filter          =   "Declarações Geradas|*.DEC"
   End
End
Attribute VB_Name = "TINT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents SendEmail As vbSendMail.clsSendMail
Attribute SendEmail.VB_VarHelpID = -1
Dim strStatus As String

Private Sub AtivarControls(Optional Ativar As Boolean = True)
    On Error Resume Next
    Dim Controle
    
    For Each Controle In Me.Controls
        If Not Controle = lblProgresso And Not Controle = ProgressBar1 Then
            Controle.Enabled = Ativar
        End If
    Next
End Sub



Public Function inetReady(Message As Boolean) As Boolean
On Error GoTo errhandler
Dim msg As String

If Inet1.StillExecuting Then
    If Message Then
        msg = "O sistema não terminou "
        msg = msg & "de execuatar o último pedido. Por favor, aguarde"
        Util.Avisa msg
    
    End If
    inetReady = False
Else
    inetReady = True
End If
Exit Function
errhandler:
    Util.Erro Err.Source & " " & Err.Number & " " & Err.Description
End Function

Private Sub EnviarViaFTP()
    On Error GoTo TRATA
    Dim strFile As String
    
    
    Screen.MousePointer = vbHourglass
    picProgresso.Visible = True
    
    With Inet1
        AtivarControls False
        .Protocol = icFTP
        .URL = Temp.PegaParametro(Bdados, "TRANSMISSAO_DESTINO")
        .UserName = Temp.PegaParametro(Bdados, "TRANSMISSAO_USUARIO")
        .Password = Temp.PegaParametro(Bdados, "TRANSMISSAO_SENHA")
        If inetReady(True) Then
            bolFlag = True
            strFile = Temp.PegaParametro(Bdados, "TRANSMISSAO_PASTA") & "DMS-" & Format(Date, "dd-mm-yyyy") & ".DEC"
            ProgressBar1.Max = FileLen(txtArquivo.Text)
            .Execute , "PUT " & txtArquivo.Text & " " & strFile
        End If
        '.Execute , "pwd"  'Força o erro
    End With
    
    Exit Sub
    
TRATA:
    Select Case Err.Number
        Case 35764        '  Still executes last command
            DoEvents
            If bolFlag Then        ' File transfer
                Inet1.Execute , "size " & strFile
                ProgressBar1.Value = Val(Inet1.GetChunk(1024))
                ProgressBar1.ToolTipText = CInt(ProgressBar1.Value * 100 / ProgressBar1.Max) & "% transmitido"
            End If
            Resume
        Case 0
            Util.Erro "Não foi possível detectar uma conexão com a internet"
        Case Else
            AtivarControls
            Screen.MousePointer = vbDefault
            picProgresso.Visible = False
    End Select
    
    Exit Sub
End Sub
    
Private Sub AtualizaGrade()
    Dim Sql As String
    Sql = "select ted_arquivo as Arquivo, ted_data_hora as Data_Envio, ted_tus_cod_usuario as Usuario  from tab_envio_declaracao ORDER BY TED_CODIGO DESC"
    grdArquivos.Preencher Bdados, Sql, 4000, 1800, 1600
End Sub

Public Sub EnviarEmail(Assunto As String)
    Screen.MousePointer = vbHourglass
        
    With SendEmail
        AtivarControls False
        .SMTPHost = Temp.PegaParametro(Bdados, "servidor_smtp")                 ' Required the fist time, optional thereafter
        .From = Temp.PegaParametro(Bdados, "email")                       ' Required the fist time, optional thereafter
        .FromDisplayName = "Gestor Municipal"         ' Optional, saved after first use
        .Recipient = Temp.PegaParametro(Bdados, "TRANSMISSAO_DESTINO")                    ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = Temp.PegaParametro(Bdados, "cliente")     ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = .From              ' Optional, used when different than 'From' address
        .Subject = Assunto                  ' Optional
        .Message = "Segue anexo o arquivo de declarações"                      ' Optional
        .Attachment = txtArquivo.Text          ' Optional, separate multiple entries with delimiter character
        .Receipt = False                        ' Optional, default = FALSE
        .UseAuthentication = True             ' Optional, default = FALSE
        .UsePopAuthentication = True           ' Optional, default = FALSE
        .UserName = Temp.PegaParametro(Bdados, "TRANSMISSAO_USUARIO")                    ' Optional, default = Null String
        .Password = Temp.PegaParametro(Bdados, "TRANSMISSAO_SENHA")                    ' Optional, default = Null String, value is NOT saved
        .POP3Host = Temp.PegaParametro(Bdados, "SERVIDOR_POP3")
        '.Connect
        
        ProgressBar1.Value = 0
        picProgresso.Visible = True
        .Send
    End With
  
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdEnviar_Click()
    Dim ArqDeclaracao As New ArquivoDeclaracao
    If txtArquivo.Text = "" Then
        Util.Avisa "Informe o arquivo a ser enviado"
        cmdLocalizarArquivo_Click
        Exit Sub
    End If
    
    If Util.Confirma("Confirma o envio do arquivo?") = True Then
        If ArqDeclaracao.ValidaArquivo(txtArquivo.Text) Then
            If Temp.PegaParametro(Bdados, "TRANSMISSAO") = "EMAIL" Then
                EnviarEmail " ARQUIVO DMS.DEC - " & ArqDeclaracao.DataArquivo
            ElseIf Temp.PegaParametro(Bdados, "TRANSMISSAO") = "FTP" Then
                EnviarViaFTP
            Else
                Util.Avisa "Não existe modo de transmissão definido. Vá no menu 'Gestão de Configurações - Parâmetros - Parâmetros do Sistema' e informe esta opção."
                Exit Sub
            End If
        Else
            Avisa "Arquivo de declaracão inválido."
            txtArquivo.SetFocus
        End If
    End If
End Sub

Private Sub cmdLocalizarArquivo_Click()
    With Dialogo
        .InitDir = "C:\"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        txtArquivo.Text = .FileName
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If cmdEnviar.Enabled = True Then
        txtInformacoes.Text = "Modo Transmissão: " & Temp.PegaParametro(Bdados, "TRANSMISSAO") & vbCrLf _
        & "Endereço: " & Temp.PegaParametro(Bdados, "TRANSMISSAO_DESTINO")
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo TRATA
    
    Set SendEmail = New vbSendMail.clsSendMail
    cabVISUAL1.Exibir Bdados, Me.Tag, App.Path
    rodVISUAL1.Exibir Bdados, Me.Tag, App.Major, App.Minor, App.Revision
    AtualizaGrade
    
    Exit Sub
TRATA:
    Util.Erro Err.Description
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
On Error GoTo TRATA
Dim strValores As String
Dim strCampos As String

Select Case State
    
    Case icResolvingHost
        strStatus = "Inicializando..."

    Case icHostResolved
        strStatus = "Inicializando..."
        
    Case icConnecting
        strStatus = "Conectando..."


    Case icConnected
        strStatus = "Conectado"

    Case icRequesting
        strStatus = "Enviando..."
        
    Case icRequestSent
        strStatus = "Enviando..."

    Case icReceivingResponse
        strStatus = "Recebendo Resposta..."

    Case icResponseReceived
        strStatus = "Resposta Recebida"

    Case icDisconnecting
        strStatus = "Desconectando..."


    Case icDisconnected
        strStatus = "Desconectado"

    Case icError
        strStatus = "Erro"
        'IF Inet1.ResponseCode
        Debug.Print Err.Number

    Case icResponseCompleted
        strStatus = "Concluído"
        
End Select
    lblProgresso.Caption = strStatus
        
    If strStatus = "Erro" Then
        If Inet1.ResponseCode = 12007 Then
            Util.Erro "Não foi possível detectar uma conexão com a internet"
        Else
            Util.Erro Inet1.ResponseCode & Inet1.ResponseInfo
        End If
        AtivarControls
    ElseIf strStatus = "Concluído" Then
        Util.Avisa "O arquivo foi enviado com sucesso"
        Inet1.Cancel
        lblProgresso.Caption = ""
        picProgresso.Visible = False
        Screen.MousePointer = vbDefault
        
        'GRAVA AS INFORMAÇÃOES NO BANCO DE DADOS
        strValores = Bdados.PreparaValor(txtArquivo.Text, AplicacoesVTFuncoes.Usuario, "FTP")
        strCampos = "TED_ARQUIVO, TED_TUS_COD_USUARIO, TED_MODO_ENVIO"
        Bdados.InsereDados "TAB_ENVIO_DECLARACAO", strValores, strCampos
        txtArquivo.Text = ""
        AtualizaGrade
        AtivarControls
    End If
Exit Sub
TRATA:
    Util.Erro Err.Source & " " & Err.Number & " " & Err.Description
End Sub

Private Sub SendEmail_Progress(PercentComplete As Long)
    ProgressBar1.Value = PercentComplete
End Sub

Private Sub SendEmail_SendFailed(Explanation As String)
    strStatus = "Erro"
    SendEmail_Status "Erro"
    If Left(Explanation, 10) = "Valid name" Then
        Util.Erro "Não foi possível detetar uma conexão com a Internet"
    Else
        Util.Erro "Não foi possível enviar o arquivo:" & vbCrLf & Explanation
    End If
    ProgressBar1.Value = 0
    picProgresso.Visible = False
    AtivarControls
End Sub

Private Sub SendEmail_SendSuccesful()
    Dim strValores As String
    Dim strCampos As String
    
    strStatus = "Concluído"
    SendEmail_Status strStatus
    
    Util.Avisa "Arquivo enviado com sucesso."
    ProgressBar1.Value = 0
    picProgresso.Visible = False
    
    strValores = Bdados.PreparaValor(txtArquivo.Text, AplicacoesVTFuncoes.Usuario)
    strCampos = "TED_ARQUIVO, TED_TUS_COD_USUARIO"
    Bdados.InsereDados "TAB_ENVIO_DECLARACAO", strValores, strCampos
    txtArquivo.Text = ""
    AtualizaGrade
    AtivarControls
End Sub

Private Sub SendEmail_Status(Status As String)
    If Left(Status, Len("Connecting")) = "Connecting" Then
        strStatus = "Conectando..."
    ElseIf Status = "Initializing Communications..." Then
        strStatus = "Inicializando..."
    ElseIf Status = "Sending Message..." Then
        strStatus = "Enviando..."
    ElseIf Status = "Transmission Complete..." Then
        strStatus = "Enviado"
    ElseIf Status = "Closing Connection..." Then
        strStatus = "Desconectando..."
    End If
    
    lblProgresso.Caption = strStatus
End Sub

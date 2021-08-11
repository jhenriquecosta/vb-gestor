VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDEC104 
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
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   979
      Begin VB.PictureBox picProgresso 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   2640
         ScaleHeight     =   465
         ScaleWidth      =   2475
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   2475
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   0
            Width           =   2385
         End
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   6450
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Localizar Arquivo"
      Filter          =   "Declarações Geradas|*.DEC"
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   3240
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   660
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5715
      Altura          =   1905
      Caption         =   " Enviar Arquivo"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdEnviar 
         Height          =   345
         Left            =   5730
         TabIndex        =   6
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         Caption         =   "&Enviar"
         Acao            =   3
         CorBorda        =   16711680
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLocalizarArquivo 
         Height          =   345
         Left            =   5310
         TabIndex        =   2
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtArquivo 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   503
         Caption         =   "Arquivo"
         Text            =   ""
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.grdVISUAL grdArquivos 
         Height          =   2595
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   4577
         CorBorda        =   16711680
         Caption         =   "Arquivos enviados"
         CorTitulo       =   16711680
         OcultarRodape   =   -1  'True
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1138
      Icone           =   "TDEC104.frx":0000
   End
End
Attribute VB_Name = "TDEC104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents SendEmail As vbSendMail.clsSendMail
Dim strStatus As String

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
        .Protocol = icFTP
        .URL = Temp.PegaParametro(Bdados, "TRANSMISSAO_DESTINO")
        .Username = Temp.PegaParametro(Bdados, "TRANSMISSAO_USUARIO")
        .Password = Temp.PegaParametro(Bdados, "TRANSMISSAO_SENHA")
        If inetReady(True) Then
            bolFlag = True
            strFile = Temp.PegaParametro(Bdados, "TRANSMISSAO_PASTA") & "DMS.DEC"
            ProgressBar1.Max = FileLen(txtArquivo.Text)
            .Execute , "PUT " & txtArquivo.Text & " " & strFile
        Else
            Exit Sub
        End If
        .Execute , "pwd"  'Força o erro
    End With
    
    Exit Sub
    
TRATA:
    Select Case Err.Number
        Case 35764        '  Still executes last command
            DoEvents
            If bolFlag Then        ' File transfer
                If Not (Dir(strFile) = "") Then
                    Inet1.Execute , "size " & strFile
                    ProgressBar1.Value = Inet1.GetChunk(1024)
                    ProgressBar1.ToolTipText = CInt(ProgressBar1.Value * 100 / ProgressBar1.Max) & "% transmitido"
                End If
            End If
            Resume
        Case 0
            Util.Erro "Não foi possivel detectar uma conexão com a internet"
        Case Else
            Stop
    End Select
End Sub
    
Private Sub AtualizaGrade()
    Dim Sql As String
    Sql = "select ted_arquivo as Arquivo, ted_data_hora as Data_Envio, ted_tus_cod_usuario as Usuario  from tab_envio_declaracao ORDER BY TED_CODIGO DESC"
    grdArquivos.Preencher Bdados, Sql, 4000, 1800, 1600
End Sub

Public Sub EnviarEmail(Assunto As String)
    Screen.MousePointer = vbHourglass
        
'    With SendEmail
'        .SMTPHost = Temp.PegaParametro(Bdados, "servidor_smtp")                 ' Required the fist time, optional thereafter
'        .From = Temp.PegaParametro(Bdados, "email")                       ' Required the fist time, optional thereafter
'        .FromDisplayName = "Gestor Municipal"         ' Optional, saved after first use
'        .Recipient = Temp.PegaParametro(Bdados, "TRANSMISSAO_DESTINO")                    ' Required, separate multiple entries with delimiter character
'        .RecipientDisplayName = Temp.PegaParametro(Bdados, "cliente")     ' Optional, separate multiple entries with delimiter character
'        .ReplyToAddress = .From              ' Optional, used when different than 'From' address
'        .Subject = Assunto                  ' Optional
'        .Message = "Segue anexo o arquivo de declarações"                      ' Optional
'        .Attachment = txtArquivo.Text          ' Optional, separate multiple entries with delimiter character
'        .Receipt = False                        ' Optional, default = FALSE
'        .UseAuthentication = True             ' Optional, default = FALSE
'        .UsePopAuthentication = True           ' Optional, default = FALSE
'        .UserName = Temp.PegaParametro(Bdados, "TRANSMISSAO_USUARIO")                    ' Optional, default = Null String
'        .Password = Temp.PegaParametro(Bdados, "TRANSMISSAO_SENHA")                    ' Optional, default = Null String, value is NOT saved
'        .POP3Host = Temp.PegaParametro(Bdados, "SERVIDOR_POP3")
'        '.Connect
'
'        ProgressBar1.Value = 0
'        picProgresso.Visible = True
'        .Send
'    End With
'
'    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdEnviar_Click()
    Dim ArqDeclaracao As New ArquivoDeclaracao
    If txtArquivo.Text = "" Then
        Util.Avisa "Informe o arquivo a ser enviado"
        cmdLocalizarArquivo_Click
        Exit Sub
    End If
    
    If Util.Confirma("confirma o envio do arquivo?") = True Then
        'If ArqDeclaracao.ValidaArquivo(txtArquivo.Text) Then
            If Temp.PegaParametro(Bdados, "TRANSMISSAO") = "EMAIL" Then
                EnviarEmail " ARQUIVO DMS.DEC - " & ArqDeclaracao.DataArquivo
            ElseIf Temp.PegaParametro(Bdados, "TRANSMISSAO") = "FTP" Then
                EnviarViaFTP
            Else
                Util.Avisa "Não existe modo de transmissão definido. Vá no menu 'Gestão de Configurações - Parâmetros - Parâmetros do Sistema' e informe esta opção."
                Exit Sub
            End If
        'Else
        '    Avisa "Arquivo de declaracão inválido."
        '    txtArquivo.SetFocus
        'End If
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

Private Sub Form_Load()
    On Error GoTo TRATA
    
'    Set SendEmail = New vbSendMail.clsSendMail
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
        strStatus = "Finalizando..."

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

    Case icResponseCompleted
        strStatus = "Pedido Completado"
        
End Select
    lblProgresso.Caption = strStatus
        
    If strStatus = "Erro" Then
        Util.Erro Inet1.ResponseCode & Inet1.ResponseInfo
    ElseIf strStatus = "Pedido Completado" Then
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
        
    End If
Exit Sub
TRATA:
    Util.Erro Err.Source & " " & Err.Number & " " & Err.Description
End Sub
'
'Private Sub SendEmail_Progress(PercentComplete As Long)
'    ProgressBar1.Value = PercentComplete
'End Sub

'Private Sub SendEmail_SendFailed(Explanation As String)
'    strStatus = "Erro"
'    SendEmail_Status "Erro"
'    If Left(Explanation, 10) = "Valid name" Then
'        Util.Erro "Não foi possível detetar uma conexão com a Internet"
'    Else
'        Util.Erro "Não foi possível enviar o arquivo:" & vbCrLf & Explanation
'    End If
'    ProgressBar1.Value = 0
'    picProgresso.Visible = False
'End Sub
'
'Private Sub SendEmail_SendSuccesful()
'    Dim strValores As String
'    Dim strCampos As String
'
'    strStatus = "Concluído"
'    SendEmail_Status strStatus
'
'    Util.Avisa "Arquivo enviado com sucesso."
'    ProgressBar1.Value = 0
'    picProgresso.Visible = False
'
'    strValores = Bdados.PreparaValor(txtArquivo.Text, AplicacoesVTFuncoes.Usuario)
'    strCampos = "TED_ARQUIVO, TED_TUS_COD_USUARIO"
'    Bdados.InsereDados "TAB_ENVIO_DECLARACAO", strValores, strCampos
'    txtArquivo.Text = ""
'    AtualizaGrade
'End Sub
'
'Private Sub SendEmail_Status(Status As String)
'    If Left(Status, Len("Connecting")) = "Connecting" Then
'        strStatus = "Conectando..."
'    ElseIf Status = "Initializing Communications..." Then
'        strStatus = "Inicializando..."
'    ElseIf Status = "Sending Message..." Then
'        strStatus = "Enviando..."
'    ElseIf Status = "Transmission Complete..." Then
'        strStatus = "Enviado"
'    ElseIf Status = "Closing Connection..." Then
'        strStatus = "Desconectando..."
'    End If
'
'    lblProgresso.Caption = strStatus
'End Sub

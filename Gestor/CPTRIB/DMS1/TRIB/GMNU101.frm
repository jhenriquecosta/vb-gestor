VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form GMNU101 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CIAP"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GMNU101.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   649
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTree 
      BackColor       =   &H00DC7E5A&
      BorderStyle     =   0  'None
      Height          =   3660
      Left            =   375
      ScaleHeight     =   3660
      ScaleWidth      =   6975
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.TreeView treSis 
         Height          =   3630
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   6403
         _Version        =   393217
         Indentation     =   538
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlMenu"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer tmrSis 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30
      Top             =   4890
   End
   Begin VB.PictureBox picTrocaSenha 
      BackColor       =   &H00FFD3B3&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   4200
      ScaleHeight     =   1050
      ScaleWidth      =   2775
      TabIndex        =   18
      Top             =   2610
      Width           =   2775
      Begin VB.TextBox txtConfirmar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Digite novamente sua Nova Senha"
         Top             =   720
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtNova 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Digite uma Nova Senha se desejar alterar sua Senha"
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Troca de senha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   45
         TabIndex        =   21
         Top             =   30
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nova"
         ForeColor       =   &H00A65C02&
         Height          =   195
         Index           =   2
         Left            =   690
         TabIndex        =   20
         Top             =   375
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmação"
         ForeColor       =   &H00A65C02&
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   19
         Top             =   720
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   285
         Index           =   4
         Left            =   1170
         Top             =   330
         Width           =   1485
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   285
         Index           =   5
         Left            =   1170
         Top             =   690
         Width           =   1485
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00A65C02&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   2835
      End
   End
   Begin VB.PictureBox picUsuario 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3570
      ScaleHeight     =   735
      ScaleWidth      =   5715
      TabIndex        =   14
      Top             =   1575
      Width           =   5715
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   660
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Senha do Usuário"
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Nome do Usuário"
         Top             =   30
         Width           =   3690
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   660
         TabIndex        =   0
         ToolTipText     =   "Código do Usuário"
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   405
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   45
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   285
         Index           =   2
         Left            =   630
         Top             =   360
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   285
         Index           =   0
         Left            =   630
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.PictureBox picProgress 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1860
      ScaleHeight     =   315
      ScaleWidth      =   6570
      TabIndex        =   11
      Top             =   4050
      Visible         =   0   'False
      Width           =   6570
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   105
         Left            =   30
         TabIndex        =   12
         Top             =   195
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando ..."
         ForeColor       =   &H00868893&
         Height          =   195
         Index           =   4
         Left            =   -30
         TabIndex        =   13
         Top             =   -30
         Width           =   1155
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   45
      Left            =   10680
      TabIndex        =   8
      Top             =   180
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   79
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   8130
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   7995
      Picture         =   "GMNU101.frx":030A
      Top             =   345
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DC7E5A&
      Index           =   1
      X1              =   648
      X2              =   75
      Y1              =   318
      Y2              =   318
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DC7E5A&
      Index           =   0
      X1              =   420
      X2              =   0
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label lblSistema 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEMA GESTOR MUNICIPAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A65C02&
      Height          =   240
      Left            =   540
      TabIndex        =   27
      Top             =   90
      Width           =   5775
   End
   Begin VB.Label lblSis 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00868893&
      Height          =   900
      Left            =   7380
      TabIndex        =   26
      Top             =   1455
      Visible         =   0   'False
      Width           =   2340
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescrSis 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00868893&
      Height          =   1140
      Left            =   7575
      TabIndex        =   25
      Top             =   2010
      Visible         =   0   'False
      Width           =   2055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUsuario 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00868893&
      Height          =   195
      Left            =   5925
      TabIndex        =   24
      Top             =   390
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00868893&
      Height          =   195
      Left            =   60
      TabIndex        =   23
      Top             =   5160
      Width           =   45
   End
   Begin VB.Label lblResp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Responsável"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   1155
      TabIndex        =   10
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ClienteClienteClienteClienteClienteClienteClienteCliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A65C02&
      Height          =   360
      Left            =   1110
      TabIndex        =   9
      Top             =   4425
      Width           =   7875
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   8205
      TabIndex        =   7
      ToolTipText     =   "Sair do sistema"
      Top             =   4950
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   11683841
      BackColor       =   12648384
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "GMNU101.frx":112F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "GMNU101.frx":114B
      Caption         =   "Sair do sistema"
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdLogoff 
      Height          =   345
      Left            =   7245
      TabIndex        =   6
      ToolTipText     =   "Mudar de usuário"
      Top             =   4950
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   11683841
      BackColor       =   12648384
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "GMNU101.frx":1167
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "GMNU101.frx":1183
      Caption         =   "&Logoff"
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdLogon 
      Height          =   345
      Left            =   7245
      TabIndex        =   4
      ToolTipText     =   "Entrar no sistema"
      Top             =   4950
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   11683841
      BackColor       =   16765875
      PictureFrames   =   1
      Windowless      =   -1  'True
      MouseIcon       =   "GMNU101.frx":119F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "GMNU101.frx":11BB
      Caption         =   "&Logon"
      ButtonStyle     =   4
      PictureAlignment=   1
   End
End
Attribute VB_Name = "GMNU101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cont As Byte
Public SenhaInicial As String

Private Sub cmdLogoff_Click()
On Error GoTo Trata
    cmdLogoff.Enabled = False
    DoEvents
    If Util.Confirma("Tem certeza que deseja efetuar logoff de " & Usuario & "?") Then
        BdSis.FechaBanco
        TrocaTela
    

    End If
    cmdLogoff.Enabled = True
    DoEvents
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
    cmdLogoff.Enabled = True
    DoEvents
End Sub

Private Sub cmdLogon_Click()
    
    On Error GoTo Trata
    Sistema = ""
    Usuario = txtCodigo
    cmdLogon.Enabled = False
    Screen.MousePointer = 11
    cmdSair.SetFocus

    CarregaBar prgBar, (Usuario)
    CarregaTree treSis, (Usuario)
    lblUsuario = "para  " & txtNome
    lblSis = ""
    lblDescrSis = ""
    TrocaTela
    LimpaTudo
    cmdLogon.Enabled = True
    cmdLogoff.SetFocus
    Screen.MousePointer = 0
    
        'POVOACAO DA TAB_LOGON
    Dim Comando As Object
'    Set Comando = CreateObject("VSClass.VSComando")
'    Comando.Texto BdSis, "sp_logon", 4 '4=adCmdStoredProc
'    Comando.setarParametro "tlg_usuario_sistema", 201, 1, 20, Usuario  '201=adLongVarChar,1=adParamInput
'    Comando.Executa
'    Set Comando = Nothing
    
    DoEvents
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub CarregaBar(Bar As ProgressBar, User As String)
    Dim SQL As String
    Dim RS As Object
    
    SQL = "SELECT COUNT(*) FROM " & _
        " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & User & "'"
    
    Bar.Min = 0
    Bar.Value = 0
'    Bar.Visible = True
    
    DoEvents
    
    If BdSis.AbreTabela(SQL, RS) Then
        Bar.Max = RS(0) + 1
    Else
        Bar.Max = 1
    End If

End Sub

Private Sub cmdSair_Click()
On Error Resume Next
    Dim Comando As Object
    Set Comando = CreateObject("VSClass.VSComando")
    Comando.Texto BdSis, "sp_logoff", 4 '4=adCmdStoredProc
    Comando.setarParametro "tlg_usuario_sistema", 201, 1, 20, Usuario '201=adLongVarChar,1=adParamInput
    Comando.Executa
    Set Comando = Nothing
    
    Unload Me
End Sub

Private Sub TrocaTela()
    lblUsuario.Visible = Not lblUsuario.Visible
    picUsuario.Visible = Not picUsuario.Visible
    picTrocaSenha.Visible = Not picTrocaSenha.Visible
    picTree.Visible = Not picTree.Visible
    lblDescrSis.Visible = Not lblDescrSis.Visible
    lblSis.Visible = Not lblSis.Visible
    cmdLogon.Visible = Not cmdLogon.Visible
    cmdLogoff.Visible = Not cmdLogoff.Visible
    txtCodigo.Enabled = Not txtCodigo.Enabled
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If UCase(Me.ActiveControl.Name) <> "TRESIS" Then
            SendKeys "{TAB}"
        Else
            Call treSis_DblClick
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Trata
    SenhaInicial = Temp.PegaParametro(BdSis, "SENHA INICIAL")
    Me.Caption = Temp.PegaParametro(BdSis, "SISTEMA") & " - " & Temp.PegaParametro(BdSis, "FANTASIA")
    lblSistema = Temp.PegaParametro(BdSis, "SISTEMA")
    lblCliente = Temp.PegaParametro(BdSis, "CLIENTE")
    lblResp = Temp.PegaParametro(BdSis, "RESPONSAVEL")
    
    Util.CarregaFig imlMenu, App.Path, 4, 7, 16
    Screen.MousePointer = 0
    DoEvents
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub CarregaTree(lst As TreeView, User As String)
    On Error GoTo Trata
    Dim Fig As String
    Dim RS1 As Object
    Dim RS2 As Object
    Dim RS3 As Object
    Dim SQL As String
    lst.Nodes.Clear
    
    picProgress.Visible = True
    
    Fig = "VISUAL"
    lst.Nodes.Add , tvwFirst, "0VSIS", Temp.PegaParametro(BdSis, "SISTEMA"), Fig
    
    SQL = "SELECT TSI_COD_SISTEMA, TSI_NOME FROM TAB_SISTEMA WHERE " & _
    " TSI_COD_SISTEMA IN (SELECT DISTINCT TAU_TSI_COD_SISTEMA FROM " & _
    " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & User & "') " & _
    " ORDER BY TSI_NOME"
    
    If BdSis.AbreTabela(SQL, RS1) Then
        Do Until RS1.EOF
            Fig = "CLOSE"
            lst.Nodes.Add "0VSIS", tvwChild, "1" & RS1(0), RS1(1), Fig
            
            
            SQL = "SELECT TMO_COD_MODULO, TMO_NOME FROM TAB_MODULO WHERE " & _
            " TMO_COD_MODULO IN (SELECT DISTINCT TAU_TMO_COD_MODULO FROM " & _
            " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & _
            User & "' AND TAU_TSI_COD_SISTEMA = '" & RS1(0) & "')" & _
            " ORDER BY TMO_NOME"
            
            
            
            If BdSis.AbreTabela(SQL, RS2) Then
                Do Until RS2.EOF
                    Fig = "CLOSE"
                    lst.Nodes.Add "1" & CStr(RS1(0)), tvwChild, "2" & CStr(RS2(0)), CStr(RS2(1)), Fig
                    
                    
                    SQL = "SELECT TFO_COD_FORMULARIO, TFO_NOME FROM TAB_FORMULARIO WHERE " & _
                    " TFO_TMO_COD_MODULO " & BdSis.Concatena & " TFO_COD_FORMULARIO IN (SELECT DISTINCT TAU_TMO_COD_MODULO " & BdSis.Concatena & " TAU_TFO_COD_FORMULARIO FROM " & _
                    " TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & _
                     User & "' AND TAU_TMO_COD_MODULO = '" & RS2(0) & "')" & _
                     " ORDER BY TFO_NOME"

                    If BdSis.AbreTabela(SQL, RS3) Then
                        Do Until RS3.EOF
                            Fig = RS2(0) & RS3(0)
                            lst.Nodes.Add "2" & CStr(RS2(0)), tvwChild, "3" & CStr(RS2(0) & RS3(0)), CStr(RS3(1)), Fig
                            lst.Nodes.Item("3" & CStr(RS2(0) & RS3(0))).Tag = RS1(0)
                            RS3.MoveNext
                            
                            prgBar.Value = prgBar.Value + 1
                            DoEvents
                        Loop
                    End If
                    BdSis.FechaTabela RS3
                    
                    
                    RS2.MoveNext
                Loop
            End If
            BdSis.FechaTabela RS2
            RS1.MoveNext
        Loop
    End If
    BdSis.FechaTabela RS1
    lst.Nodes.Item("0VSIS").Expanded = True
    
    picProgress.Visible = False
    
    Exit Sub
Trata:
    If Err.Number = 35601 Then
        Fig = "NAOTEM"
        Resume
    ElseIf Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
    picProgress.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Trata
    If Not Util.Confirma("Deseja realmente finalizar o sistema?") Then
        Cancel = True
    Else
        BdSis.FechaBanco
        BdSis.FechaBanco
'        End
    End If
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub lblEmp_Click()

End Sub

Private Sub tmrSis_Timer()
     lblData = Format(Now, "Long Date")
End Sub

Private Sub treSis_Collapse(ByVal Node As MSComctlLib.Node)
    If treSis.Nodes(Node.Index).Image = "OPEN" Then treSis.Nodes(Node.Index).Image = "CLOSE"
End Sub

Private Sub treSis_DblClick()
On Error GoTo Trata
    
    If Mid(treSis.SelectedItem.Key, 1, 1) = "3" Then
        Screen.MousePointer = 11
        
        Call ChamaAplicacao(Mid(treSis.SelectedItem.Key, 2), Mid(treSis.SelectedItem.Parent.Parent.Key, 2), treSis.SelectedItem.Parent.Parent.Text, treSis.SelectedItem.Text)
        Screen.MousePointer = 0
    End If

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 91 Then
            Util.Informa "Módulo " & Sistema & " não encontrado."
        Else
            Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Sub treSis_Expand(ByVal Node As MSComctlLib.Node)
    If treSis.Nodes(Node.Index).Image = "CLOSE" Then treSis.Nodes(Node.Index).Image = "OPEN"
End Sub

Private Sub treSis_NodeClick(ByVal Node As MSComctlLib.Node)
    If Not Node Is Nothing Then
        lblSis = Node
        lblDescrSis = Descricao(Node.Key)
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    LimpaTudo
    cmdLogon.Enabled = False
End Sub

Private Sub LimpaTudo()
    txtCodigo = ""
    txtNome = ""
    txtSenha = ""
    txtNova = ""
    txtConfirmar = ""
End Sub

Private Sub txtCodigo_LostFocus()
    On Error GoTo Trata
    
    If UCase(Me.ActiveControl.Name) = "CMDSAIR" Then Exit Sub
    
    txtCodigo = Trim(txtCodigo)
    txtNome = Seguranca.ExisteUsuario(BdSis, txtCodigo)
    If txtNome <> "" Then
        Cont = 0
    Else
        Util.Informa "Usuário '" & txtCodigo & "' não Cadastrado."
        Cont = Cont + 1
        If Cont >= 3 Then
            Util.AVISA "Tentativas Esgotadas. O Sistema será Finalizado."
'            End
        End If
        txtCodigo.SetFocus
    End If
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtConfirmar_LostFocus()
    On Error GoTo Trata
    If txtNova <> txtConfirmar Then
        Util.AVISA "Confirmação de senha inválida."
        txtNova = ""
        txtNova.SetFocus
    Else
         AlteraSenha txtCodigo, txtNova
        Util.Informa "Senha Alterada com Segurança."
        cmdLogon.Enabled = True
        Confirmar False
        txtNova = ""
        cmdLogon.SetFocus
    End If
        
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtNova_GotFocus()
    txtNova = ""
    Confirmar False
End Sub

Private Sub txtNova_LostFocus()
    On Error GoTo Trata

    If UCase(Me.ActiveControl.Name) = "TXTSENHA" Then txtSenha.SetFocus: Exit Sub
    txtNova = Trim(txtNova)
    
    
    If txtNova <> "" Then
        If NovaValida(txtNova) Then
            Confirmar True
            txtConfirmar.SetFocus
            cmdLogon.Enabled = False
        Else
            Util.Informa "Nova senha inválida. Deve possuir de " & Temp.PegaParametro(BdSis, "SENHA TAMANHO MIN") & " a " & Temp.PegaParametro(BdSis, "SENHA TAMANHO MAX") & " caracteres alfa-numéricos."
            txtNova = ""
            txtNova.SetFocus
        End If
    End If
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha = ""
    txtNova = ""
    cmdLogon.Enabled = False
    Confirmar False
End Sub

Private Sub Confirmar(Valor As Boolean)
    txtConfirmar.Visible = Valor
    txtConfirmar = ""
End Sub

Private Sub Travar()
    Util.AVISA "Tentativas Esgotadas. O usuário '" & txtCodigo & "' será bloqueado."
    Call BdSis.AtualizaDados("tab_usuario", "0'", "tus_ativo", "tus_cod_usuario= '" & txtCodigo & "'")
End Sub

Private Sub txtSenha_LostFocus()
    On Error GoTo Trata
    
    If UCase(Me.ActiveControl.Name) = "CMDSAIR" Or _
    UCase(Me.ActiveControl.Name) = "TXTCODIGO" Then Exit Sub
    txtSenha = Trim(txtSenha)
    
    If Not SenhaValida(txtCodigo, txtSenha) Then
        Cont = Cont + 1
        Util.AVISA "Senha inválida."
        If Cont = 3 Then
            If Temp.PegaParametro(BdSis, "TRAVAR") = "SIM" Then Travar
            Util.AVISA "O Sistema será finalizado."
'            End
        End If
        txtSenha.SetFocus
    Else
        If DeveTrocarSenha(txtSenha) Then
            Util.Informa "Sua senha expirou. Troque sua senha."
        Else
            cmdLogon.Enabled = True
            Cont = 0
        End If
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Public Function DeveTrocarSenha(Senha As String) As Boolean
    If txtSenha = SenhaInicial Then
        DeveTrocarSenha = True
    End If
End Function

Public Function SenhaValida(User As String, Senha As String) As Boolean
    On Error GoTo Trata
    Dim RS As Object
    Dim SQL As String
    
    SQL = "SELECT TUS_SENHA FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & (User) & "'"
    If BdSis.AbreTabela(SQL, RS) Then
        If RS(0) = Seguranca.Criptografa(Senha) Then
            SenhaValida = True
            Exit Function
        End If
    End If
    BdSis.FechaTabela RS
    SenhaValida = False
    
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Sub AlteraSenha(User As String, Nova As String)
    On Error GoTo Trata
    Dim Valor As String
    Valor = BdSis.PreparaValor(Seguranca.Criptografa(Nova))
    BdSis.AtualizaDados "TAB_USUARIO", Valor, "TUS_SENHA", _
     "TUS_COD_USUARIO = '" & (User) & "'"
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Public Function NovaValida(Senha As String) As Boolean
    NovaValida = False
    Senha = Trim(Senha)
    If Len(Senha) >= CInt(Temp.PegaParametro(BdSis, "SENHA TAMANHO MIN")) And Len(Senha) <= CInt(Temp.PegaParametro(BdSis, "SENHA TAMANHO MAX")) Then
        If Senha <> SenhaInicial Then
            NovaValida = True
        End If
    End If
    
End Function

Public Function Descricao(Sistema As String) As String
    Dim RS As Object
    Dim Tipo As Byte: Tipo = Mid(Sistema, 1, 1)
    Dim SQL As String
    Select Case Tipo
        Case 0 'Nada
            Descricao = Temp.PegaParametro(BdSis, "DESCRICAO")
        Case 1 'Sistema
            SQL = "SELECT TSI_DESCR FROM TAB_SISTEMA WHERE TSI_COD_SISTEMA = '" & Mid(Sistema, 2) & "'"
        Case 2 'Módulo
            SQL = "SELECT TMO_DESCR FROM TAB_MODULO WHERE TMO_COD_MODULO = '" & Mid(Sistema, 2) & "'"
        Case 3 'Formulário
            SQL = "SELECT TFO_DESCR FROM TAB_FORMULARIO WHERE TFO_TMO_COD_MODULO " & BdSis.Concatena & " TFO_COD_FORMULARIO = '" & Mid(Sistema, 2) & "'"
    End Select
    If SQL <> "" Then
        If BdSis.AbreTabela(SQL, RS) Then
            Descricao = "" & RS(0)
        End If
        BdSis.FechaTabela RS
    End If
End Function

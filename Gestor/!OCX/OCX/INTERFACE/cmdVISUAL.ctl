VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl cmdVISUAL 
   BackColor       =   &H80000016&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   DefaultCancel   =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   1200
   ToolboxBitmap   =   "cmdVISUAL.ctx":0000
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -360
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":0312
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":201E
            Key             =   "adicionar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":25B8
            Key             =   "excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":42C4
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":485E
            Key             =   "salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":656A
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":6B04
            Key             =   "sair"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":9810
            Key             =   "limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cmdVISUAL.ctx":9DAA
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin Threed.SSCommand cmdBotao 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   8388608
      BackColor       =   14737632
      PictureMaskColor=   12583104
      PictureUseMask  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Rotulo"
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Shape shpBorda 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "cmdVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum eAcao
    Nenhum = 0
    Adicionar = 1
    Excluir = 2
    Salvar = 3
    Imprimir = 4
    Buscar = 5
    Limpar = 6
    Sair = 7
    Usuario = 8
    Cancelar = 9
End Enum
Public Event Click()
Private m_Acao As eAcao
Private m_Enabled As Boolean
Private m_CorBorda As OLE_COLOR
Private m_CorFrente As OLE_COLOR
Private m_CorFundo As OLE_COLOR
Private m_CorFoco As OLE_COLOR
Private m_Icone As IPictureDisp
Private m_Cancel As Boolean

Public Property Get Cancel() As Boolean
    On Error GoTo Trata

1   If bRegistrado Then Cancel = m_Cancel
    Exit Property
Trata:
End Property

Public Property Let Cancel(ByVal Value As Boolean)
    On Error GoTo Trata

1   If bRegistrado Then
2       m_Cancel = Value
3       cmdBotao.Cancel = m_Cancel
4       PropertyChanged "Cancel"
5   End If

    Exit Property
Trata:
End Property

Public Property Get Icone() As IPictureDisp
    On Error GoTo Trata

1   If bRegistrado Then Set Icone = m_Icone
    Exit Property
Trata:
End Property

Public Property Set Icone(ByVal Value As IPictureDisp)
    On Error GoTo Trata

1   If bRegistrado Then
2       Set m_Icone = Value
        '    ImageList1.ListImages.Remove "user"
        '    ImageList1.ListImages.Add , "user", m_Icone
        '    Dim img As ListImage

3       If m_Icone <> 0 Then
4           ImageList1.ListImages.Remove "user"
5           ImageList1.ListImages.Add , "user", m_Icone
6       End If

        'Set ImageList1.ListImages("user") = ImageList1.ListImages("temp")
        '    Set ImageList1.ListImages("user") = m_Icone
7       AplicarFigura
8       PropertyChanged "Icone"
9   End If
    Exit Property
Trata:
End Property

Public Property Get CorFoco() As OLE_COLOR
    On Error GoTo Trata

1   If bRegistrado Then CorFoco = m_CorFoco
    Exit Property
Trata:
End Property

Public Property Let CorFoco(ByVal Value As OLE_COLOR)
    On Error GoTo Trata
1   If bRegistrado Then
2       m_CorFoco = Value
3       PropertyChanged "CorFoco"
4   End If

    Exit Property
Trata:
End Property

Public Property Get corFundo() As OLE_COLOR
    On Error GoTo Trata
1   If bRegistrado Then corFundo = m_CorFundo
    Exit Property
Trata:
End Property

Public Property Let corFundo(ByVal Value As OLE_COLOR)
    On Error GoTo Trata
1   If bRegistrado Then
2       m_CorFundo = Value
3       cmdBotao.BackColor = m_CorFundo
4       UserControl.BackColor = m_CorFundo
5       PropertyChanged "CorFundo"
6   End If

    Exit Property
Trata:
End Property

Public Property Get CorFrente() As OLE_COLOR
    On Error GoTo Trata
1   If bRegistrado Then CorFrente = m_CorFrente
    Exit Property
Trata:
End Property

Public Property Let CorFrente(ByVal Value As OLE_COLOR)
    On Error GoTo Trata
1   If bRegistrado Then
2       m_CorFrente = Value
3       cmdBotao.ForeColor = m_CorFrente
4       PropertyChanged "CorFrente"
5   End If
    Exit Property
Trata:
End Property

Public Property Get CorBorda() As OLE_COLOR
    On Error GoTo Trata
1   If bRegistrado Then CorBorda = m_CorBorda
    Exit Property
Trata:
End Property

Public Property Let CorBorda(ByVal Value As OLE_COLOR)
    On Error GoTo Trata
1   If bRegistrado Then
2       m_CorBorda = Value
3       shpBorda.BorderColor = m_CorBorda
4       PropertyChanged "CorBorda"
5   End If
    Exit Property
Trata:
End Property

Public Property Get Enabled() As Boolean
    On Error GoTo Trata
1   If bRegistrado Then Enabled = m_Enabled
    Exit Property
Trata:
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    On Error GoTo Trata
1   If bRegistrado Then
2       m_Enabled = Value
3       cmdBotao.Enabled = Value
4       cmdBotao.TabStop = Value
5       PropertyChanged "Enabled"
6   End If
    Exit Property
Trata:
End Property

Property Get Acao() As eAcao
    On Error GoTo Trata

1   If bRegistrado Then
2       Acao = m_Acao
3       AplicarFigura
4   End If

    Exit Property
Trata:
End Property

Property Let Acao(ByVal Value As eAcao)
    On Error GoTo Trata
1   If bRegistrado Then
2       m_Acao = Value
3       PropertyChanged "Acao"
4       AplicarFigura
5   End If

    Exit Property
Trata:
End Property

Private Sub cmdBotao_Click()
    On Error GoTo Trata

1   If bRegistrado Then RaiseEvent Click
    Exit Sub
Trata:
End Sub

Private Sub cmdBotao_GotFocus()
    On Error GoTo Trata
1   If bRegistrado Then cmdBotao.BackColor = m_CorFoco
    Exit Sub
Trata:
End Sub

Private Sub cmdBotao_LostFocus()
    On Error GoTo Trata
1   If bRegistrado Then cmdBotao.BackColor = m_CorFundo
    Exit Sub
Trata:
End Sub

Private Sub cmdBotao_MouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error GoTo Trata
1   If bRegistrado Then cmdBotao.BackColor = m_CorFoco
    Exit Sub
Trata:
End Sub

Private Sub cmdBotao_MouseExit(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error GoTo Trata
1   If bRegistrado Then cmdBotao.BackColor = m_CorFundo
    Exit Sub
Trata:
End Sub

Private Sub UserControl_Initialize()
    On Error GoTo Trata
1   ValidaComponente "INTERFACE"

'2   If bRegistrado Then
3       Set Util = New VSUtil
'4   End If

    Exit Sub
Trata:
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Trata
1   If bRegistrado Then
2       shpBorda.Width = Width
3       shpBorda.Height = Height
4       cmdBotao.Width = Width - 60
5       cmdBotao.Height = Height - 60
6   End If

    Exit Sub
Trata:
End Sub

Private Sub UserControl_InitProperties()
    On Error GoTo Trata

1   If bRegistrado Then
2       Caption = "Rotulo"
3       Acao = Nenhum
4       Set Icone = LoadPicture("")
5       Enabled = True
6       CorBorda = Ambient.ForeColor
7       CorFrente = Ambient.ForeColor
8       corFundo = Ambient.BackColor
9       CorFoco = &HF5F5F5
10      Cancel = False
11  End If

    Exit Sub
Trata:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next

1   If bRegistrado Then
2       Caption = PropBag.ReadProperty("Caption", "Rotulo")
3       Acao = PropBag.ReadProperty("Acao", Nenhum)
4       Set Icone = PropBag.ReadProperty("Icone", LoadPicture(""))
5       Enabled = PropBag.ReadProperty("Enabled", True)
6       CorBorda = PropBag.ReadProperty("CorBorda", Ambient.ForeColor)
7       CorFrente = PropBag.ReadProperty("CorFrente", Ambient.ForeColor)
8       corFundo = PropBag.ReadProperty("CorFundo", Ambient.BackColor)
9       CorFoco = PropBag.ReadProperty("CorFoco", &HF5F5F5)
10      Cancel = PropBag.ReadProperty("Cancel", False)
11  End If
    Exit Sub
Trata:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo Trata

1   If bRegistrado Then
2       Call PropBag.WriteProperty("Caption", cmdBotao.Caption, "Rotulo")
3       Call PropBag.WriteProperty("Acao", m_Acao, Nenhum)
4       AplicarFigura
5       Call PropBag.WriteProperty("Enabled", m_Enabled, True)
6       Call PropBag.WriteProperty("CorBorda", m_CorBorda, Ambient.ForeColor)
7       Call PropBag.WriteProperty("CorFrente", m_CorFrente, Ambient.ForeColor)
8       Call PropBag.WriteProperty("CorFundo", m_CorFundo, Ambient.BackColor)
9       Call PropBag.WriteProperty("CorFoco", m_CorFoco, &HF5F5F5)
10      Call PropBag.WriteProperty("Icone", m_Icone, LoadPicture(""))
11      Call PropBag.WriteProperty("Cancel", m_Cancel, False)
12  End If

    Exit Sub
Trata:
End Sub

Public Property Get Caption() As String
    On Error GoTo Trata
1   If bRegistrado Then Caption = cmdBotao.Caption
    Exit Property
Trata:
End Property

Public Property Let Caption(ByVal vnewvalue As String)
    On Error GoTo Trata

1   If bRegistrado Then cmdBotao.Caption = vnewvalue
    Exit Property
Trata:
End Property

Public Sub AplicarFigura()
    On Error GoTo Trata

1   If bRegistrado Then
2       Dim figura As String

3       Select Case m_Acao

            Case Nenhum: figura = ""

4           Case Adicionar: figura = "adicionar"

5           Case Excluir: figura = "excluir"

6           Case Salvar: figura = "salvar"

7           Case Imprimir: figura = "imprimir"

8           Case Buscar: figura = "buscar"

9           Case Limpar: figura = "limpar"

10          Case Sair: figura = "sair"

11          Case Usuario: figura = "user"

12          Case Cancelar: figura = "cancelar"
13      End Select

14      If figura = "" Then
            '        If Icone Is Nothing Then
15          cmdBotao.Picture = LoadPicture("")
16          cmdBotao.Alignment = ssCenterMiddle
            '         Else
            '            cmdBotao.Picture = ImageList1.ListImages("user").ExtractIcon
            '            cmdBotao.Alignment = ssRightMiddle
            '         End If

17      Else
18          cmdBotao.Picture = ImageList1.ListImages(figura).ExtractIcon
19          cmdBotao.Alignment = ssRightMiddle
20      End If

21  End If

    Exit Sub
Trata:
End Sub


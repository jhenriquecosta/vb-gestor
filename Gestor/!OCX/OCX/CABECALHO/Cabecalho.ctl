VERSION 5.00
Begin VB.UserControl ctlCabecalho 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   765
   ScaleWidth      =   9510
   ToolboxBitmap   =   "Cabecalho.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1620
      Top             =   1200
   End
   Begin VB.Image imgIcone 
      Height          =   750
      Left            =   30
      Picture         =   "Cabecalho.ctx":0312
      Stretch         =   -1  'True
      Top             =   30
      Width           =   900
   End
   Begin VB.Label lblFormulario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Formulário"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1050
      TabIndex        =   2
      Top             =   360
      Width           =   3150
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   8835
      TabIndex        =   1
      Top             =   30
      Width           =   630
   End
   Begin VB.Label lblSistema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1080
      TabIndex        =   0
      Top             =   30
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   -30
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9570
   End
End
Attribute VB_Name = "ctlCabecalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Timer1_Timer()
    If bRegistrado Then
        lblHora = Time
    End If
End Sub
Public Function Exibe(Sistema As String, Formulario As String, Optional Icone As String) As Boolean
    On Error GoTo erro
    
    If bRegistrado Then
        lblSistema = Sistema
        lblFormulario = Formulario
        If Icone <> "" Then
            imgIcone.Picture = LoadPicture(Icone)
        End If
        
        Exibe = True
    End If
    Exit Function
erro:
End Function

Private Sub UserControl_Initialize()
    ValidaComponente "CABECALHO"
End Sub

Private Sub UserControl_InitProperties()
    If bRegistrado Then
        BackColor = &HFFFFFF
        lblFormulario.ForeColor = &H800000
        Shape1.BackColor = &HE0E0E0
        lblSistema.ForeColor = &H808080
        lblHora.ForeColor = &H808080
        Sistema = "Sistema"
        Formulario = "Formulario"
    End If
End Sub

Private Sub UserControl_Resize()
    If bRegistrado Then
        Shape1.Width = Width
        lblHora.Left = Shape1.Width - lblHora.Width - 100
    End If
End Sub

Public Property Get CorFundo() As OLE_COLOR
    If bRegistrado Then CorFundo = BackColor
End Property

Public Property Let CorFundo(ByVal vNewValue As OLE_COLOR)
    If bRegistrado Then BackColor = vNewValue
End Property

Public Property Get CorFrente() As OLE_COLOR
    If bRegistrado Then CorFrente = lblFormulario.ForeColor
End Property

Public Property Let CorFrente(ByVal vNewValue As OLE_COLOR)
    If bRegistrado Then lblFormulario.ForeColor = vNewValue
End Property
Public Property Get CorFundoFaixa() As OLE_COLOR
    If bRegistrado Then CorFundoFaixa = Shape1.BackColor
End Property

Public Property Let CorFundoFaixa(ByVal vNewValue As OLE_COLOR)
    If bRegistrado Then Shape1.BackColor = vNewValue
End Property
Public Property Get CorFrenteFaixa() As OLE_COLOR
    If bRegistrado Then CorFrenteFaixa = lblSistema.ForeColor
End Property

Public Property Let CorFrenteFaixa(ByVal vNewValue As OLE_COLOR)
    If bRegistrado Then
        lblSistema.ForeColor = vNewValue
        lblHora.ForeColor = vNewValue
    End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If bRegistrado Then
        BackColor = PropBag.ReadProperty("CorFundo", &HFFFFFF)
        lblFormulario.ForeColor = PropBag.ReadProperty("CorFrente", &H800000)
        Shape1.BackColor = PropBag.ReadProperty("CorFundoFaixa", &HE0E0E0)
        lblSistema.ForeColor = PropBag.ReadProperty("CorFrenteFaixa", &H808080)
        lblHora.ForeColor = PropBag.ReadProperty("CorFrenteFaixa", &H808080)
        lblSistema = PropBag.ReadProperty("Sistema", "Sistema")
        lblFormulario = PropBag.ReadProperty("Formulario", "Formulario")
    End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Call PropBag.WriteProperty("CorFundo", BackColor, &HFFFFFF)
        Call PropBag.WriteProperty("CorFrente", lblFormulario.ForeColor, &H800000)
        Call PropBag.WriteProperty("CorFundoFaixa", Shape1.BackColor, &HE0E0E0)
        Call PropBag.WriteProperty("CorFrenteFaixa", lblSistema.ForeColor, &H808080)
        Call PropBag.WriteProperty("Sistema", lblSistema.Caption, "Sistema")
        Call PropBag.WriteProperty("Formulario", lblFormulario.Caption, "Formulario")
    End If
End Sub

Public Property Get Sistema() As String
    If bRegistrado Then Sistema = lblSistema
End Property

Public Property Let Sistema(ByVal vNewValue As String)
    If bRegistrado Then
        lblSistema = vNewValue
        Exibe lblSistema, lblFormulario
    End If
End Property

Public Property Get Formulario() As String
    If bRegistrado Then Formulario = lblFormulario
End Property

Public Property Let Formulario(ByVal vNewValue As String)
    If bRegistrado Then
        lblFormulario = vNewValue
        Exibe lblSistema, lblFormulario
    End If
End Property

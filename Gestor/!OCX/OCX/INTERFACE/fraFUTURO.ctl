VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.UserControl fraFUTURO 
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   4800
   ToolboxBitmap   =   "fraFUTURO.ctx":0000
   Begin Threed.SSCommand cmdFechar 
      Height          =   225
      Left            =   4500
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   -2147483647
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "+"
      ButtonStyle     =   4
   End
   Begin VB.Label lblCabecalho 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cabecalho"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3750
   End
   Begin VB.Label lblDescricao 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descricao"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3765
   End
   Begin VB.Image imgIcone 
      Height          =   480
      Left            =   120
      Top             =   90
      Width           =   480
   End
   Begin VB.Shape shpCabecalho 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   645
      Left            =   30
      Top             =   30
      Width           =   4755
   End
   Begin VB.Shape shpCorpo 
      BorderColor     =   &H80000001&
      Height          =   1215
      Left            =   30
      Top             =   630
      Width           =   4755
   End
End
Attribute VB_Name = "fraFUTURO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Altura As Single, m_AlturaAutomatica As Boolean
Private m_Status As Boolean
Public Event mudancaStatus()

Public Property Get Status() As stat
Attribute Status.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then
        If m_Status Then
            Status = staFechado
        Else
            Status = staAberto
        End If
    End If
End Property

Public Property Let Status(ByVal vnewvalue As stat)
    If bRegistrado Then
        If Ocultavel Then
            If vnewvalue = staAberto Then
                m_Status = True
            ElseIf vnewvalue = staFechado Then
                m_Status = False
            End If
            cmdFechar_Click
        End If
    End If
End Property

Public Property Let Enabled(Valor As Boolean)
    Dim Controle As Object
    
    If bRegistrado Then
        UserControl.Enabled = Valor
        For Each Controle In UserControl.ContainedControls
            Controle.Enabled = Valor
        Next
    End If
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Enabled = UserControl.Enabled
End Property

Public Property Get alturaAutomatica() As Boolean
Attribute alturaAutomatica.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then alturaAutomatica = m_AlturaAutomatica
End Property

Public Property Let alturaAutomatica(ByVal Value As Boolean)
    If bRegistrado Then m_AlturaAutomatica = Value
End Property
Public Property Get Altura() As Single
Attribute Altura.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Altura = m_Altura
End Property

Public Property Let Altura(ByVal Value As Single)
    If bRegistrado Then m_Altura = Value
End Property
Public Property Get Ocultavel() As Boolean
Attribute Ocultavel.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Ocultavel = cmdFechar.Visible
End Property

Public Property Let Ocultavel(Value As Boolean)
    If bRegistrado Then cmdFechar.Visible = Value
End Property

Public Property Get Alinhamento() As AlignmentConstants
Attribute Alinhamento.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Alinhamento = lblCabecalho.Alignment
End Property

Public Property Let Alinhamento(Value As AlignmentConstants)
    If bRegistrado Then
        lblCabecalho.Alignment = Value
        lblDescricao.Alignment = Value
    End If
End Property

Public Property Get Icone() As IPictureDisp
Attribute Icone.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Set Icone = imgIcone.Picture
End Property

Public Property Set Icone(ByVal Value As IPictureDisp)
    If bRegistrado Then
        imgIcone.Picture = Value
        If Value Is Nothing Then
            lblCabecalho.Left = imgIcone.Left
            lblDescricao.Left = imgIcone.Left
        Else
            If Value > 0 Then
                lblCabecalho.Left = imgIcone.Left + imgIcone.Width + 30
                lblDescricao.Left = imgIcone.Left + imgIcone.Width + 30
            Else
                lblCabecalho.Left = imgIcone.Left
                lblDescricao.Left = imgIcone.Left
            End If
        End If
        PropertyChanged "Icone"
    End If
End Property

Public Property Get corTexto() As OLE_COLOR
Attribute corTexto.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then corTexto = lblCabecalho.ForeColor
End Property

Public Property Let corTexto(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        lblCabecalho.ForeColor = Value
        lblDescricao.ForeColor = lblCabecalho.ForeColor
        PropertyChanged "corTexto"
    End If
End Property

Public Property Get corFundo() As OLE_COLOR
Attribute corFundo.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then corFundo = shpCabecalho.FillColor
End Property

Public Property Let corFundo(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        shpCabecalho.FillColor = Value
        cmdFechar.BackColor = shpCabecalho.FillColor
        PropertyChanged "corFundo"
    End If
End Property

Public Property Get corFaixa() As OLE_COLOR
Attribute corFaixa.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then corFaixa = shpCabecalho.BorderColor
End Property

Public Property Let corFaixa(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        shpCabecalho.BorderColor = Value
        shpCorpo.BorderColor = shpCabecalho.BorderColor
        cmdFechar.ForeColor = shpCabecalho.BorderColor
        PropertyChanged "corFaixa"
    End If
End Property

Public Property Get Descricao() As String
Attribute Descricao.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Descricao = lblDescricao
End Property

Public Property Let Descricao(ByVal Value As String)
    If bRegistrado Then
        lblDescricao = Value
        PropertyChanged "Descricao"
    End If
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Futuro"
    If bRegistrado Then Caption = lblCabecalho
End Property

Public Property Let Caption(ByVal Value As String)
    If bRegistrado Then
        lblCabecalho = Value
        PropertyChanged "Caption"
    End If
End Property

Private Sub cmdFechar_Click()
    If bRegistrado Then
        If Ocultavel Then
            If m_Status Then
                cmdFechar.Caption = "-"
                Height = m_Altura
            Else
                cmdFechar.Caption = "+"
                Height = shpCabecalho.Height + 30
            End If
            m_Status = Not m_Status
            RaiseEvent mudancaStatus
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    ValidaComponente "INTERFACE"
End Sub

Private Sub UserControl_InitProperties()
    If bRegistrado Then
        Caption = "Cabecalho"
        Descricao = "Descricao"
        corFaixa = &H80000002 'ActiveTitleBar
        corFundo = &H80000005 'WindowBackground
        corTexto = &H80000012 'ButtonText
        Ocultavel = True
        Altura = 1905
        shpCorpo.Top = shpCabecalho.Height - 30
        Enabled = True
        Status = staAberto
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim aux As stat
    
    If bRegistrado Then
        aux = PropBag.ReadProperty("Status", staAberto)
        If aux = staAberto Then
            m_Status = False
            cmdFechar.Caption = "-"
        Else
            m_Status = True
            cmdFechar.Caption = "+"
        End If
        
        Caption = PropBag.ReadProperty("Caption", "Cabecalho")
        Descricao = PropBag.ReadProperty("Descricao", "Descricao")
        corFaixa = PropBag.ReadProperty("corFaixa", &H80000002)
        corFundo = PropBag.ReadProperty("corFundo", &H80000005)
        corTexto = PropBag.ReadProperty("corTexto", &H80000012)
        Set Icone = PropBag.ReadProperty("Icone", imgIcone.Picture)
        Alinhamento = PropBag.ReadProperty("Alinhamento", vbLeftJustify)
        Ocultavel = PropBag.ReadProperty("Ocultavel", True)
        Altura = PropBag.ReadProperty("Altura", 1890)
        Enabled = PropBag.ReadProperty("Enabled", True)
    End If
End Sub

Private Sub UserControl_Resize()
    If bRegistrado Then
        If alturaAutomatica Then Altura = Height
        cmdFechar.Left = Width - 300
        shpCorpo.Height = IIf(Height - shpCabecalho.Height > 0, Height - shpCabecalho.Height, 0)
        shpCabecalho.Width = Width - 30
        shpCorpo.Width = Width - 30
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim aux As stat

    If bRegistrado Then
        If m_Status Then
            aux = staFechado
            cmdFechar.Caption = "+"
        Else
            aux = staAberto
            cmdFechar.Caption = "-"
        End If
    
        Call PropBag.WriteProperty("Status", aux, staAberto)
        Call PropBag.WriteProperty("Caption", lblCabecalho.Caption, "Cabecalho")
        Call PropBag.WriteProperty("Descricao", lblDescricao.Caption, "Descricao")
        Call PropBag.WriteProperty("corFaixa", shpCabecalho.BorderColor, &H80000002)
        Call PropBag.WriteProperty("corFundo", shpCabecalho.FillColor, &H80000005)
        Call PropBag.WriteProperty("corTexto", lblCabecalho.ForeColor, &H80000012)
        Call PropBag.WriteProperty("Icone", imgIcone.Picture, "")
        Call PropBag.WriteProperty("Alinhamento", lblCabecalho.Alignment, vbLeftJustify)
        Call PropBag.WriteProperty("Ocultavel", cmdFechar.Visible, True)
        Call PropBag.WriteProperty("Altura", m_Altura, 1890)
        Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    End If
End Sub

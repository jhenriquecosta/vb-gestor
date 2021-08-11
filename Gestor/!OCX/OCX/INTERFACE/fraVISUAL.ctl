VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.UserControl fraVISUAL 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   1890
   ScaleWidth      =   3255
   ToolboxBitmap   =   "fraVISUAL.ctx":0000
   Begin Threed.SSCommand cmdFechar 
      Height          =   195
      Left            =   2970
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   344
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ButtonStyle     =   4
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   1875
      Left            =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblRotulo 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Frame"
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
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3195
   End
End
Attribute VB_Name = "fraVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'============================================
'PROPÓSITO fraVISUAL
'   Criar um frame no padrão de interface dos componentes
'da Visual Tecnologia, com recurso de roll-up
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 13:07
'============================================
'DEPENDÊNCIAS RUN-TIME
'   SSA3D30.ocx
'============================================
'METODOS
'   MudancaStatus
'   Status
'   Caption
'   CorTexto
'   CorFaixa
'   CorFundo
'   Alinhamento
'   Ocultavel
'============================================
'CONTROLES
'   cmdFechar
'   Shape1
'   lblRotulo
'============================================
Option Explicit
Private m_Altura As Single, m_AlturaAutomatica As Boolean
Private m_Status As Boolean, m_Ocultavel As Boolean
Private m_InverterCores As Boolean
Public Enum stat
    staAberto
    staFechado
End Enum
Public Event mudancaStatus()
Attribute mudancaStatus.VB_Description = "Disparado quando ocorre mudanca no status (aberto ou fechado) do frame"


Public Property Let Enabled(Valor As Boolean)
    If bRegistrado Then
        UserControl.Enabled = Valor
        
        Dim Controle As Object
        For Each Controle In UserControl.ContainedControls
            Controle.Enabled = Valor
        Next
    End If
End Property

Public Property Get Enabled() As Boolean
    If bRegistrado Then Enabled = UserControl.Enabled
End Property

'============================================
'PROPOSITO cmdFechar_Click
'   Fechar (ou abrir) ou o frame
'============================================
'Autor: Sergio Queiroz
'Data: 12/03/01 20:30
'============================================
Private Sub cmdFechar_Click()
    If bRegistrado Then
        Dim auxCorAnterior As OLE_COLOR
        
        If InverterCores Then
    
            auxCorAnterior = corFaixa
            corFaixa = corTexto
            corTexto = auxCorAnterior
    
        End If
        
        If m_Status Then
    
            cmdFechar.Caption = "-"
            Height = m_Altura
    
        Else
    
            cmdFechar.Caption = "+"
            Height = lblRotulo.Height + 50
    
        End If
    
        m_Status = Not m_Status
        RaiseEvent mudancaStatus
    End If

End Sub

'============================================
'PROPOSITO lblRotulo_DblClick
'   Fornecer ao usuário uma opção para a abertura ou
'fechamento do frame
'============================================
'Autor: Sergio Queiroz
'Data: 12/03/01 20:32
'============================================
Private Sub lblRotulo_DblClick()
    If bRegistrado Then
        If Ocultavel Then cmdFechar_Click
    End If

End Sub

Private Sub UserControl_Initialize()
    ValidaComponente "INTERFACE"
'    If bRegistrado Then
        Set Util = New VSUtil
'    End If
End Sub

'============================================
'PROPOSITO UserControl_InitProperties
'   Iniciar os valores padrões das propriedades
'============================================
'Autor: Sergio Queiroz
'Data: 12/03/01 20:37
'============================================
Private Sub UserControl_InitProperties()
    If bRegistrado Then
        corFundo = &HFFFFFF
        corFaixa = &HC0C0C0
        corTexto = &H0&
        Caption = "Grupo"
        Borda = vbBSSolid
        Ocultavel = True
        Status = staAberto
        cmdFechar.Caption = "-"
        Altura = 1905
        InverterCores = False
        Enabled = True
    End If

End Sub

'============================================
'PROPÓSITO UserControl_ReadProperties
'   Atualizar os componentes do controle com os valores das
'propriedades
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 13:06
'============================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Dim aux As stat
        aux = PropBag.ReadProperty("Status", staAberto)
    
        If aux = staAberto Then
    
            m_Status = False
            cmdFechar.Caption = "-"
    
        Else
    
            m_Status = True
            cmdFechar.Caption = "+"
    
        End If
    
        Altura = PropBag.ReadProperty("Altura", 1890)
        Caption = PropBag.ReadProperty("Caption", " Grupo")
        corTexto = PropBag.ReadProperty("CorTexto", &HE0E0E0)
        corFaixa = PropBag.ReadProperty("CorFaixa", &H80000010)
        corFundo = PropBag.ReadProperty("CorFundo", &HE0E0E0)
        Alinhamento = PropBag.ReadProperty("Alinhamento", vbLeftJustify)
        Ocultavel = PropBag.ReadProperty("Ocultavel", True)
        alturaAutomatica = PropBag.ReadProperty("AlturaAutomatica", False)
        InverterCores = PropBag.ReadProperty("InverterCores", False)
        Borda = PropBag.ReadProperty("Borda", vbBSSolid)
        Enabled = PropBag.ReadProperty("Enabled", True)
    End If
End Sub

'============================================
'PROPOSITO UserControl_Resize
'   Redimensionar o label e o contorno do frame e, ainda,
' reposicionar o botão de fechar
'============================================
'Autor: Sergio Queiroz
'Data: 12/03/01 20:35
'============================================
Private Sub UserControl_Resize()
    If bRegistrado Then
        Shape1.Height = Height
        Shape1.Width = Width
        lblRotulo.Width = Width - 60
    
        If cmdFechar.Caption = "+" Then
    
            lblRotulo.Height = Height - 60
    
        Else
    
            lblRotulo.Height = 255
    
        End If
    
        cmdFechar.Left = Width - 285
    
        If alturaAutomatica Then Altura = Height
    End If

End Sub

'============================================
'PROPÓSITO Ocultavel
'   Permitir ao programador o controle sobre a capacidade de
'abrir e fechar o frame
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 12:56
'============================================
Public Property Get Ocultavel() As Boolean
Attribute Ocultavel.VB_Description = "Permitir ao programador o controle sobre a capacidade de abrir e fechar o frame"
Attribute Ocultavel.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then
        Ocultavel = m_Ocultavel
        cmdFechar.Visible = m_Ocultavel
    End If

End Property

Public Property Let Ocultavel(Valor As Boolean)
    If bRegistrado Then
        m_Ocultavel = Valor
        cmdFechar.Visible = m_Ocultavel
    End If

End Property

'============================================
'PROPÓSITO Alinhamento
'   Definir o alinhamento horizontal do titulo
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 12:57
'============================================
Public Property Get Alinhamento() As AlignmentConstants
Attribute Alinhamento.VB_Description = "Definir o alinhamento horizontal do titulo"
Attribute Alinhamento.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then Alinhamento = lblRotulo.Alignment

End Property

Public Property Let Alinhamento(Valor As AlignmentConstants)
    If bRegistrado Then
        lblRotulo.Alignment = Valor
        lblRotulo = Trim$(lblRotulo)
    
        Select Case Valor
    
            Case vbRightJustify
    
                If Ocultavel Then
    
                    lblRotulo = lblRotulo & Space$(7)
    
                Else
    
                    lblRotulo = lblRotulo & Space$(1)
    
                End If
    
            Case vbLeftJustify
                lblRotulo = Space$(1) & lblRotulo
    
        End Select
    End If

End Property

'============================================
'PROPÓSITO CorFundo
'   Especificar a cor do frame
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 12:58
'============================================
Public Property Get corFundo() As OLE_COLOR
Attribute corFundo.VB_Description = "Especificar a cor do frame"
Attribute corFundo.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then corFundo = BackColor

End Property

Public Property Let corFundo(ByVal vnewvalue As OLE_COLOR)
    If bRegistrado Then BackColor = vnewvalue

End Property

'============================================
'PROPÓSITO CorFaixa
'   Especificar a cor do cabecalho do frame
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 12:59
'============================================
Public Property Get corFaixa() As OLE_COLOR
Attribute corFaixa.VB_Description = "Especificar a cor do cabecalho do frame"
Attribute corFaixa.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then corFaixa = lblRotulo.BackColor

End Property

Public Property Let corFaixa(ByVal vnewvalue As OLE_COLOR)
    If bRegistrado Then
        lblRotulo.BackColor = vnewvalue
        cmdFechar.BackColor = vnewvalue
        Shape1.BorderColor = vnewvalue
    End If

End Property

'============================================
'PROPÓSITO CorTexto
'   Especificar a cor do titulo
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 12:59
'============================================
Public Property Get corTexto() As OLE_COLOR
Attribute corTexto.VB_Description = "Especificar a cor do titulo"
Attribute corTexto.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then corTexto = lblRotulo.ForeColor

End Property

Public Property Let corTexto(ByVal vnewvalue As OLE_COLOR)
    If bRegistrado Then
        lblRotulo.ForeColor = vnewvalue
        cmdFechar.ForeColor = vnewvalue
    End If

End Property

'============================================
'PROPÓSITO Caption
'   Titulo do frame
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 13:00
'============================================
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Titulo do frame"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then Caption = Trim$(lblRotulo.Caption)

End Property

Public Property Let Caption(ByVal vnewvalue As String)
    If bRegistrado Then lblRotulo.Caption = " " & vnewvalue

End Property

'============================================
'PROPÓSITO Status
'   Informar a situacao (aberto ou fechado) em que se encontra
' o frame
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 13:00
'============================================
Public Property Get Status() As stat
Attribute Status.VB_Description = "Informar a situacao (aberto ou fechado) em que se encontra o frame"
Attribute Status.VB_ProcData.VB_Invoke_Property = ";Visual"
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

'============================================
'PROPÓSITO UserControl_WriteProperties
'   Escrever nas propriedades os valores dos componentes
'============================================
'Autor: Sergio Queiroz
'Data: 08/12/2001 13:18
'============================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Dim aux As stat
    
        If m_Status Then
    
            aux = staFechado
            cmdFechar.Caption = "+"
    
        Else
    
            aux = staAberto
            cmdFechar.Caption = "-"
    
        End If
    
        Call PropBag.WriteProperty("Status", aux, staAberto)
        
        Call PropBag.WriteProperty("Altura", m_Altura, 1890)
        Call PropBag.WriteProperty("Caption", lblRotulo.Caption, " Grupo")
        Call PropBag.WriteProperty("CorTexto", lblRotulo.ForeColor, &HE0E0E0)
        Call PropBag.WriteProperty("CorFaixa", lblRotulo.BackColor, &H80000010)
        Call PropBag.WriteProperty("CorFundo", BackColor, &HE0E0E0)
        Call PropBag.WriteProperty("Alinhamento", lblRotulo.Alignment, vbLeftJustify)
        Call PropBag.WriteProperty("Ocultavel", m_Ocultavel, True)
        cmdFechar.Visible = m_Ocultavel
        Call PropBag.WriteProperty("AlturaAutomatica", m_AlturaAutomatica, False)
        Call PropBag.WriteProperty("InverterCores", m_InverterCores, False)
        Call PropBag.WriteProperty("Borda", Shape1.BorderStyle, vbBSSolid)
        Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    End If

End Sub

Public Property Get Borda() As BorderStyleConstants
Attribute Borda.VB_Description = "Define se o controle possui ou não borda"
Attribute Borda.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then Borda = Shape1.BorderStyle

End Property

Public Property Let Borda(ByVal vnewvalue As BorderStyleConstants)
    If bRegistrado Then Shape1.BorderStyle = vnewvalue

End Property

Public Property Get Altura() As Single
Attribute Altura.VB_Description = "Altura que o controle vai ter quando aberto"
Attribute Altura.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then Altura = m_Altura

End Property

Public Property Let Altura(ByVal vnewvalue As Single)
    If bRegistrado Then m_Altura = vnewvalue

End Property

Public Property Get alturaAutomatica() As Boolean
Attribute alturaAutomatica.VB_Description = "Define a propriedade Altura automaticamente quando dimensionado o controle"
Attribute alturaAutomatica.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then alturaAutomatica = m_AlturaAutomatica

End Property

Public Property Let alturaAutomatica(ByVal vnewvalue As Boolean)
    If bRegistrado Then m_AlturaAutomatica = vnewvalue

End Property

Public Property Get InverterCores() As Boolean
Attribute InverterCores.VB_Description = "Define o comportamento das cores do título"
Attribute InverterCores.VB_ProcData.VB_Invoke_Property = ";Visual"
    If bRegistrado Then InverterCores = m_InverterCores

End Property

Public Property Let InverterCores(ByVal vnewvalue As Boolean)
    If bRegistrado Then m_InverterCores = vnewvalue

End Property


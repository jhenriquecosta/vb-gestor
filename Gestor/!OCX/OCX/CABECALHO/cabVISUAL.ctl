VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl cabVISUAL 
   Alignable       =   -1  'True
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   780
   ScaleWidth      =   4800
   ToolboxBitmap   =   "cabVISUAL.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   150
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2130
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cabVISUAL.ctx":0312
            Key             =   "visual"
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame fraCabecalho 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   1138
      _Version        =   196610
      BackColor       =   16777215
      Begin VB.Label lblHora 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   4665
         TabIndex        =   3
         Top             =   30
         Width           =   45
      End
      Begin VB.Label lblDescricao 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         TabIndex        =   2
         Top             =   270
         Width           =   4005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFormulario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formulário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   690
         TabIndex        =   1
         Top             =   60
         Width           =   915
      End
      Begin VB.Image imgIcone 
         Height          =   480
         Left            =   120
         Picture         =   "cabVISUAL.ctx":062C
         Top             =   60
         Width           =   480
      End
   End
End
Attribute VB_Name = "cabVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Formulario As String
Private m_Descricao As String
Private m_Icone As IPictureDisp
Private m_Codigo As String

Public Property Get Codigo() As String
    If bRegistrado Then Codigo = m_Codigo
End Property

Public Property Let Codigo(ByVal Value As String)
    If bRegistrado Then
        m_Codigo = Value
        PropertyChanged "Codigo"
    End If
End Property

Public Property Get Icone() As IPictureDisp
    If bRegistrado Then Set Icone = m_Icone
End Property

Public Property Set Icone(ByVal Value As IPictureDisp)
    If bRegistrado Then
        Set m_Icone = Value
        Set imgIcone.Picture = m_Icone
        PropertyChanged "Icone"
    End If
End Property

Public Property Get Descricao() As String
    If bRegistrado Then
        Descricao = m_Descricao
    End If
End Property

Public Property Let Descricao(ByVal Value As String)
    If bRegistrado Then
        m_Descricao = Value
        lblDescricao = m_Descricao
        PropertyChanged "Descricao"
    End If
End Property

Public Property Get Formulario() As String
    If bRegistrado Then
        Formulario = m_Formulario
    End If
End Property

Public Property Let Formulario(ByVal Value As String)
    If bRegistrado Then
        m_Formulario = Value
        lblFormulario = m_Formulario
        PropertyChanged "Formulario"
    End If
End Property

Private Sub Timer1_Timer()
    If bRegistrado Then lblHora = Format$(Now, "hh:mm:ss")
End Sub

Private Sub UserControl_Initialize()
    ValidaComponente "CABECALHO"
    If bRegistrado Then lblHora = Format$(Now, "hh:mm:ss")
End Sub

Private Sub UserControl_InitProperties()
    If bRegistrado Then
        Formulario = "Formulario"
        Descricao = "Descricao"
        Set Icone = imgIcone.Picture
        Codigo = "SABC101"
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Formulario = PropBag.ReadProperty("Formulario", "Formulario")
        Descricao = PropBag.ReadProperty("Descricao", "Descricao")
        Set Icone = PropBag.ReadProperty("Icone", imgIcone.Picture)
        Codigo = PropBag.ReadProperty("Codigo", "SABC101")
    End If
End Sub

Private Sub UserControl_Resize()
    If bRegistrado Then
        fraCabecalho.Width = Width
        lblHora.Left = fraCabecalho.Width - lblHora.Width - 70
        lblHora.Left = lblHora.Left
        lblDescricao.Width = Width - lblDescricao.Left
        Height = fraCabecalho.Height
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Call PropBag.WriteProperty("Formulario", m_Formulario, "Formulario")
        Call PropBag.WriteProperty("Descricao", m_Descricao, "Descricao")
        Call PropBag.WriteProperty("Icone", m_Icone, "")
        Call PropBag.WriteProperty("Codigo", m_Codigo, "SABC101")
    End If
End Sub

Public Function Exibir(Bdados As Object, CodForm As String, Path As String) As String
    On Error Resume Next
    
    If bRegistrado Then
        Codigo = CodForm
        Exibir = CodForm
        UserControl.Parent.Caption = CodForm
        Set Icone = ImageList1.ListImages("visual").ExtractIcon
        If Bdados.AbreTabela("SELECT * FROM TAB_FORMULARIO WHERE TFO_TMO_COD_MODULO='" & Mid(CodForm, 1, 4) & "' AND TFO_COD_FORMULARIO=" & Mid(CodForm, 5, 3)) Then
            Formulario = "" & Bdados.Tabela!TFO_NOME
            Descricao = "" & Bdados.Tabela!TFO_DESCR
            If Right$(Path, 1) = "\" Then
                Set Icone = LoadPicture(Path & "\imagens\" & CodForm)
            Else
                Set Icone = LoadPicture(Path & "\imagens\" & CodForm & ".ico")
            End If
        End If
        Bdados.FechaTabela
    End If
End Function

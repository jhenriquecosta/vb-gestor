VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.UserControl rodVISUAL 
   Alignable       =   -1  'True
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   4800
   ToolboxBitmap   =   "rodVISUAL.ctx":0000
   Begin Threed.SSFrame fraRodape 
      Height          =   465
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   820
      _Version        =   196610
      ForeColor       =   -2147483633
      Begin VB.Label lblSistema 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema"
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
         Left            =   210
         TabIndex        =   2
         Top             =   30
         Width           =   690
      End
      Begin VB.Label lblModulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modulo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   165
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   450
      End
   End
End
Attribute VB_Name = "rodVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Sistema As String
Private m_Modulo As String
Private m_VerMaior As Integer
Private m_VerMenor As Integer
Private m_VerRevisao As Integer
Private m_CorFundo As OLE_COLOR
Private m_CorFrente As OLE_COLOR

Public Property Get CorFrente() As OLE_COLOR
    If bRegistrado Then CorFrente = m_CorFrente
End Property

Public Property Let CorFrente(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        m_CorFrente = Value
        lblModulo.ForeColor = m_CorFrente
        lblSistema.ForeColor = m_CorFrente
        PropertyChanged "CorFrente"
    End If
End Property

Public Property Get CorFundo() As OLE_COLOR
    If bRegistrado Then CorFundo = m_CorFundo
End Property

Public Property Let CorFundo(ByVal Value As OLE_COLOR)
    If bRegistrado Then
        m_CorFundo = Value
        fraRodape.BackColor = m_CorFundo
        PropertyChanged "CorFundo"
    End If
End Property

Public Property Get VerRevisao() As Integer
    If bRegistrado Then VerRevisao = m_VerRevisao
End Property

Public Property Let VerRevisao(ByVal Value As Integer)
    If bRegistrado Then
        m_VerRevisao = Value
        If m_VerMaior = 0 And m_VerMenor = 0 And m_VerRevisao = 0 Then
            lblSistema = m_Sistema
        Else
            lblSistema = m_Sistema & " " & m_VerMaior & "." & m_VerMenor & "." & m_VerRevisao
        End If
        PropertyChanged "VerRevisao"
    End If
End Property

Public Property Get VerMenor() As Integer
    If bRegistrado Then VerMenor = m_VerMenor
End Property

Public Property Let VerMenor(ByVal Value As Integer)
    If bRegistrado Then
        m_VerMenor = Value
        PropertyChanged "VerMenor"
    End If
End Property

Public Property Get VerMaior() As Integer
    If bRegistrado Then VerMaior = m_VerMaior
End Property

Public Property Let VerMaior(ByVal Value As Integer)
    If bRegistrado Then
        m_VerMaior = Value
        PropertyChanged "VerMaior"
    End If
End Property


Public Property Get Modulo() As String
    If bRegistrado Then Modulo = m_Modulo
End Property

Public Property Let Modulo(ByVal Value As String)
    If bRegistrado Then
        m_Modulo = Value
        lblModulo = m_Modulo
        PropertyChanged "Modulo"
    End If
End Property

Public Property Get Sistema() As String
    If bRegistrado Then Sistema = m_Sistema
End Property

Public Property Let Sistema(ByVal Value As String)
    If bRegistrado Then
        m_Sistema = Value
        If m_VerMaior = 0 And m_VerMenor = 0 And m_VerRevisao = 0 Then
            lblSistema = m_Sistema
        Else
            lblSistema = m_Sistema & " " & m_VerMaior & "." & m_VerMenor & "." & m_VerRevisao
        End If
        PropertyChanged "Sistema"
    End If
End Property

Private Sub UserControl_Initialize()
    ValidaComponente "CABECALHO"
End Sub

Private Sub UserControl_InitProperties()
    If bRegistrado Then
        Sistema = "Sistema"
        Modulo = "Modulo"
        VerMaior = 0
        VerMenor = 0
        VerRevisao = 0
        CorFundo = Ambient.BackColor
        CorFrente = Ambient.ForeColor
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Modulo = PropBag.ReadProperty("Modulo", "Modulo")
        VerMaior = PropBag.ReadProperty("VerMaior", 0)
        VerMenor = PropBag.ReadProperty("VerMenor", 0)
        VerRevisao = PropBag.ReadProperty("VerRevisao", 0)
        Sistema = PropBag.ReadProperty("Sistema", "Sistema")
        CorFundo = PropBag.ReadProperty("CorFundo", Ambient.BackColor)
        CorFrente = PropBag.ReadProperty("CorFrente", Ambient.ForeColor)
    End If
End Sub

Private Sub UserControl_Resize()
    If bRegistrado Then
        fraRodape.Width = Width + 150
        fraRodape.Height = Height + 30
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    If bRegistrado Then
        Call PropBag.WriteProperty("Sistema", m_Sistema, "Sistema")
        Call PropBag.WriteProperty("Modulo", m_Modulo, "Modulo")
        Call PropBag.WriteProperty("VerMaior", m_VerMaior, 0)
        Call PropBag.WriteProperty("VerMenor", m_VerMenor, 0)
        Call PropBag.WriteProperty("VerRevisao", m_VerRevisao, 0)
        Call PropBag.WriteProperty("CorFundo", m_CorFundo, Ambient.BackColor)
        Call PropBag.WriteProperty("CorFrente", m_CorFrente, Ambient.ForeColor)
    End If
End Sub

Public Sub Exibir(Bdados As Object, CodForm As String, Optional Maior, Optional Menor, Optional Revisao)
    On Error Resume Next
    Dim sql As String
    
    If bRegistrado Then
        sql = "SELECT TSI_NOME, TMO_NOME " & _
                " FROM TAB_SISTEMA, TAB_MODULO" & _
                " WHERE TMO_COD_MODULO='" & Mid(CodForm, 1, 4) & "'" & _
                    " AND TMO_TSI_COD_SISTEMA=TSI_COD_SISTEMA"
        If Bdados.AbreTabela(sql) Then
            Sistema = "" & Bdados.Tabela!TSI_NOME
            Modulo = "" & Bdados.Tabela!TMO_NOME
            If Not IsMissing(Maior) Then VerMaior = Maior
            If Not IsMissing(Menor) Then VerMenor = Menor
            If IsMissing(Revisao) Then
                VerRevisao = m_VerRevisao
            Else
                VerRevisao = Revisao
            End If
        End If
        Bdados.FechaTabela
    End If
End Sub

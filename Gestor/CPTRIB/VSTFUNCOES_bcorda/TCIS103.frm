VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Begin VB.Form TCIS103 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TCIS103"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VTOcx.cmdVISUAL CmdOk 
      Height          =   345
      Left            =   960
      TabIndex        =   3
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   "OK"
      Acao            =   3
   End
   Begin VTOcx.cboVISUAL cboDia 
      Height          =   315
      Left            =   495
      TabIndex        =   0
      Top             =   840
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "Dia"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cboVISUAL cboMes 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "Mês"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cboVISUAL cboAno 
      Height          =   315
      Left            =   2820
      TabIndex        =   2
      Top             =   840
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   "Ano"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   1138
      Formulario      =   "Nova data de vencimento"
      Descricao       =   "Vencimento"
      Icone           =   "TCIS103.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdCancela 
      Height          =   345
      Left            =   2190
      TabIndex        =   4
      Top             =   1290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      Caption         =   "CANCELA"
      Acao            =   7
   End
End
Attribute VB_Name = "TCIS103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancela_Click()
    NovaDataVencimento = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cboDia.ListIndex < 0 Then
        Util.Avisa "Selecione Dia."
        cboDia.SetFocus
        Exit Sub
    End If
    
    If cboMes.ListIndex < 0 Then
        Util.Avisa "Selecione Mês."
        cboMes.SetFocus
        Exit Sub
    End If
    
    If cboAno.ListIndex < 0 Then
        Util.Avisa "Selecione Ano."
        cboAno.SetFocus
        Exit Sub
    End If
    NovaDataVencimento = cboDia & "/" & cboMes & "/" & cboAno
    If Not IsDate(NovaDataVencimento) Then
        Util.Avisa "Data Inválida."
        cboDia.SetFocus
        Exit Sub
    End If
    If Not (AplicacoesVTFuncoes.Municipio = "PETROLINA" Or AplicacoesVTFuncoes.Municipio = "VERDEJANTE" Or AplicacoesVTFuncoes.Municipio = "LAGOA GRANDE") Then
        If CDate(NovaDataVencimento) < Date Then
            Util.Avisa "Vencimento inválido."
            cboDia.SetFocus
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim I As Integer
    NovaDataVencimento = ""
    For I = 1 To 31
        'dias
        cboDia.AddItem Format(I, "00")
    Next
    For I = 1 To 12
        'Meses
        cboMes.AddItem Format(I, "00")
    Next
    For I = 1994 To 3000
        'Meses
        cboAno.AddItem Format(I, "0000")
    Next
    Seta
End Sub
Private Sub Seta()
    Dim I As Integer
    For I = 0 To cboDia.ListCount
        If cboDia.List(I) = Format(Day(Me.Tag), "00") Then
            cboDia.ListIndex = I
        End If
    Next
    
    For I = 0 To cboMes.ListCount
        If cboMes.List(I) = Format(Month(Me.Tag), "00") Then
            cboMes.ListIndex = I
        End If
    Next
    
    For I = 0 To cboAno.ListCount
        If cboAno.List(I) = Format(Year(Me.Tag), "00") Then
            cboAno.ListIndex = I
        End If
    Next
End Sub

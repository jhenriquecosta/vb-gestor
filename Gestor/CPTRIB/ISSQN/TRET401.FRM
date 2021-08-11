VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#2.0#0"; "VTControles.ocx"
Begin VB.Form TRET401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRET401"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraPeriodo 
      Height          =   900
      Left            =   45
      TabIndex        =   7
      Top             =   690
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   1588
      Altura          =   1905
      Caption         =   " Período(mm/yyyy)"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483626
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6945
         TabIndex        =   2
         Top             =   405
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtPerFin 
         Height          =   285
         Left            =   4440
         TabIndex        =   1
         Tag             =   "Data Emissao"
         Top             =   450
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   503
         Caption         =   "Final"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   7
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPerImic 
         Height          =   285
         Left            =   1095
         TabIndex        =   0
         Tag             =   "Data Emissao"
         Top             =   435
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   503
         Caption         =   "Inicial"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   7
         RetirarMascara  =   0   'False
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   1138
      Icone           =   "TRET401.frx":0000
   End
   Begin VTOcx.grdVISUAL Grid 
      Height          =   3105
      Left            =   45
      TabIndex        =   3
      Top             =   1635
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   5477
      CorBorda        =   32768
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   8
      Top             =   4785
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7125
         TabIndex        =   5
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   5955
         TabIndex        =   4
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
End
Attribute VB_Name = "TRET401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nota As cNota

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdBuscar_Click()
    Set Nota = New cNota
    With Nota
        If .PreencherGrid(Grid, txtPerImic, txtPerFin) Then
            Grid.Mensagem = ColocarEspacos((CInt(Grid.Width) / 54) - (Len(.Arrecadacao) + 2)) & .Arrecadacao
        Else
            Avisa "Nenhuma nota."
            Grid.ListItems.Clear
            Grid.Mensagem = ""
            txtPerImic.SetFocus
        End If
    End With
    Set Nota = Nothing
End Sub
Private Function ColocarEspacos(Num As Integer) As String
Dim i As Integer
i = 0
Do While i <= Num
    ColocarEspacos = ColocarEspacos & " "
    i = i + 1
Loop
End Function


Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    Grid.ListItems.Clear
    Grid.Mensagem = ColocarEspacos(4)
    txtPerImic.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSair_LostFocus()
    txtPerImic.SetFocus
End Sub

Private Sub txtPerImic_LostFocus()
    If Trim(txtPerImic) = "" Then Exit Sub
    If IsNumeric(txtPerImic) Then
        If Len(txtPerImic) = 6 Then
            txtPerImic = Left(txtPerImic, 2) & "/" & Right(txtPerImic, 4)
        Else
            Avisa "Período inválido."
            txtPerImic.SetFocus
        End If
    End If
End Sub

Private Sub txtPerFin_LostFocus()
    If Trim(txtPerFin) = "" Then Exit Sub
    If IsNumeric(txtPerFin) Then
        If Len(txtPerFin) = 6 Then
            txtPerFin = Left(txtPerFin, 2) & "/" & Right(txtPerFin, 4)
        Else
            Avisa "Período inválido."
            txtPerFin.SetFocus
        End If
    End If
End Sub

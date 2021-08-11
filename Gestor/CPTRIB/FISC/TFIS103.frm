VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TFIS103 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "TFIS103.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPar 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1185
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "Parâmetro"
      Top             =   960
      Width           =   7320
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   1349
      CorFundo        =   16774636
      CorFrente       =   12632064
   End
   Begin VTOcx.cmdVISUAL cmdGravar 
      Height          =   375
      Left            =   2670
      TabIndex        =   1
      Top             =   1380
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdDeletar 
      Height          =   375
      Left            =   4140
      TabIndex        =   2
      Top             =   1380
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Deletar"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdNovo 
      Height          =   375
      Left            =   5610
      TabIndex        =   3
      Top             =   1380
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "&Novo"
      Acao            =   1
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   1380
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   3990
      Left            =   60
      TabIndex        =   6
      Tag             =   "Documento gerencial"
      Top             =   1860
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   7038
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      Tabs            =   "TFIS103.frx":08CA
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3600
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   6350
         _Version        =   131082
         TabGuid         =   "TFIS103.frx":094C
         Begin VTOcx.grdVISUAL Grid 
            Height          =   3450
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   6085
            CorBorda        =   32768
            Caption         =   "Parametros"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            OcultarRodape   =   -1  'True
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3600
         Index           =   1
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   6350
         _Version        =   131082
         TabGuid         =   "TFIS103.frx":0974
         Begin VB.TextBox txtDes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3315
            Left            =   0
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Tag             =   "Descrição"
            Top             =   150
            Width           =   8310
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parâmetro"
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
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1035
      Width           =   915
   End
End
Attribute VB_Name = "TFIS103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Parametro As New Parametros
Private Sub cmdDeletar_Click()
    If txtPar.Enabled = False And Trim(txtPar) <> "" Then
        If Util.Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If Parametro.Deletar(Grid.SelectedItem) Then
                Util.Informa "Parâmentro apagado."
            Else
                Util.Avisa "Parâmetro não apagado."
            End If
        End If
    Else
        Util.Avisa "Selecione um registro gravado."
    End If
    Call cmdNovo_Click
End Sub

Private Sub cmdNovo_Click()
    txtPar.Enabled = True
    txtPar = ""
    txtDes = ""
    Parametro.vCodigo = ""
    Parametro.PreencheGrid Grid
    
    txtPar.SetFocus
End Sub

Private Sub Grid_DblClick()
    On Error GoTo Trata
    TabDados.Tabs(2).Selected = True
    txtDes.SetFocus
    Parametro.vCodigo = Grid.SelectedItem
    txtPar = Grid.SelectedItem.SubItems(1)
    txtDes = Grid.SelectedItem.SubItems(2)
    
    txtPar.Enabled = False
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
    Exit Sub
    Resume
End Sub

Private Sub txtPar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    
    If Edita.CriticaCampos(Me) Then
        If Util.Confirma("Deseja salvar os dados de " & txtPar & "?") Then
            Screen.MousePointer = 11
            
            Parametro.vParametro = txtPar
            Parametro.vDescricao = txtDes
            If Parametro.Gravar Then
                Call cmdNovo_Click
                Util.Informa "Operação realizada."
            End If
        End If
    End If
    TabDados.Tabs(1).Selected = True
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo Trata
    
    Instala.NovoPerfil Me, ctlCabecalho1, Cod_sis, Sistema, Desc_Form, App.Path & "\imagens\"
    Parametro.PreencheGrid Grid
    Screen.MousePointer = 0
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtPar_LostFocus()
    Dim i As Integer
    txtPar = Trim(txtPar)
    For i = 1 To Grid.ListItems.Count
        If Grid.ListItems(i).Text = Trim(txtPar) Then
            txtDes.Text = Grid.ListItems(i).ListSubItems.Item(1).Text
            Exit For
        End If
    Next
    
End Sub

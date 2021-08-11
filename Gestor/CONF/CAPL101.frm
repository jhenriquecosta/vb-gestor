VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CAPL101 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CAPL101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comDig 
      Left            =   45
      Top             =   3990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "Localização do arquivo"
      FileName        =   "*.mdb"
      InitDir         =   "C:\"
   End
   Begin VB.ComboBox cboSis 
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
      Height          =   315
      ItemData        =   "CAPL101.frx":08CA
      Left            =   2805
      List            =   "CAPL101.frx":08DA
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4935
      Width           =   1305
   End
   Begin VB.ComboBox cboTipo 
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
      Height          =   315
      ItemData        =   "CAPL101.frx":0900
      Left            =   630
      List            =   "CAPL101.frx":0910
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4950
      Width           =   1305
   End
   Begin VB.TextBox txtArq 
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
      Height          =   1560
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   5700
      Width           =   7185
   End
   Begin VB.TextBox txtCat 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4845
      MaxLength       =   50
      TabIndex        =   13
      Top             =   5310
      Width           =   1290
   End
   Begin VB.TextBox txtPass 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2820
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   5310
      Width           =   1275
   End
   Begin VB.TextBox txtUser 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   630
      MaxLength       =   50
      TabIndex        =   11
      Top             =   5325
      Width           =   1290
   End
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
      Height          =   690
      Left            =   1050
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1905
      Width           =   6180
   End
   Begin VB.TextBox txtDSN 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4845
      TabIndex        =   10
      Top             =   4935
      Width           =   2385
   End
   Begin VB.TextBox txtNome 
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
      Height          =   315
      Left            =   1050
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Nome"
      Top             =   1515
      Width           =   3585
   End
   Begin VB.TextBox txtCod 
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
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Código"
      Top             =   1125
      Width           =   1395
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   1620
      Left            =   60
      TabIndex        =   6
      Top             =   2970
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   2858
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   12582912
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   0
   End
   Begin Cabecalho.ctlCabecalho ctlCabecalho1 
      Align           =   1  'Align Top
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1349
      CorFundo        =   16774636
      CorFrente       =   12632064
   End
   Begin Threed.SSCommand cmdCop 
      Height          =   315
      Left            =   75
      TabIndex        =   16
      ToolTipText     =   "Deseja sair?"
      Top             =   7320
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL101.frx":0936
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL101.frx":0952
      Caption         =   "Copiar para a área de &transferência"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   2190
      TabIndex        =   28
      Top             =   5010
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   27
      Top             =   5040
      Width           =   300
   End
   Begin Threed.SSCommand cmdPro 
      Height          =   315
      Left            =   6225
      TabIndex        =   14
      Top             =   5310
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL101.frx":096E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL101.frx":098A
      Caption         =   "&Processar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Geração de código de Cofiguração"
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
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   26
      Top             =   4665
      Width           =   7185
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Catalog"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4200
      TabIndex        =   25
      Top             =   5385
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2055
      TabIndex        =   24
      Top             =   5400
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   225
      TabIndex        =   23
      Top             =   5385
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   22
      Top             =   1980
      Width           =   690
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Sistema selecionado"
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
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   21
      Top             =   810
      Width           =   7185
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDSN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dsn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4470
      TabIndex        =   20
      Top             =   4995
      Width           =   270
   End
   Begin Threed.SSCommand cmdDeletar 
      Height          =   315
      Left            =   5595
      TabIndex        =   4
      Top             =   1515
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL101.frx":09A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL101.frx":09C2
      Caption         =   "&Apagar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdGravar 
      Height          =   315
      Left            =   4740
      TabIndex        =   3
      Top             =   1515
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL101.frx":09DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL101.frx":09FA
      Caption         =   "&Salvar"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   390
      TabIndex        =   19
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   525
      TabIndex        =   18
      Top             =   1575
      Width           =   405
   End
   Begin Threed.SSCommand cmdSair 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6225
      TabIndex        =   7
      ToolTipText     =   "Deseja sair?"
      Top             =   7335
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL101.frx":0A16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL101.frx":0A32
      Caption         =   "Sai&r"
      ButtonStyle     =   4
      PictureAlignment=   6
   End
   Begin Threed.SSCommand cmdNovo 
      Height          =   315
      Left            =   6435
      TabIndex        =   5
      Top             =   1515
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   196610
      MousePointer    =   16
      ForeColor       =   12632064
      PictureFrames   =   1
      BackStyle       =   1
      Windowless      =   -1  'True
      MouseIcon       =   "CAPL101.frx":0A4E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "CAPL101.frx":0A6A
      Caption         =   "&Novo"
      ButtonStyle     =   4
      PictureAlignment=   6
      ShapeSize       =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " Tabela de Sistemas"
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
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   2685
      Width           =   7185
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "CAPL101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COD As String

Private Sub cboTipo_LostFocus()
    If cboTipo = "Access" Then
        lblDSN.FontBold = True
        lblDSN.FontUnderline = True
        lblDSN.ForeColor = &HFF0000
    Else
        LimpaDSN
    End If
End Sub

Private Sub LimpaDSN()
    lblDSN.FontBold = False
    lblDSN.FontUnderline = False
    lblDSN.ForeColor = &H0&
End Sub

Private Sub cmdCop_Click()
    Clipboard.SetText txtArq
End Sub

Private Sub cmdDeletar_Click()
    If COD <> "" And Trim(txtCod) <> "" Then
        If Util.Confirma("Deseja mesmo apagar '" & txtCod & "'?") Then
            If Bdados.DeletaDados("TAB_SISTEMA", "TSI_COD_SISTEMA='" & txtCod & "'") Then
                Util.Informa "Registro apagado."
            Else
                Util.Avisa "Registro não apagado."
            End If
        End If
    Else
        Util.Avisa "Selecione um registro gravado."
    End If
    Call cmdNovo_Click
End Sub

Private Sub cmdNovo_Click()
    COD = ""
    LimpaDSN
    Edita.LimpaCampos Me
    combos
    AtualizaG
    
    txtCod.SetFocus
End Sub

Private Sub combos()
    cboTipo.Clear
    cboTipo.AddItem "Access"
    cboTipo.AddItem "Sql Server"
    cboTipo.AddItem "Oracle"
    cboTipo.AddItem "ODBC"
    
    Edita.AtualizaCombo Bdados, cboSis, "SELECT TSI_COD_SISTEMA FROM TAB_SISTEMA"
    cboSis.AddItem "SEG"
End Sub


Private Sub AtualizaG()
    Call Util.MontaGrid(Bdados, Grid, _
    "SELECT TSI_COD_SISTEMA AS Código, TSI_NOME as Nome,TSI_DESCR AS Descrição FROM TAB_SISTEMA", 900, 2500, 7000)
End Sub

Private Sub cmdPro_Click()
    If cboTipo = "" Or cboSis = "" Or txtDSN = "" Then
        Util.Avisa "Informe o tipo, sistema e DSN para gerar o arquivo."
    Else
        txtArq = Instala.GeraConfig(cboTipo.ListIndex, cboSis, txtDSN, txtUser, txtPass, txtCat)
    End If
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call Util.OrdenaGrid(Grid, ColumnHeader)
End Sub


Private Sub Grid_DblClick()
    On Error GoTo Trata
    txtCod.SetFocus
    
    txtCod = Grid.SelectedItem.Text
    txtNome = Grid.SelectedItem.SubItems(1)
    txtDes = Grid.SelectedItem.SubItems(2)
    
    COD = Grid.SelectedItem.Text
    
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub lblDSN_Click()
    If lblDSN.FontBold Then
        comDig.ShowOpen
        txtDSN = comDig.FileName
    End If
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Atualizar(Codigo As String)
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    
    If COD = "" Then COD = Codigo
    
    Campos = "TSI_COD_SISTEMA,TSI_NOME,TSI_DESCR"
    Valores = Bdados.PreparaValor(txtCod, txtNome, txtDes)
    
    Call Bdados.GravaDados("TAB_SISTEMA", Valores, Campos, _
    "TSI_COD_SISTEMA = '" & COD & "'")

    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdGravar_Click()
    On Error GoTo Trata
    txtCod = Trim(txtCod)
    If Edita.CriticaCampos(Me) Then
        If IsNumeric(txtCod) Or Len(txtCod) < 4 Then
            Util.Avisa "Código deve conter 4 caracteres."
        Else
            If Util.Confirma("Deseja salvar os dados de '" & txtCod & "'?") Then
                Screen.MousePointer = 11
                Call Atualizar(txtCod)
                Call cmdNovo_Click
                Util.Informa "Operação realizada."
            End If
        End If
    End If
    
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
    
    AtualizaG
    combos
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

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Grid_DblClick
End Sub

Private Sub txtCod_LostFocus()
    On Error GoTo Trata
    Dim I As Integer
    txtCod = Trim(txtCod)
    If COD = "" Then
        For I = 1 To Grid.ListItems.Count
            If Grid.ListItems(I).Text = Trim(txtCod) Then
                txtNome.Text = Grid.ListItems(I).ListSubItems.Item(1).Text
                txtDSN.Text = Grid.ListItems(I).ListSubItems.Item(2).Text
                txtDes.Text = Grid.ListItems(I).ListSubItems.Item(3).Text
                COD = txtCod
                Exit For
            End If
        Next
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 35600 Then Resume Next
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Private Sub txtDSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And lblDSN.FontBold Then Call lblDSN_Click
End Sub

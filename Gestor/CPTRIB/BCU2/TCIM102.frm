VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#1.1#0"; "VTControles.ocx"
Begin VB.Form TCIM102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.txtVISUAL txtDistrito 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Tag             =   "Codigo"
      Top             =   750
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "Distrito"
      Text            =   ""
      Restricao       =   2
      AlinhamentoTexto=   2
      MaxLen          =   2
      MinLen          =   2
   End
   Begin Cabecalho.cabVISUAL cabCabecalho 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   1138
      Formulario      =   "CODIGO"
      Icone           =   "TCIM102.frx":0000
   End
   Begin VTOcx.grdVISUAL grdDistritoSetor 
      Height          =   2625
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4339
      CorBorda        =   -2147483646
      Caption         =   "Distritos/Setores"
      CorTitulo       =   -2147483646
      CorCaption      =   -2147483639
      CorDica         =   -2147483646
   End
   Begin Cabecalho.rodVISUAL rodRodape 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   3795
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   926
      CorFundo        =   -2147483632
      CorFrente       =   -2147483633
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   405
         Left            =   2820
         TabIndex        =   3
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   714
         Caption         =   "&Apagar"
         Acao            =   2
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   3900
         TabIndex        =   4
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   4980
         TabIndex        =   5
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
      Begin VTOcx.cmdVISUAL cmdGravar 
         Height          =   405
         Left            =   1740
         TabIndex        =   2
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   714
         Caption         =   "&Gravar"
         Acao            =   3
         CorBorda        =   -2147483645
         CorFrente       =   -2147483630
         CorFoco         =   -2147483628
      End
   End
   Begin VTOcx.txtVISUAL txtSetor 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "Codigo"
      Top             =   750
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "Setor"
      Text            =   ""
      Restricao       =   2
      AlinhamentoTexto=   2
      MaxLen          =   2
      MinLen          =   2
   End
End
Attribute VB_Name = "TCIM102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Capa As cDistritoSetor

Private Sub cmdExcluir_Click()
    Set Capa = New cDistritoSetor
    If Not (txtDistrito <> "" And txtSetor <> "") Then Exit Sub
    If Capa.Excluir(txtDistrito, txtSetor) Then
        Util.Informa "Registro Apagado."
        Capa.PreencherGrd grdDistritoSetor
        cmdLimpar_Click
    Else
        Util.Erro "Erro ao apagar."
    End If
End Sub

Private Sub cmdGravar_Click()
    If Edita.CriticaCampos(Me) Then
        Set Capa = New cDistritoSetor
        With Capa
            .Distrito = txtDistrito
            .Setor = txtSetor
            If .Gravar Then
                Util.Informa "Registro Cadastrado."
                Capa.PreencherGrd grdDistritoSetor
                cmdLimpar_Click
            Else
                Util.Erro "Erro ao gravar."
            End If
        End With
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    Set Capa = Nothing
    txtDistrito.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabCabecalho.Exibir Bdados, Me.Name, App.Path
    rodRodape.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Set Capa = New cDistritoSetor
    Capa.PreencherGrd grdDistritoSetor
    Set Capa = Nothing
End Sub

Private Sub grdDistritoSetor_Click()
    If Not (grdDistritoSetor.SelectedItem Is Nothing) Then
        txtDistrito = grdDistritoSetor.SelectedItem.Text
        txtSetor = grdDistritoSetor.SelectedItem.SubItems(1)
        txtDistrito.SetFocus
    End If
End Sub

VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCAF102 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   12
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCAF102.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1455
      Left            =   75
      TabIndex        =   3
      Top             =   3435
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      Altura          =   1905
      Caption         =   " Dados Livro"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483626
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtOrdem 
         Height          =   300
         Left            =   195
         TabIndex        =   8
         Tag             =   "Total"
         Top             =   1065
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         Caption         =   "Proximo Nº Ordem"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtSituacao 
         Height          =   300
         Left            =   4245
         TabIndex        =   7
         Top             =   705
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         Caption         =   "Situação"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtFolha 
         Height          =   300
         Left            =   885
         TabIndex        =   5
         Top             =   705
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   529
         Caption         =   "Folha Atual"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtTotal 
         Height          =   300
         Left            =   2685
         TabIndex        =   6
         Tag             =   "Total"
         Top             =   705
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         Caption         =   "Total"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         Mascara         =   "0000"
      End
      Begin VTOcx.txtVISUAL txtLivro 
         Height          =   300
         Left            =   360
         TabIndex        =   4
         Tag             =   "Livro"
         Top             =   360
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         Caption         =   "Livro"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         Mascara         =   "0000"
      End
   End
   Begin VTOcx.grdVISUAL grdAforamento 
      Height          =   2715
      Left            =   45
      TabIndex        =   1
      Top             =   690
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   4339
      Caption         =   "Livros Aforamento"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   255
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   1138
      Icone           =   "TCAF102.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   4980
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   3030
         TabIndex        =   9
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   4980
         TabIndex        =   11
         Top             =   120
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4005
         TabIndex        =   10
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
   End
End
Attribute VB_Name = "TCAF102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AforManu As New cAforManu
Dim Aforamento As New cAforamento

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    AforManu.PreencherGrid grdAforamento
    txtOrdem = Aforamento.ProximoAforamento
    txtFolha = 0
    txtSituacao = 1
    txtLivro.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valida As Boolean
    Valida = False
    If Edita.CriticaCampos(Me) Then
        With AforManu
            If .Buscar(txtLivro) = False Then
                .DataAbertura = Format(Date, "DD/MM/YYYY")
            End If
            .FolhaAtual = txtFolha
            .FolhaTotal = txtTotal
            .Status = Nvl(txtSituacao, 1)
            .CodUsuario = Aplicacoes.Usuario
            If .Gravar(txtLivro) = False Then
                Valida = False
            Else
                Valida = True
            End If
            If .GravarCorrelativo(txtOrdem) = False Then
                Valida = False
            Else
                Valida = True
            End If
        End With
        If Valida = True Then
            Util.Informa "Dados Atualizados com sucesso."
            cmdNovo_Click
        Else
            Util.Informa "Não foi possivel atualziar dados."
        End If
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    Set AforManu = New cAforManu
    AforManu.PreencherGrid grdAforamento
    txtOrdem = Aforamento.ProximoAforamento()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AforManu = Nothing
    Set Aforamento = Nothing
End Sub

Private Sub grdAforamento_Click()
    If Not grdAforamento.SelectedItem Is Nothing Then
        With grdAforamento.SelectedItem
            'If .SubItems(3) = 1 Then
                txtLivro = .Text
                txtFolha = .SubItems(1)
                txtTotal = .SubItems(2)
                txtSituacao = .SubItems(3)
            'End If
        End With
    End If
End Sub

Private Sub txtLivro_LostFocus()
'    On Error Resume Next
'    If Trim$(txtLivro) <> "" Then
'        Dim Item As ListItem
'        Set Item = grdAforamento.FindItem(CDbl(Nvl(Trim(txtLivro), 0)))
'        If Not Item Is Nothing Then
'            Item.Selected = True
'            Item.EnsureVisible
'            grdAforamento_Click
'        End If
'    End If
End Sub

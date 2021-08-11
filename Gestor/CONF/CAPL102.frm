VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form CAPL102 
   BackColor       =   &H00FFF5EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
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
   ScaleHeight     =   7770
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   15
      Top             =   7260
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   900
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   7770
         TabIndex        =   25
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   0
         CorFrente       =   0
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9000
         TabIndex        =   16
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   0
         CorFrente       =   0
      End
   End
   Begin ActiveTabs.SSActiveTabs tabGeral 
      Height          =   6420
      Left            =   60
      TabIndex        =   17
      Top             =   720
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   11324
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "CAPL102.frx":0000
      Images          =   "CAPL102.frx":00B6
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6000
         Index           =   0
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   10583
         _Version        =   131082
         TabGuid         =   "CAPL102.frx":0709
         Begin VTOcx.txtVISUAL txtSistema 
            Height          =   315
            Left            =   1785
            TabIndex        =   1
            Top             =   4380
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            TipoLetras      =   0
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtCodSistema 
            Height          =   315
            Left            =   60
            TabIndex        =   0
            Top             =   4380
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            Caption         =   "Sistema"
            Text            =   ""
            Restricao       =   1
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarSistema 
            Height          =   405
            Left            =   8865
            TabIndex        =   3
            Top             =   5505
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirSistema 
            Height          =   405
            Left            =   9255
            TabIndex        =   4
            Top             =   5505
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdSistemas 
            Height          =   4260
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   7514
            CorFundo        =   -2147483633
            Caption         =   "Sistemas"
            CorTitulo       =   4210688
         End
         Begin VTOcx.txtVISUAL txtDescSistema 
            Height          =   705
            Left            =   900
            TabIndex        =   2
            Top             =   4740
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   1244
            Caption         =   "Descrição"
            Text            =   ""
            TipoLetras      =   0
            AlinhamentoRotuloVertical=   0
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdLimparSistema 
            Height          =   405
            Left            =   8475
            TabIndex        =   26
            Top             =   5505
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   6
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6000
         Index           =   2
         Left            =   -99969
         TabIndex        =   20
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   10583
         _Version        =   131082
         TabGuid         =   "CAPL102.frx":0731
         Begin VTOcx.txtVISUAL txtFormulario 
            Height          =   315
            Left            =   1785
            TabIndex        =   11
            Top             =   4380
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            TipoLetras      =   0
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtCodFormulario 
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   4380
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            Caption         =   "Formulario"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarFormulario 
            Height          =   405
            Left            =   8865
            TabIndex        =   13
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirFormulario 
            Height          =   405
            Left            =   9255
            TabIndex        =   14
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdFormularios 
            Height          =   4260
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   7514
            CorFundo        =   -2147483633
            Caption         =   "Formularios"
            CorTitulo       =   4210688
         End
         Begin VTOcx.txtVISUAL txtDescFormulario 
            Height          =   705
            Left            =   900
            TabIndex        =   12
            Top             =   4740
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   1244
            Caption         =   "Descrição"
            Text            =   ""
            TipoLetras      =   0
            AlinhamentoRotuloVertical=   0
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdLimparFormulario 
            Height          =   405
            Left            =   8475
            TabIndex        =   28
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   6
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6000
         Index           =   1
         Left            =   -99969
         TabIndex        =   22
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   10583
         _Version        =   131082
         TabGuid         =   "CAPL102.frx":0759
         Begin VTOcx.txtVISUAL txtModulo 
            Height          =   315
            Left            =   1770
            TabIndex        =   6
            Top             =   4380
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            TipoLetras      =   0
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtCodModulo 
            Height          =   315
            Left            =   60
            TabIndex        =   5
            Top             =   4380
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            Caption         =   "Modulo"
            Text            =   ""
            Restricao       =   1
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdSalvarModulo 
            Height          =   405
            Left            =   8865
            TabIndex        =   8
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   3
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdExcluirModulo 
            Height          =   405
            Left            =   9255
            TabIndex        =   9
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   2
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
         Begin VTOcx.grdVISUAL grdModulos 
            Height          =   4260
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   7514
            CorFundo        =   -2147483633
            Caption         =   "Modulos"
            CorTitulo       =   4210688
         End
         Begin VTOcx.txtVISUAL txtDescModulo 
            Height          =   705
            Left            =   900
            TabIndex        =   7
            Top             =   4740
            Width           =   8700
            _ExtentX        =   15346
            _ExtentY        =   1244
            Caption         =   "Descrição"
            Text            =   ""
            TipoLetras      =   0
            AlinhamentoRotuloVertical=   0
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdLimparModulo 
            Height          =   405
            Left            =   8475
            TabIndex        =   27
            Top             =   5520
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   714
            Caption         =   ""
            Acao            =   6
            CorBorda        =   -2147483632
            CorFundo        =   -2147483633
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1138
      Icone           =   "CAPL102.frx":0781
   End
End
Attribute VB_Name = "CAPL102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCodSistema As String
Private strCodModulo As String

Private Sub cmdAdicionarSistema_Click()
    PrepararSistema
End Sub

Private Sub cmdExcluirFormulario_Click()
    If ExcluirFormulario(txtCodFormulario, txtFormulario) Then
        PrepararFormulario
    End If
End Sub

Private Sub cmdExcluirModulo_Click()
    If ExcluirModulo(txtCodModulo, txtModulo) Then
        PrepararModulo
    End If
End Sub

Private Sub cmdExcluirSistema_Click()
    If ExcluirSistema(txtCodSistema, txtSistema) Then
        PrepararSistema
    End If
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo Trata
    Screen.MousePointer = vbHourglass
    With Relatorio
        .DefinirArquivo Bdados, App.Path & "\CAPL301.rpt"
        .Formulas "VT_Sistema", "'" & Temp.PegaParametro(Bdados, "SISTEMA") & "'"
        .Formulas "VT_Descricao", "'" & Temp.PegaParametro(Bdados, "DESCRICAO") & "'"
        If txtCodSistema = "" Then
            .Selecao = ""
        Else
            .Selecao = "{TAB_SISTEMA.TSI_COD_SISTEMA}='" & txtCodSistema & "'"
        End If
        .Visualizar
    End With
    Screen.MousePointer = vbDefault
Trata:
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdLimparFormulario_Click()
    txtCodFormulario = ""
    txtFormulario = ""
    txtDescFormulario = ""
    txtCodFormulario.SetFocus
End Sub

Private Sub cmdLimparModulo_Click()
    txtCodModulo = ""
    txtModulo = ""
    txtDescModulo = ""
    txtCodModulo.SetFocus
End Sub

Private Sub cmdLimparSistema_Click()
    txtCodSistema = ""
    txtSistema = ""
    txtDescSistema = ""
    txtCodSistema.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvarFormulario_Click()
    If GravarFormulario(txtCodFormulario, txtFormulario, txtDescFormulario) Then
        PrepararFormulario
    End If
End Sub

Private Sub cmdSalvarModulo_Click()
    If GravarModulo(txtCodModulo, txtModulo, txtDescModulo) Then
        PrepararModulo
    End If
End Sub

Private Sub cmdSalvarSistema_Click()
    If GravarSistema(txtCodSistema, txtSistema, txtDescSistema) Then
        PrepararSistema
    End If
End Sub

Private Sub Form_Load()
    PreencherSistemas
End Sub

Private Sub PreencherSistemas()
    Dim sql As String
    
    sql = "SELECT TSI_COD_SISTEMA AS Codigo, TSI_NOME as Sistema, TSI_DESCR as Descricao" & _
            " FROM TAB_SISTEMA" & _
            " ORDER BY TSI_COD_SISTEMA"
    grdSistemas.Preencher Bdados, sql
End Sub

Private Sub PreencherModulos(CodSistema As String)
    Dim sql As String
    
    sql = "SELECT TMO_COD_MODULO AS Codigo, TMO_NOME as Sistema, TMO_DESCR as Descricao" & _
            " FROM TAB_MODULO" & _
            " WHERE TMO_TSI_COD_SISTEMA='" & CodSistema & "'" & _
            " ORDER BY TMO_COD_MODULO"
    grdModulos.Preencher Bdados, sql
End Sub

Private Sub PreencherFormularios(CodModulo As String)
    Dim sql As String
    
    sql = "SELECT TFO_COD_FORMULARIO AS Codigo, TFO_NOME as Formulario, TFO_DESCR as Descricao" & _
            " FROM TAB_FORMULARIO" & _
            " WHERE TFO_TMO_COD_MODULO='" & CodModulo & "'" & _
            " ORDER BY TFO_COD_FORMULARIO"
    grdFormularios.Preencher Bdados, sql
End Sub

Private Sub grdFormularios_Click()
    If Not grdFormularios.SelectedItem Is Nothing Then
        With grdFormularios.SelectedItem
            txtCodFormulario = .Text
            txtFormulario = .SubItems(1)
            txtDescFormulario = .SubItems(2)
        End With
    End If
End Sub

Private Sub grdModulos_Click()
    If Not grdModulos.SelectedItem Is Nothing Then
        With grdModulos.SelectedItem
            strCodModulo = .Text
            txtCodModulo = .Text
            txtModulo = .SubItems(1)
            txtDescModulo = .SubItems(2)
            grdFormularios.Caption = "Formularios : " & .Text
            PreencherFormularios .Text
        End With
    End If
End Sub

Private Sub grdModulos_DblClick()
    tabGeral.Tabs(3).Selected = True
    txtCodFormulario.SetFocus
End Sub

Private Sub grdSistemas_DblClick()
    tabGeral.Tabs(2).Selected = True
    txtCodModulo.SetFocus
End Sub

Private Sub grdSistemas_Click()
    If Not grdSistemas.SelectedItem Is Nothing Then
        With grdSistemas.SelectedItem
            strCodSistema = .Text
            txtCodSistema = .Text
            txtSistema = .SubItems(1)
            txtDescSistema = .SubItems(2)
            grdModulos.Caption = "Modulos : " & .Text
            PreencherModulos .Text
        End With
    End If
End Sub

Private Sub PrepararSistema()
    PreencherSistemas
    grdModulos.Preencher Bdados, ""
    grdFormularios.Preencher Bdados, ""
    txtCodSistema = ""
    txtSistema = ""
    txtDescSistema = ""
    txtCodSistema.SetFocus
End Sub


Private Function GravarSistema(Codigo As String, Nome As String, Descricao) As Boolean
    Dim Campos As String, Valores As String
    
    If Trim$(Nome) = "" Then Exit Function
    
    Campos = "TSI_COD_SISTEMA,TSI_NOME,TSI_DESCR"
    Valores = Bdados.PreparaValor(Codigo, Nome, Descricao)
    GravarSistema = Bdados.GravaDados("TAB_SISTEMA", Valores, Campos, "TSI_COD_SISTEMA='" & Codigo & "'")
End Function

Private Function ExcluirSistema(Codigo As String, Nome As String) As Boolean
    If Trim$(Nome) = "" Then Exit Function
    
    If Util.Confirma("Apagar " & Nome & " ?") Then
        ExcluirSistema = Bdados.DeletaDados("TAB_ACESSO_USUARIO", "TAU_TSI_COD_SISTEMA='" & Codigo & "'")
        'ExcluirSistema = ExcluirSistema And Bdados.DeletaDados("TAB_FORMULARIO", "TFO_TMO_COD_MODULO='" & Codigo & "'")
        ExcluirSistema = ExcluirSistema And Bdados.DeletaDados("TAB_MODULO", "TMO_TSI_COD_SISTEMA='" & Codigo & "'")
        ExcluirSistema = ExcluirSistema And Bdados.DeletaDados("TAB_SISTEMA", "TSI_COD_SISTEMA='" & Codigo & "'")
    End If
End Function

Private Function GravarModulo(Codigo As String, Nome As String, Descricao) As Boolean
    Dim Campos As String, Valores As String
    
    If Trim$(Nome) = "" Then Exit Function
    
    Campos = "TMO_TSI_COD_SISTEMA,TMO_COD_MODULO,TMO_NOME,TMO_DESCR"
    Valores = Bdados.PreparaValor(strCodSistema, Codigo, Nome, Descricao)
    GravarModulo = Bdados.GravaDados("TAB_MODULO", Valores, Campos, "TMO_TSI_COD_SISTEMA='" & strCodSistema & "' AND TMO_COD_MODULO='" & Codigo & "'")
End Function

Private Function ExcluirModulo(Codigo As String, Nome As String) As Boolean
    If Trim$(Nome) = "" Then Exit Function
    
    If Util.Confirma("Apagar " & Nome & " ?") Then
        ExcluirModulo = Bdados.DeletaDados("TAB_ACESSO_USUARIO", "TAU_TSI_COD_SISTEMA='" & strCodSistema & "' AND TAU_TMO_COD_MODULO='" & Codigo & "'")
        ExcluirModulo = ExcluirModulo And Bdados.DeletaDados("TAB_FORMULARIO", "TFO_TMO_COD_MODULO='" & Codigo & "'")
        ExcluirModulo = ExcluirModulo And Bdados.DeletaDados("TAB_MODULO", "TMO_TSI_COD_SISTEMA='" & strCodSistema & "' AND TMO_COD_MODULO='" & Codigo & "'")
    End If
End Function

Private Function GravarFormulario(Codigo As String, Nome As String, Descricao) As Boolean
    Dim Campos As String, Valores As String
    
    If Trim$(Nome) = "" Then Exit Function
    
    Campos = "TFO_TMO_COD_MODULO,TFO_COD_FORMULARIO,TFO_NOME,TFO_DESCR"
    Valores = Bdados.PreparaValor(strCodModulo, Codigo, Nome, Descricao)
    GravarFormulario = Bdados.GravaDados("TAB_FORMULARIO", Valores, Campos, "TFO_TMO_COD_MODULO='" & strCodModulo & "' AND TFO_COD_FORMULARIO='" & Codigo & "'")
End Function

Private Function ExcluirFormulario(Codigo As String, Nome As String) As Boolean
    If Trim$(Nome) = "" Then Exit Function
    
    If Util.Confirma("Apagar " & Nome & " ?") Then
        ExcluirFormulario = Bdados.DeletaDados("TAB_ACESSO_USUARIO", "TAU_TMO_COD_MODULO='" & strCodModulo & "' AND TAU_TFO_COD_FORMULARIO='" & Codigo & "'")
        ExcluirFormulario = ExcluirFormulario And Bdados.DeletaDados("TAB_FORMULARIO", "TFO_TMO_COD_MODULO='" & strCodModulo & "' AND TFO_COD_FORMULARIO='" & Codigo & "'")
    End If
End Function

Private Sub PrepararModulo()
    PreencherModulos strCodSistema
    grdFormularios.Preencher Bdados, ""
    txtCodModulo = ""
    txtModulo = ""
    txtDescModulo = ""
    txtCodModulo.SetFocus
End Sub

Private Sub PrepararFormulario()
    PreencherFormularios strCodModulo
    txtCodFormulario = ""
    txtFormulario = ""
    txtDescFormulario = ""
    txtCodFormulario.SetFocus
End Sub


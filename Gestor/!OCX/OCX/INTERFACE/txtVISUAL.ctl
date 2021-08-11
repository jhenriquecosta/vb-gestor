VERSION 5.00
Begin VB.UserControl txtVISUAL 
   BackColor       =   &H80000016&
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   2565
   ToolboxBitmap   =   "txtVISUAL.ctx":0000
   Begin VB.TextBox txtTexto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   630
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblRotulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rotulo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   540
   End
End
Attribute VB_Name = "txtVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'============================================
'PROPÓSITO txtVISUAL
'   Combinar um label com um textbox, formando um padrão de
'interface destes componentes na Visual Tecnologia
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 13:45
'============================================
'METODOS
'   CorTexto
'   CorRotulo
'   CorFundo
'   EnterEqvTab
'   AlinhamentoRotulo
'   Descricao
'   Requerido
'   Restricao
'   Formato
'   TipoLetras
'   AutoFocaliza
'   Text
'   Caption
'============================================
'CONTROLES
'   txtTexto
'   lblRotulo
'============================================
Option Explicit
Private Edita As VSTexto
Private Util As VSUtil
Private m_AutoFocaliza As Boolean
Private m_ValorPadrao As String
Private m_TipoLetras As vtCase
Private m_Formato As TipoFormato
Private m_Restricao As TipoChar
Private m_Requerido As Boolean
Private m_AlinhamentoRotulo As AlinhamtoLabel
Private m_AlinhamentoRotuloVertical As AlinhamtoVert
Private m_EnterEqvTab As Boolean
Public Enum vtCase
    letrTodas = 0
    letrMaiusculas = 1
    letrMinusculas = 2
End Enum
Public Enum TipoFormato
    formNenhum = -1
    formData = 0
    formCPF = 1
    formCGC = 2
    formTelefone = 3
    formCEP = 4
    formMonetario = 5
    formHora = 6
    formUmDigito = 7
    formDoisDigitos = 8
    formPASEP = 9
    formDocumento = 10
End Enum
Public Enum TipoChar
    restrNenhuma
    restrLetras
    restrNumeros
    restrValores
End Enum
Public Enum AlinhamtoLabel
    alinhEsquerdo
    alinhAcima
End Enum
Public Enum AlinhamtoVert
    alinhTopo
    alinhMeio
    alinhBase
End Enum
Public Event Change()
Public Event KeyPress(KeyAscii As Integer)
Private m_ValorMinimo As Long
Private m_ValorMaximo As Long
Private m_MinLen As Long
Private m_AgruparValores As Boolean
Private m_Mascara As String
Private m_CaracterSenha As String
Private m_RetirarMascara As Boolean
Private m_AutoTAB As Boolean

Public Property Get AutoTAB() As Boolean
    If bRegistrado Then AutoTAB = m_AutoTAB
End Property
Public Property Let AutoTAB(Value As Boolean)
    If bRegistrado Then
        m_AutoTAB = Value
        PropertyChanged "AutoTAB"
    End If
End Property

Public Property Get RetirarMascara() As Boolean
    If bRegistrado Then RetirarMascara = m_RetirarMascara
End Property
Public Property Let RetirarMascara(Value As Boolean)
    If bRegistrado Then
        m_RetirarMascara = Value
        PropertyChanged "RetirarMascara"
    End If
End Property
Public Property Get CaracterSenha() As String
    If bRegistrado Then CaracterSenha = m_CaracterSenha
End Property

Public Property Let CaracterSenha(ByVal Value As String)
    If bRegistrado Then
        m_CaracterSenha = Value
        txtTexto.PasswordChar = m_CaracterSenha
        PropertyChanged "CaracterSenha"
    End If
End Property

Public Property Get Mascara() As String
    If bRegistrado Then Mascara = m_Mascara
End Property

Public Property Let Mascara(ByVal Value As String)
    If bRegistrado Then
        m_Mascara = Value
        PropertyChanged "Mascara"
    End If
End Property

Public Property Get AgruparValores() As Boolean
    If bRegistrado Then AgruparValores = m_AgruparValores
End Property

Public Property Let AgruparValores(ByVal Value As Boolean)
    If bRegistrado Then
        m_AgruparValores = Value
        PropertyChanged "AgruparValores"
    End If
End Property


Public Property Get MinLen() As Long

    If bRegistrado Then MinLen = m_MinLen

End Property

Public Property Let MinLen(ByVal Value As Long)
' VTOcx.txtVISUAL.Property MinLen
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Comprimento do menor texto que o controle aceita
'
' Parametros : Value (Long) - Comprimento minimo
'
' Ex: MinLen = 6, o texto deve ter, no minimo, 06 caracteres
'--------------------------------------------------------------------------------

    If bRegistrado Then m_MinLen = Value

End Property

Public Property Get ValorMaximo() As Long

    If bRegistrado Then ValorMaximo = m_ValorMaximo

End Property

Public Property Let ValorMaximo(ByVal Value As Long)
' VTOcx.txtVISUAL.Property ValorMaximo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Maior valor que o conteudo pode assumir
'
' Parametros : Value (Long)
'
' Ex: ValorMaximo = 100, o usuario nao pode inserir o valor 101, p.ex.
'--------------------------------------------------------------------------------

    If bRegistrado Then m_ValorMaximo = Value

End Property

Public Property Get ValorMinimo() As Long

    If bRegistrado Then ValorMinimo = m_ValorMinimo

End Property

Public Property Let ValorMinimo(ByVal Value As Long)
' VTOcx.txtVISUAL.Property ValorMinimo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Menor valor que o conteudo pode assumir
'
' Parametros : Value (Long)
'
' Ex: ValorMinimo = 10, o controle nao aceita o valor 09, p.ex.
'--------------------------------------------------------------------------------

    If bRegistrado Then m_ValorMinimo = Value

End Property

Private Sub txtTexto_Change()
' VTOcx.txtVISUAL.Sub txtTexto_Change
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Dispara o evento Change do controle quando o conteudo da caixa de
'               texto é modificado
'
'--------------------------------------------------------------------------------

    If bRegistrado Then RaiseEvent Change

End Sub

Private Sub txtTexto_GotFocus()
' VTOcx.txtVISUAL.Sub txtTexto_GotFocus
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Destaca a caixa de texto automaticamente
'
'--------------------------------------------------------------------------------
    
    If bRegistrado Then
        If m_RetirarMascara Then
            txtTexto = Edita.TiraPic(txtTexto, ".")
            txtTexto = Edita.TiraPic(txtTexto, "-")
            txtTexto = Edita.TiraPic(txtTexto, "/")
            txtTexto = Edita.TiraPic(txtTexto, "(")
            txtTexto = Edita.TiraPic(txtTexto, ")")
        End If
        
        If m_AutoFocaliza Then
    
            txtTexto.BackColor = cteAmarelo
            Edita.SelecionaTexto txtTexto
        Else
    
            txtTexto.BackColor = cteBranco
    
        End If
    End If

End Sub

Private Sub txtTexto_KeyDown(KeyCode As Integer, Shift As Integer)
' VTOcx.txtVISUAL.Sub txtTexto_KeyDown
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Simula um updown no apertar das setas para cima e para baixo do teclado
'
' Parametros : KeyCode (Integer) - Tecla pressionada
'              Shift (Integer) - Pressionamento das teclas Shift, Ctrl e Alt
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        If MinLen <> MaxLen Then
    
            If IsNumeric(txtTexto) Then
    
                Select Case KeyCode
    
                    Case 38
                        txtTexto = CLng(txtTexto) + 1
    
                        If txtTexto > ValorMaximo Then txtTexto = ValorMaximo
    
                    Case 40
                        txtTexto = CLng(txtTexto) - 1
    
                        If txtTexto < ValorMinimo Then txtTexto = ValorMinimo
    
                End Select
    
            End If
    
        End If
    End If
End Sub

Private Sub txtTexto_KeyPress(KeyAscii As Integer)
' VTOcx.txtVISUAL.Sub txtTexto_KeyPress
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Formata e valida o caracter pressionado
'
' Parametros : KeyAscii (Integer) - Codigo ascii da tecla
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        RaiseEvent KeyPress(KeyAscii)
        If KeyAscii = 39 Then
            KeyAscii = 0
        End If
        If m_Restricao > 0 Then
    
            KeyAscii = Edita.AceitaDig(KeyAscii, m_Restricao - 1)
    
        End If
    
        Select Case TipoLetras
    
            Case letrMaiusculas
                KeyAscii = Edita.Maiuscula(KeyAscii)
    
            Case letrMinusculas
                KeyAscii = Edita.Minuscula(KeyAscii)
    
        End Select
    End If
End Sub

Private Sub txtTexto_LostFocus()
' VTOcx.txtVISUAL.Sub txtTexto_LostFocus
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Retira o destaque da caixa de texto e insere o valor padrao se a
'               caixa tiver sido deixada vazia
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        txtTexto = Util.Nvl(txtTexto, m_ValorPadrao)
        txtTexto_Validate False
        txtTexto.BackColor = cteBranco
        '    If m_Requerido And m_ValorPadrao = "" Then
        '        If Trim(Descricao) <> "" And Trim(txtTexto.Text) = "" Then
        '            MsgBox Descricao & ": informação requerida!", vbInformation, "VISUAL Tecnologia"
        '        End If
        '    End If
    End If
End Sub


Private Sub txtTexto_Validate(Cancel As Boolean)
' VTOcx.txtVISUAL.Sub txtTexto_Validate
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Retira espacos, formata e valida o texto
'
' Parametros : Cancel (Boolean) - Permite cancelar o LostFocus
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        txtTexto = Trim$(txtTexto)
    
        Formatar
        
        If ValorMinimo <> ValorMaximo Then
    
            If IsNumeric(txtTexto) Then
    
                If txtTexto > ValorMaximo Then txtTexto = ValorMaximo
    
                If txtTexto < ValorMinimo Then txtTexto = ValorMinimo
    
            End If
    
        End If
        
        If MinLen > 0 And Len(txtTexto) < MinLen And txtTexto <> "" Then
    
            Cancel = True
    
        Else
    
            Edita.DestacaCaixa txtTexto, False
    
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
' VTOcx.txtVISUAL.Sub UserControl_Initialize
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Incializa as variaveis de ambiente
'
'--------------------------------------------------------------------------------

    ValidaComponente "INTERFACE"
'    If bRegistrado Then
        Set Edita = New VSTexto
        Set Util = New VSUtil
'    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
' VTOcx.txtVISUAL.Sub UserControl_KeyPress
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Simula o Tab no pressionamento do Enter
'
' Parametros : KeyAscii (Integer) - Tecla pressionada
'
'--------------------------------------------------------------------------------

    If bRegistrado Then
        If m_EnterEqvTab Then
            If AutoTAB Then
                If Len(txtTexto) = txtTexto.MaxLength - 1 Then SendKeys "{TAB}"
            Else
                If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
            End If
        End If
    End If
End Sub

Private Sub UserControl_LostFocus()
    'If bRegistrado Then txtTexto_Validate False
End Sub

'============================================
'PROPÓSITO UserControl_Resize
'   Redimensionar a caixa de texto de forma que o controle
'possa comportar todos os seus componentes
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 13:47
'============================================
Private Sub UserControl_Resize()
' VTOcx.txtVISUAL.Sub UserControl_Resize
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Dim Tam As Integer
        lblRotulo.Left = 0
    
        Select Case m_AlinhamentoRotulo
    
            Case alinhEsquerdo
                txtTexto.Top = 0
    
                If Trim$(lblRotulo) = "" Then
    
                    txtTexto.Left = 0
                    Tam = Width
    
                Else
    
                    txtTexto.Left = lblRotulo.Width + 40
                    Tam = Width - lblRotulo.Width - 40
    
                End If
    
                If Tam > 0 Then
    
                    txtTexto.Width = Tam
    
                Else
    
                    txtTexto.Width = 0
    
                End If
    
                txtTexto.Height = Height
    
                If m_AlinhamentoRotuloVertical = alinhMeio Then
    
                    lblRotulo.Top = (Height - lblRotulo.Height) / 2
    
                ElseIf m_AlinhamentoRotuloVertical = alinhTopo Then
    
                    lblRotulo.Top = 0
    
                ElseIf m_AlinhamentoRotuloVertical = alinhBase Then
    
                    lblRotulo.Top = txtTexto.Height - lblRotulo.Height
    
                End If
                
            Case alinhAcima
                lblRotulo.Top = 0
                txtTexto.Left = 0
                txtTexto.Top = lblRotulo.Height
                txtTexto.Height = Height - lblRotulo.Height
                Tam = Width - 40
    
                If Tam > 0 Then
    
                    txtTexto.Width = Tam
    
                Else
    
                    txtTexto.Width = 0
    
                End If
    
        End Select
    End If
End Sub

'============================================
'PROPÓSITO Caption
'   Prover acesso ao conteudo do lblRotulo
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:06
'============================================
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Conteudo do rotulo"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property Caption
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then Caption = lblRotulo.Caption

End Property

Public Property Let Caption(ByVal vnewvalue As String)
' VTOcx.txtVISUAL.Property Caption
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (String)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        lblRotulo.Caption = vnewvalue
        UserControl_Resize
    End If
End Property

'============================================
'PROPÓSITO Text
'   Prover acesso ao conteudo do txtTexto
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:06
'============================================
Public Property Get Text() As String
Attribute Text.VB_Description = "Conteudo da caixa de texto"
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Visual"
Attribute Text.VB_UserMemId = 0
' VTOcx.txtVISUAL.Property Text
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then Text = txtTexto.Text

End Property

Public Property Let Text(ByVal vnewvalue As String)
' VTOcx.txtVISUAL.Property Text
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (String)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Select Case TipoLetras
    
            Case letrMaiusculas
                vnewvalue = UCase$(vnewvalue)
    
            Case letrMinusculas
                vnewvalue = LCase$(vnewvalue)
    
        End Select
    
        txtTexto.Text = vnewvalue
        txtTexto_GotFocus
        txtTexto_Validate False
    End If
End Property

Public Property Get Enabled() As Boolean

    If bRegistrado Then Enabled = txtTexto.Enabled

End Property

Public Property Let Enabled(ByVal vnewvalue As Boolean)
' VTOcx.txtVISUAL.Property Enabled
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:25
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (Boolean)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        txtTexto.Enabled = vnewvalue
        'lblRotulo.Enabled = vnewvalue
        txtTexto.TabStop = vnewvalue
    End If
End Property

'============================================
'PROPÓSITO UserControl_InitProperties
'   Iniciar os valores padrões das propriedades
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:43
'============================================
Private Sub UserControl_InitProperties()
' VTOcx.txtVISUAL.Sub UserControl_InitProperties
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Text = ""
        
        AlinhamentoRotulo = alinhEsquerdo
        AlinhamentoRotuloVertical = alinhMeio
        AlinhamentoTexto = AlignmentConstants.vbLeftJustify
        AutoFocaliza = True
        Caption = "Rotulo"
        corFundo = UserControl.Ambient.BackColor
        CorRotulo = &H80000012
        corTexto = &H800000
        '    Descricao = "Rotulo"
        EnterEqvTab = True
        Enabled = True
        Formato = formNenhum
        Restricao = restrNenhuma
        TipoLetras = letrMaiusculas
        Requerido = True
        MaxLen = 0
        MinLen = 0
        ValorMinimo = 0
        ValorMaximo = 0
        
        Width = lblRotulo.Width + 40 + txtTexto.Width
    
        AgruparValores = True
        Mascara = ""
        CaracterSenha = ""
        RetirarMascara = True
        AutoTAB = False
    End If
End Sub

'============================================
'PROPÓSITO UserControl_ReadProperties
'   Atualizar os componentes do controle com os valores das
'propriedades
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:44
'============================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' VTOcx.txtVISUAL.Sub UserControl_ReadProperties
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : PropBag (PropertyBag)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Caption = PropBag.ReadProperty("Caption", "Rotulo")
        Text = PropBag.ReadProperty("Text", "Texto")
        Enabled = PropBag.ReadProperty("Enabled", True)
        AutoFocaliza = PropBag.ReadProperty("AutoFocaliza", True)
        TipoLetras = PropBag.ReadProperty("TipoLetras", letrMaiusculas)
        Formato = PropBag.ReadProperty("Formato", formNenhum)
        Restricao = PropBag.ReadProperty("Restricao", restrNenhuma)
        Requerido = PropBag.ReadProperty("Requerido", True)
        AlinhamentoRotulo = PropBag.ReadProperty("AlinhamentoRotulo", alinhEsquerdo)
        AlinhamentoRotuloVertical = PropBag.ReadProperty("AlinhamentoRotuloVertical", alinhMeio)
        AlinhamentoTexto = PropBag.ReadProperty("AlinhamentoTexto", AlignmentConstants.vbLeftJustify)
        EnterEqvTab = PropBag.ReadProperty("EnterEqvTab", True)
        corFundo = PropBag.ReadProperty("CorFundo", UserControl.Ambient.BackColor)
        CorRotulo = PropBag.ReadProperty("CorRotulo", &H80000012)
        corTexto = PropBag.ReadProperty("CorTexto", &H800000)
        ValorPadrao = PropBag.ReadProperty("ValorPadrao", "")
        MaxLen = PropBag.ReadProperty("MaxLen", 0)
        MinLen = PropBag.ReadProperty("MinLen", 0)
        ValorMinimo = PropBag.ReadProperty("ValorMinimo", 0)
        ValorMaximo = PropBag.ReadProperty("ValorMaximo", 0)
        AgruparValores = PropBag.ReadProperty("AgruparValores", True)
        Mascara = PropBag.ReadProperty("Mascara", "")
        CaracterSenha = PropBag.ReadProperty("CaracterSenha", "")
        RetirarMascara = PropBag.ReadProperty("RetirarMascara", True)
        AutoTAB = PropBag.ReadProperty("AutoTAB", False)
    End If
End Sub

'============================================
'PROPÓSITO UserControl_Terminate
'   Destruir as variaveis de memoria utilizadas
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:45
'============================================
Private Sub UserControl_Terminate()
' VTOcx.txtVISUAL.Sub UserControl_Terminate
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Set Edita = Nothing
        Set Util = Nothing
    End If
End Sub

'============================================
'PROPÓSITO UserControl_WriteProperties
'   Escrever nas propriedades os valores dos componentes
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:45
'============================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' VTOcx.txtVISUAL.Sub UserControl_WriteProperties
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : PropBag (PropertyBag)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        Call PropBag.WriteProperty("Caption", lblRotulo.Caption, "Rotulo")
        Call PropBag.WriteProperty("Text", txtTexto.Text, "Texto")
        Call PropBag.WriteProperty("Enabled", txtTexto.Enabled, True)
        Call PropBag.WriteProperty("AutoFocaliza", m_AutoFocaliza, True)
        Call PropBag.WriteProperty("TipoLetras", m_TipoLetras, letrMaiusculas)
        Call PropBag.WriteProperty("Formato", m_Formato, formNenhum)
        Call PropBag.WriteProperty("Restricao", m_Restricao, restrNenhuma)
        Call PropBag.WriteProperty("Requerido", m_Requerido, True)
        Call PropBag.WriteProperty("Descricao", txtTexto.Tag, "")
        Call PropBag.WriteProperty("AlinhamentoRotulo", m_AlinhamentoRotulo, alinhEsquerdo)
        Call PropBag.WriteProperty("AlinhamentoRotuloVertical", m_AlinhamentoRotuloVertical, alinhMeio)
        Call PropBag.WriteProperty("AlinhamentoTexto", txtTexto.Alignment, AlignmentConstants.vbLeftJustify)
        Call PropBag.WriteProperty("EnterEqvTab", m_EnterEqvTab, True)
        Call PropBag.WriteProperty("CorFundo", BackColor, UserControl.Ambient.BackColor)
        Call PropBag.WriteProperty("CorRotulo", lblRotulo.ForeColor, &H80000012)
        Call PropBag.WriteProperty("CorTexto", txtTexto.ForeColor, &H800000)
        Call PropBag.WriteProperty("ValorPadrao", m_ValorPadrao, "")
        Call PropBag.WriteProperty("ValorMinimo", m_ValorMinimo, 0)
        Call PropBag.WriteProperty("ValorMaximo", m_ValorMaximo, 0)
        Call PropBag.WriteProperty("MaxLen", txtTexto.MaxLength, 0)
        Call PropBag.WriteProperty("MinLen", m_MinLen, 0)
    
        Call PropBag.WriteProperty("AgruparValores", m_AgruparValores, True)
        Call PropBag.WriteProperty("Mascara", m_Mascara, "")
        Call PropBag.WriteProperty("CaracterSenha", m_CaracterSenha, "")
        Call PropBag.WriteProperty("RetirarMascara", m_RetirarMascara, True)
        Call PropBag.WriteProperty("AutoTAB", m_AutoTAB, False)
        
    End If
End Sub

'============================================
'PROPÓSITO AutoFocaliza
'   Setar flag para chamada automatica de VSTexto.FocalizaTexto
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:07
'============================================
Public Property Get AutoFocaliza() As Boolean
Attribute AutoFocaliza.VB_Description = "Focaliza automaticamente a caixa de texto que recebe o foco"
Attribute AutoFocaliza.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property AutoFocaliza
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then AutoFocaliza = m_AutoFocaliza

End Property

Public Property Let AutoFocaliza(ByVal vnewvalue As Boolean)
' VTOcx.txtVISUAL.Property AutoFocaliza
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (Boolean)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then m_AutoFocaliza = vnewvalue

End Property

'============================================
'PROPÓSITO TipoLetras
'   Definir flag para chamada de VSTexto.AceitaMaiuscula
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 14:50
'============================================
Public Property Get TipoLetras() As vtCase
Attribute TipoLetras.VB_Description = "Case das letras pressionadas"
Attribute TipoLetras.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property TipoLetras
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then TipoLetras = m_TipoLetras

End Property

Public Property Let TipoLetras(ByVal vnewvalue As vtCase)
' VTOcx.txtVISUAL.Property TipoLetras
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (vtCase)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_TipoLetras = vnewvalue
    
        Select Case vnewvalue
    
            Case letrMaiusculas
                Text = UCase$(Text)
    
            Case letrMinusculas
                Text = LCase$(Text)
    
        End Select
    End If
End Property

'============================================
'PROPÓSITO Formato
'   Definir flag para formatacao do texto
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 15:14
'============================================
Public Property Get Formato() As TipoFormato
Attribute Formato.VB_Description = "Tipo do formato a ser aplicado na caixa de texto"
Attribute Formato.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property Formato
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then Formato = m_Formato

End Property

Public Property Let Formato(ByVal vnewvalue As TipoFormato)
' VTOcx.txtVISUAL.Property Formato
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (TipoFormato)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_Formato = vnewvalue
        Formatar
    End If

End Property

'============================================
'PROPÓSITO Restricao
'   Definir flag para aceitação de teclas pressionadas
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 15:27
'============================================
Public Property Get Restricao() As TipoChar
Attribute Restricao.VB_Description = "Restringe as teclas que podem ser pressionadas na caixa"
Attribute Restricao.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property Restricao
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then Restricao = m_Restricao

End Property

Public Property Let Restricao(ByVal vnewvalue As TipoChar)
' VTOcx.txtVISUAL.Property Restricao
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (TipoChar)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then m_Restricao = vnewvalue

End Property

'============================================
'PROPÓSITO Requerido
'   Flag que indica se o campo é requerido ou não
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 16:08
'============================================
Public Property Get Requerido() As Boolean
Attribute Requerido.VB_Description = "Informacao do caixa de texto é requerida"
Attribute Requerido.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property Requerido
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then Requerido = m_Requerido

End Property

Public Property Let Requerido(ByVal vnewvalue As Boolean)
' VTOcx.txtVISUAL.Property Requerido
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (Boolean)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then m_Requerido = vnewvalue
    '    If crit Then
    '        If Descricao = "" Then Descricao = txtTexto.Name
    '    Else
    '        Descricao = ""
    '    End If

End Property

'============================================
'PROPÓSITO Descricao
'   Interfacear a propriedade Tag da caixa de texto
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 16:12
'============================================
'Public Property Get Descricao() As String
'    Descricao = txtTexto.Tag
'End Property
'
'Public Property Let Descricao(ByVal vNewValue As String)
'    txtTexto.Tag = vNewValue
'End Property

'============================================
'PROPÓSITO AlinhamentoRotulo
'   Definir a posicao do label em relacao à caixa
'de texto
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 16:39
'============================================
Public Property Get AlinhamentoRotulo() As AlinhamtoLabel
Attribute AlinhamentoRotulo.VB_Description = "Alinhamento do rotulo em relacao à caixa de texto"
Attribute AlinhamentoRotulo.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property AlinhamentoRotulo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then AlinhamentoRotulo = m_AlinhamentoRotulo

End Property

Public Property Let AlinhamentoRotulo(ByVal vnewvalue As AlinhamtoLabel)
' VTOcx.txtVISUAL.Property AlinhamentoRotulo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (AlinhamtoLabel)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_AlinhamentoRotulo = vnewvalue
    
        Select Case m_AlinhamentoRotulo
    
            Case alinhEsquerdo
                lblRotulo.Alignment = AlignmentConstants.vbRightJustify
    
            Case alinhAcima
                lblRotulo.Alignment = AlignmentConstants.vbLeftJustify
    
        End Select
    
        UserControl_Resize
    End If
End Property

'============================================
'PROPÓSITO EnterEqvTab
'   Definir se o pressionamento da tecla {Enter}
'é equivalente ao pressionar da tecla {Tab}
'============================================
'Autor: Sergio Queiroz
'Data: 26/11/2001 17:39
'============================================
Public Property Get EnterEqvTab() As Boolean
Attribute EnterEqvTab.VB_Description = "Define se pressionar a tecla {Enter} é equivalente a pressionar a tecla {Tab}"
Attribute EnterEqvTab.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property EnterEqvTab
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then EnterEqvTab = m_EnterEqvTab

End Property

Public Property Let EnterEqvTab(ByVal vnewvalue As Boolean)
' VTOcx.txtVISUAL.Property EnterEqvTab
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (Boolean)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then m_EnterEqvTab = vnewvalue

End Property

Public Property Get corFundo() As OLE_COLOR
Attribute corFundo.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property CorFundo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then corFundo = BackColor

End Property

Public Property Let corFundo(ByVal vnewvalue As OLE_COLOR)
' VTOcx.txtVISUAL.Property CorFundo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (OLE_COLOR)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then BackColor = vnewvalue

End Property

Public Property Get CorRotulo() As OLE_COLOR
Attribute CorRotulo.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property CorRotulo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then CorRotulo = lblRotulo.ForeColor

End Property

Public Property Let CorRotulo(ByVal vnewvalue As OLE_COLOR)
' VTOcx.txtVISUAL.Property CorRotulo
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (OLE_COLOR)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then lblRotulo.ForeColor = vnewvalue

End Property

Public Property Get corTexto() As OLE_COLOR
Attribute corTexto.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property CorTexto
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then corTexto = txtTexto.ForeColor

End Property

Public Property Let corTexto(ByVal vnewvalue As OLE_COLOR)
' VTOcx.txtVISUAL.Property CorTexto
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (OLE_COLOR)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then txtTexto.ForeColor = vnewvalue

End Property

Public Property Get ValorPadrao() As String
Attribute ValorPadrao.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property ValorPadrao
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then ValorPadrao = m_ValorPadrao

End Property

Public Property Let ValorPadrao(ByVal vnewvalue As String)
' VTOcx.txtVISUAL.Property ValorPadrao
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (String)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then m_ValorPadrao = vnewvalue

End Property

Public Function Validar() As Boolean
' VTOcx.txtVISUAL.Function Validar
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        On Error Resume Next
    
        If m_Requerido Then
    
            Validar = (txtTexto <> "")
    
            If Not Validar Then txtTexto.SetFocus
    
        End If
    End If
End Function

Public Property Get AlinhamentoTexto() As AlignmentConstants
Attribute AlinhamentoTexto.VB_Description = "Define o posicionamento do rotulo em relação à caixa de texto"
Attribute AlinhamentoTexto.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property AlinhamentoTexto
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then AlinhamentoTexto = txtTexto.Alignment

End Property

Public Property Let AlinhamentoTexto(ByVal vnewvalue As AlignmentConstants)
' VTOcx.txtVISUAL.Property AlinhamentoTexto
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (AlignmentConstants)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then txtTexto.Alignment = vnewvalue

End Property

Public Property Get AlinhamentoRotuloVertical() As AlinhamtoVert
Attribute AlinhamentoRotuloVertical.VB_Description = "Define o posicionamento vertical do rotulo em relacao à caixa de texto"
Attribute AlinhamentoRotuloVertical.VB_ProcData.VB_Invoke_Property = ";Visual"
' VTOcx.txtVISUAL.Property AlinhamentoRotuloVertical
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then AlinhamentoRotuloVertical = m_AlinhamentoRotuloVertical

End Property

Public Property Let AlinhamentoRotuloVertical(ByVal vnewvalue As AlinhamtoVert)
' VTOcx.txtVISUAL.Property AlinhamentoRotuloVertical
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (AlinhamtoVert)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then
        m_AlinhamentoRotuloVertical = vnewvalue
        UserControl_Resize
    End If
End Property

Public Property Get MaxLen() As Long
' VTOcx.txtVISUAL.Property MaxLen
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then MaxLen = txtTexto.MaxLength

End Property

Public Property Let MaxLen(ByVal vnewvalue As Long)
' VTOcx.txtVISUAL.Property MaxLen
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:24
'
' Descricao  : Descreva
'
' Parametros : vnewvalue (Long)
'
' Ex:
'--------------------------------------------------------------------------------

    If bRegistrado Then txtTexto.MaxLength = vnewvalue

End Property


Private Sub Formatar()
    If bRegistrado Then
        If Formato = formNenhum Then
            If Mascara <> "" Then
            Dim msk As String
            Dim Posicoes() As Integer
            Dim pos As Integer, i As Integer, inicio As Integer
            
            '1 - Mapeia as ocorrencias de ponto na mascara
            msk = Mascara
            inicio = 1: i = 0
            Do
                pos = InStr(inicio, msk, ".")
                ReDim Preserve Posicoes(0 To i)
                Posicoes(i) = pos
                i = i + 1
                inicio = pos + 1
            Loop While pos <> 0
            
            '2 - Aplica a mascara sem pontos
            msk = Edita.TiraPic(msk, ".")
            txtTexto = Format$(txtTexto, msk)
            
            '3 - Aplica os pontos no resultado
            For pos = 0 To i - 1
                If Posicoes(pos) > 0 Then txtTexto = Edita.BotaPic(txtTexto, ".", Posicoes(pos) - 1)
            Next
            End If
        Else
            txtTexto = Edita.TiraPic(txtTexto, ".")
            txtTexto = Edita.TiraPic(txtTexto, "-")
            txtTexto = Edita.TiraPic(txtTexto, "/")
            txtTexto = Edita.TiraPic(txtTexto, "(")
            txtTexto = Edita.TiraPic(txtTexto, ")")
            txtTexto = Edita.FormataTexto(txtTexto, m_Formato, m_AgruparValores)
        End If
    End If
End Sub

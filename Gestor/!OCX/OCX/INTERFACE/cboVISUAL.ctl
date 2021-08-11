VERSION 5.00
Begin VB.UserControl cboVISUAL 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   3510
   ToolboxBitmap   =   "cboVISUAL.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1170
      Top             =   0
   End
   Begin VB.ComboBox cboCombo 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "cboVISUAL.ctx":0312
      Left            =   705
      List            =   "cboVISUAL.ctx":0314
      TabIndex        =   0
      Text            =   "cboCombo"
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label lblRotulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rotulo"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   540
   End
End
Attribute VB_Name = "cboVISUAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private bTravaEventoClick As Boolean
Private colDados As Collection
Private Edita As VSTexto
Private autofoc As Boolean
Private padr As String
Private digito As TipoChar
Private crit As Boolean
Private alinh As AlinhamtoLabel
Private sndkys As Boolean
Private vtformato As TipoFormato
Private cas As vtCase
Private editav As Boolean
Public Event Click()
Public Type TipoColuna
    Ordem As Integer
    Nome As String
    Valor As Variant
    ListIndex As Integer
End Type
'******* Flat ***********
' keep the combobox state (dropped or not)
Private cbOpen As Boolean
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' build the mask for the control
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
' the control's styles, according to the mouse cursor position
Private Enum DrawCombo
    FC_DRAWNORMAL = 0
    FC_DRAWRAISED = 1
    FC_DRAWPRESSED = 2
    FC_DRAWDISABLED = 3
End Enum
' the style of the pen used to create the mask over the combobox
Private Const SM_CXHTHUMB = 10
Private Const PS_SOLID = 0

Public Property Get NewIndex() As Integer

    On Error GoTo Trata

1   If bRegistrado Then NewIndex = cboCombo.NewIndex
    Exit Property
Trata:
End Property

Public Property Get ListCount() As Integer
    On Error GoTo Trata

1   If bRegistrado Then ListCount = cboCombo.ListCount
    Exit Property
Trata:
End Property

Private Sub cboCombo_Click()
    On Error GoTo Trata

1   If bRegistrado Then
2       If Not bTravaEventoClick Then RaiseEvent Click
3   End If
    Exit Sub
Trata:
End Sub

Private Sub cboCombo_KeyPress(KeyAscii As Integer)
    On Error GoTo Trata

1   If bRegistrado Then
2       Edita.BuscaItemNaLista cboCombo, KeyAscii
3       KeyAscii = IIf(editav, KeyAscii, 0)

4       If digito > 0 Then
5           KeyAscii = Edita.AceitaDig(KeyAscii, digito - 1)
6       End If

7       Select Case TipoLetras

            Case letrMaiusculas
8               KeyAscii = Edita.Maiuscula(KeyAscii)

9           Case letrMinusculas
10              KeyAscii = Edita.Minuscula(KeyAscii)
11      End Select

12  End If
    Exit Sub
Trata:
End Sub

Private Sub cboCombo_LostFocus()
    On Error GoTo Trata

1   If bRegistrado Then

2       If cboCombo.Style = 0 Then cboCombo = Util.Nvl(cboCombo, padr)
3   End If

    Exit Sub
Trata:
End Sub

Private Sub cboCombo_Validate(Cancel As Boolean)
    On Error GoTo Trata

1   If bRegistrado Then
2       Edita.DestacaCaixa cboCombo, False
3       Dim pres As Boolean

4       Select Case vtformato

            Case formMonetario, formTelefone: pres = True
5       End Select

6       If vtformato > 0 Then cboCombo = Edita.FormataTexto(cboCombo, vtformato, pres)
7   End If
    Exit Sub
Trata:
End Sub

Private Sub Timer1_Timer()
    On Error GoTo Trata

1   If bRegistrado Then

2       If Ambient.UserMode Then   ' do not execute in Design mode
3           Dim pnt As POINTAPI
4           GetCursorPos pnt        ' get the cursor pos
            ' convert the coords relatively to the control
5           ScreenToClient UserControl.hWnd, pnt
            ' check if the cursor is over the control

6           If pnt.x * Screen.TwipsPerPixelX < UserControl.ScaleLeft Or _
               pnt.y * Screen.TwipsPerPixelX < UserControl.ScaleTop Or _
               pnt.x * Screen.TwipsPerPixelX > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
               pnt.y * Screen.TwipsPerPixelX > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
                ' the cursor is not over the control
                ' Debug.Print "Out"

7               If cboCombo.Enabled Then
                    ' get the Dropped state of the combobox
8                   cbOpen = SendMessageAsLong(cboCombo.hWnd, CB_GETDROPPEDSTATE, 0, 0)

9                   If Not cbOpen Then DrawCombo FC_DRAWNORMAL

10              Else
                    ' draw the disabled mask is the control is disabled
11                  DrawCombo FC_DRAWDISABLED
12              End If

13          Else
                ' cursor over the control
                ' Debug.Print "In"
                ' draw the raised mask
14              DrawCombo FC_DRAWRAISED
15          End If

16      End If
        Timer1.Enabled = False

17  End If
    Exit Sub
Trata:
End Sub

Private Sub UserControl_Initialize()
    On Error GoTo Trata
1   ValidaComponente "INTERFACE"

'2   If bRegistrado Then
3       Set Edita = New VSTexto
4       Set Util = New VSUtil
'5   End If
    Exit Sub
Trata:
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error GoTo Trata
1   If bRegistrado Then

2       If sndkys Then

3           If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
4       End If

5   End If
    Exit Sub
Trata:
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Trata
1   Dim Tam As Integer

2   If bRegistrado Then
3       lblRotulo.Left = 0

4       Select Case alinh

            Case alinhEsquerdo
5               Height = cboCombo.Height
6               cboCombo.Top = (Height - cboCombo.Height) / 2

7               If Trim$(lblRotulo) = "" Then
8                   cboCombo.Left = 0
9                   Tam = Width

10              Else
11                  cboCombo.Left = lblRotulo.Width + 40
12                  Tam = Width - lblRotulo.Width - 40
13              End If

14              If Tam > 0 Then
15                  cboCombo.Width = Tam

16              Else
17                  cboCombo.Width = 0
18              End If

                'cboCombo.Height = Height
19              lblRotulo.Top = (Height - lblRotulo.Height) / 2

20          Case alinhAcima
21              lblRotulo.Top = 0
22              cboCombo.Left = 0
23              cboCombo.Top = lblRotulo.Height

24              If Height <> lblRotulo.Height + cboCombo.Height Then
25                  Height = lblRotulo.Height + cboCombo.Height
26              End If

27              Tam = Width - 40

28              If Tam > 0 Then
29                  cboCombo.Width = Tam

30              Else
31                  cboCombo.Width = 0
32              End If

33      End Select

34  End If
    Exit Sub
Trata:
End Sub

Public Property Get Caption() As String
    On Error GoTo Trata
1   If bRegistrado Then Caption = lblRotulo.Caption
    Exit Property
Trata:
End Property

Public Property Let Caption(ByVal vnewvalue As String)
    On Error GoTo Trata
1   If bRegistrado Then
2       lblRotulo.Caption = vnewvalue
3       UserControl_Resize
4   End If
    Exit Property
Trata:
End Property

Public Property Get ToolTipText() As String
    On Error GoTo Trata
1   If bRegistrado Then
        If Trim(cboCombo.ToolTipText) = "" Then
            If Trim(cboCombo.Tag) <> "" Then
                ToolTipText = "Campo Obrigatório: Selecione " & lblRotulo.Caption
            Else
                ToolTipText = "Selecione " & lblRotulo.Caption
            End If
        Else
            ToolTipText = cboCombo.ToolTipText
        End If
    End If
    Exit Property
Trata:
End Property

Public Property Let ToolTipText(ByVal vnewvalue As String)
    On Error GoTo Trata
1   If bRegistrado Then
2       cboCombo.ToolTipText = vnewvalue
3       UserControl_Resize
4   End If
    Exit Property
Trata:
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    On Error GoTo Trata
1   If bRegistrado Then Text = cboCombo.Text
    Exit Property
Trata:
End Property

Public Property Let Text(ByVal vnewvalue As String)
    On Error GoTo Trata
1   If bRegistrado Then

2       Select Case TipoLetras

            Case letrMaiusculas
3               vnewvalue = UCase$(vnewvalue)

4           Case letrMinusculas
5               vnewvalue = LCase$(vnewvalue)
6       End Select

        'cboCombo_Validate False
7       cboCombo.Text = vnewvalue

8       If vnewvalue = "" Then
9           cboCombo.ListIndex = -1

10      Else
11          Dim i As Integer

12          For i = 0 To cboCombo.ListCount

13              If cboCombo.List(i) = vnewvalue Then
14                  cboCombo.ListIndex = i
15                  Exit For
16              End If

17          Next

18      End If

19  End If

    Exit Property
Trata:
End Property

Private Sub UserControl_InitProperties()
    On Error GoTo Trata
1   If bRegistrado Then
2       Text = ""
3       Alinhamento = alinhEsquerdo
4       Enabled = True
        '    AutoFocaliza = True
5       Caption = "Rotulo"
6       corFundo = UserControl.Ambient.BackColor
7       CorRotulo = &H80000012
8       corTexto = &H800000
        '    Descricao = "Rotulo"
9       EnterEqvTab = True
10      Formato = formNenhum
11      Restricao = restrNenhuma
12      TipoLetras = letrMaiusculas
13      Requerido = True
14      Editavel = False
15      Width = lblRotulo.Width + 40 + cboCombo.Width
16  End If

    Exit Sub
Trata:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo Trata
1   If bRegistrado Then
2       lblRotulo.Caption = PropBag.ReadProperty("Caption", "Rotulo")

3       If cboCombo.Style = 0 Then cboCombo.Text = PropBag.ReadProperty("Text", "Texto")
4       autofoc = PropBag.ReadProperty("AutoFocaliza", True)
5       cas = PropBag.ReadProperty("TipoLetras", letrMaiusculas)
6       vtformato = PropBag.ReadProperty("Formato", formNenhum)
7       digito = PropBag.ReadProperty("Restricao", restrNenhuma)
8       crit = PropBag.ReadProperty("Requerido", True)
9       cboCombo.Tag = PropBag.ReadProperty("Descricao", "")
10      alinh = PropBag.ReadProperty("Alinhamento", alinhEsquerdo)
11      sndkys = PropBag.ReadProperty("EnterEqvTab", True)
12      BackColor = PropBag.ReadProperty("CorFundo", UserControl.Ambient.BackColor)
13      lblRotulo.ForeColor = PropBag.ReadProperty("CorRotulo", &H80000012)
14      cboCombo.ForeColor = PropBag.ReadProperty("CorTexto", &H800000)
15      padr = PropBag.ReadProperty("ValorPadrao", "")
16      editav = PropBag.ReadProperty("Editavel", False)
17      cboCombo.Enabled = PropBag.ReadProperty("Enabled", True)
18      lblRotulo.Enabled = PropBag.ReadProperty("Enabled", True)
19  End If
    Exit Sub
Trata:
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Trata
1   If bRegistrado Then
2       Set Edita = Nothing
3       Set Util = Nothing
4   End If
    Exit Sub
Trata:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo Trata
1   If bRegistrado Then
2       Call PropBag.WriteProperty("Caption", lblRotulo.Caption, "Rotulo")
3       Call PropBag.WriteProperty("Text", cboCombo.Text, "Texto")
4       Call PropBag.WriteProperty("AutoFocaliza", autofoc, True)
5       Call PropBag.WriteProperty("TipoLetras", cas, letrMaiusculas)
6       Call PropBag.WriteProperty("Formato", vtformato, formNenhum)
7       Call PropBag.WriteProperty("Restricao", digito, restrNenhuma)
8       Call PropBag.WriteProperty("Requerido", crit, True)
9       Call PropBag.WriteProperty("Descricao", cboCombo.Tag, "")
10      Call PropBag.WriteProperty("Alinhamento", alinh, alinhEsquerdo)
11      Call PropBag.WriteProperty("EnterEqvTab", sndkys, True)
12      Call PropBag.WriteProperty("CorFundo", BackColor, UserControl.Ambient.BackColor)
13      Call PropBag.WriteProperty("CorRotulo", lblRotulo.ForeColor, &H80000012)
14      Call PropBag.WriteProperty("CorTexto", cboCombo.ForeColor, &H800000)
15      Call PropBag.WriteProperty("ValorPadrao", padr, "")
16      Call PropBag.WriteProperty("Editavel", editav, False)
17      Call PropBag.WriteProperty("Enabled", cboCombo.Enabled, True)
18  End If
    Exit Sub
Trata:
End Sub

'Public Property Get AutoFocaliza() As Boolean
'    AutoFocaliza = autofoc
'End Property
'
'Public Property Let AutoFocaliza(ByVal vNewValue As Boolean)
'    autofoc = vNewValue
'End Property
Public Property Get TipoLetras() As vtCase
    On Error GoTo Trata
1   If bRegistrado Then TipoLetras = cas
    Exit Property
Trata:
End Property

Public Property Let TipoLetras(ByVal vnewvalue As vtCase)
    On Error GoTo Trata
1   If bRegistrado Then
2       cas = vnewvalue

3       Select Case vnewvalue

            Case letrMaiusculas
4               Text = UCase$(Text)

5           Case letrMinusculas
6               Text = LCase$(Text)
7       End Select

8   End If
    Exit Property
Trata:
End Property

Public Property Get Formato() As TipoFormato
    On Error GoTo Trata
1   If bRegistrado Then Formato = vtformato
    Exit Property
Trata:
End Property

Public Property Let Formato(ByVal vnewvalue As TipoFormato)
    On Error GoTo Trata
1   Dim pres As Boolean

2   If bRegistrado Then
3       vtformato = vnewvalue

4       Select Case vnewvalue

            Case formMonetario, formTelefone: pres = True
5       End Select

6       cboCombo = Edita.FormataTexto(cboCombo, vnewvalue, pres)
7   End If

    Exit Property
Trata:
End Property

Public Property Get Restricao() As TipoChar
    On Error GoTo Trata
1   If bRegistrado Then Restricao = digito
    Exit Property
Trata:
End Property

Public Property Let Restricao(ByVal vnewvalue As TipoChar)
    On Error GoTo Trata

1   If bRegistrado Then digito = vnewvalue
    Exit Property
Trata:
End Property

Public Property Get Requerido() As Boolean
    On Error GoTo Trata

1   If bRegistrado Then Requerido = crit
    Exit Property
Trata:
End Property

Public Property Let Requerido(ByVal vnewvalue As Boolean)
    On Error GoTo Trata

1   If bRegistrado Then crit = vnewvalue
    '    If crit Then
    '        If Descricao = "" Then Descricao = cboCombo.Name
    '    Else
    '        Descricao = ""
    '    End If
    Exit Property
Trata:
End Property

'Public Property Get Descricao() As String
'    Descricao = cboCombo.Tag
'End Property
'
'Public Property Let Descricao(ByVal vNewValue As String)
'    cboCombo.Tag = vNewValue
'End Property
Public Property Get Alinhamento() As AlinhamtoLabel
    On Error GoTo Trata

1   If bRegistrado Then Alinhamento = alinh
    Exit Property
Trata:
End Property

Public Property Let Alinhamento(ByVal vnewvalue As AlinhamtoLabel)
    On Error GoTo Trata

1   If bRegistrado Then
2       alinh = vnewvalue

3       Select Case alinh

            Case alinhEsquerdo
4               lblRotulo.Alignment = AlignmentConstants.vbRightJustify

5           Case alinhAcima
6               lblRotulo.Alignment = AlignmentConstants.vbLeftJustify
7       End Select

8       UserControl_Resize
9   End If

    Exit Property
Trata:
End Property

Public Property Get EnterEqvTab() As Boolean
    On Error GoTo Trata

1   If bRegistrado Then EnterEqvTab = sndkys
    Exit Property
Trata:
End Property

Public Property Let EnterEqvTab(ByVal vnewvalue As Boolean)
    On Error GoTo Trata

1   If bRegistrado Then sndkys = vnewvalue
    Exit Property
Trata:
End Property

Public Property Get corFundo() As OLE_COLOR
    On Error GoTo Trata

1   If bRegistrado Then corFundo = BackColor
    Exit Property
Trata:
End Property

Public Property Let corFundo(ByVal vnewvalue As OLE_COLOR)
    On Error GoTo Trata

1   If bRegistrado Then BackColor = vnewvalue
    Exit Property
Trata:
End Property

Public Property Get CorRotulo() As OLE_COLOR
    On Error GoTo Trata

1   If bRegistrado Then CorRotulo = lblRotulo.ForeColor
    Exit Property
Trata:
End Property

Public Property Let CorRotulo(ByVal vnewvalue As OLE_COLOR)
    On Error GoTo Trata

1   If bRegistrado Then lblRotulo.ForeColor = vnewvalue
    Exit Property
Trata:
End Property

Public Property Get corTexto() As OLE_COLOR
    On Error GoTo Trata

1   If bRegistrado Then corTexto = cboCombo.ForeColor
    Exit Property
Trata:
End Property

Public Property Let corTexto(ByVal vnewvalue As OLE_COLOR)
    On Error GoTo Trata

1   If bRegistrado Then cboCombo.ForeColor = vnewvalue
    Exit Property
Trata:
End Property

Public Property Get ValorPadrao() As String
    On Error GoTo Trata

1   If bRegistrado Then ValorPadrao = padr
    Exit Property
Trata:
End Property

Public Property Let ValorPadrao(ByVal vnewvalue As String)
    On Error GoTo Trata

1   If bRegistrado Then padr = vnewvalue
    Exit Property
Trata:
End Property

Public Sub Preencher(BDados As VSDados, Tabela As String, Optional ColunaExibicao As Integer = 0)
    On Error GoTo Trata

1   If bRegistrado Then
2       Dim RS As VSRecordset
3       Dim ColunaExtra As TipoColuna
4       Dim i As Integer, j As Integer
5       cboCombo.Clear: Set colDados = New Collection

6       If BDados.AbreTabela(Tabela, RS) Then

7           Do Until RS.EOF

8               If Not IsNull(RS(ColunaExibicao)) Then

9                   If Trim$(RS(ColunaExibicao)) <> "" Then
10                      cboCombo.AddItem RS(ColunaExibicao)
11                      j = RS.Fields.Count - 1

12                      For i = 0 To j
13                          ColunaExtra.Ordem = i
14                          ColunaExtra.Nome = UCase$(RS.Fields(i).Name)
15                          ColunaExtra.Valor = RS.Fields(i).Value
16                          ColunaExtra.ListIndex = cboCombo.NewIndex
17                          colDados.Add ColunaExtra, "IDX" & ColunaExtra.ListIndex & "COL" & i
18                      Next

19                  End If

20              End If

21              RS.MoveNext
22          Loop

23      End If

24      BDados.FechaTabela RS
25  End If

    Exit Sub
Trata:
End Sub

Public Property Get Coluna(Qual As Variant) As TipoColuna
    On Error GoTo Trata

1   If bRegistrado Then

2       If cboCombo.ListIndex >= 0 Then

3           If Not colDados Is Nothing Then

4               If IsNumeric(Qual) Then
5                   Coluna = colDados("IDX" & cboCombo.ListIndex & "COL" & Qual)

6               Else
7                   Dim col

8                   For Each col In colDados

9                       If col.ListIndex = cboCombo.ListIndex Then

10                          If col.Nome = UCase$(Qual) Then
11                              Coluna = colDados("IDX" & col.ListIndex & "COL" & col.Ordem)
12                              Exit For
13                          End If

14                      End If

15                  Next

16              End If

17          End If

18      End If

19  End If

    Exit Property
Trata:
End Property

Public Sub Exibir(QualColuna As Variant, Valor As Variant)
    On Error GoTo Trata

1   If bRegistrado Then

3       If Not colDados Is Nothing Then
4           Dim col

5           For Each col In colDados

6               If IsNumeric(QualColuna) Then

7                   If col.Ordem = QualColuna Then

8                       If col.Valor = Valor Then
9                           cboCombo.ListIndex = col.ListIndex
10                          Exit For
11                      End If

12                  End If

13              Else

14                  If col.Nome = QualColuna Then

15                      If col.Valor = Valor Then
16                          cboCombo.ListIndex = col.ListIndex
17                          Exit For
18                      End If 'valor=valor

19                  End If 'nome=qualcoluna

20              End If 'qualcoluna=numeric

21          Next

22      End If 'coldados = nothing

23  End If

24  Exit Sub
Trata:
End Sub

Public Property Get ListIndex() As Integer
    On Error GoTo Trata

1   If bRegistrado Then ListIndex = cboCombo.ListIndex
    Exit Property
Trata:
End Property

Public Property Let ListIndex(ByVal vnewvalue As Integer)
    On Error GoTo Trata

1   If bRegistrado Then cboCombo.ListIndex = vnewvalue
    Exit Property
Trata:
End Property

Public Sub AddItem(Item As String, Optional Index)
    On Error GoTo Trata

1   If bRegistrado Then cboCombo.AddItem Item, Index
    Exit Sub
Trata:
End Sub

Public Sub Clear()
    On Error GoTo Trata

1   If bRegistrado Then cboCombo.Clear
    Exit Sub
Trata:
End Sub

Public Property Get Editavel() As Boolean
    On Error GoTo Trata

1   If bRegistrado Then Editavel = editav
    Exit Property
Trata:
End Property

Public Property Let Editavel(ByVal vnewvalue As Boolean)
    On Error GoTo Trata

1   If bRegistrado Then editav = vnewvalue
    Exit Property
Trata:
End Property

Public Property Get Enabled() As Boolean
    On Error GoTo Trata

1   If bRegistrado Then Enabled = cboCombo.Enabled
    Exit Property
Trata:
End Property

Public Property Let Enabled(ByVal vnewvalue As Boolean)
    On Error GoTo Trata

1   If bRegistrado Then
2       cboCombo.Enabled = vnewvalue
3       lblRotulo.Enabled = vnewvalue
4       cboCombo.TabStop = vnewvalue
5   End If

    Exit Property
Trata:
End Property

Public Sub PreencherGeral(BDados As VSDados, Tabela As String)
    On Error GoTo Trata

1   If bRegistrado Then
2       Dim RS As VSRecordset
3       Dim ColunaExtra As TipoColuna
4       Dim Sql As String
5       cboCombo.Clear: Set colDados = New Collection
6       Sql = "SELECT TGE_NOME, TGE_CODIGO, TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO>0 AND TGE_TIPO=(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='" & Tabela & "') ORDER BY TGE_NOME"

7       If BDados.AbreTabela(Sql, RS) Then

8           Do While Not RS.EOF
9               cboCombo.AddItem RS!TGE_NOME
10              ColunaExtra.Ordem = 0
11              ColunaExtra.Nome = "TGE_NOME"
12              ColunaExtra.Valor = RS!TGE_NOME
13              ColunaExtra.ListIndex = cboCombo.NewIndex
14              colDados.Add ColunaExtra, "IDX" & ColunaExtra.ListIndex & "COL" & 0
15              ColunaExtra.Ordem = 1
16              ColunaExtra.Nome = "TGE_CODIGO"
17              ColunaExtra.Valor = RS!TGE_CODIGO
18              ColunaExtra.ListIndex = cboCombo.NewIndex
19              colDados.Add ColunaExtra, "IDX" & ColunaExtra.ListIndex & "COL" & 1
20              ColunaExtra.Ordem = 2
21              ColunaExtra.Nome = "TGE_TIPO"
22              ColunaExtra.Valor = RS!TGE_TIPO
23              ColunaExtra.ListIndex = cboCombo.NewIndex
24              colDados.Add ColunaExtra, "IDX" & ColunaExtra.ListIndex & "COL" & 2
25              RS.MoveNext
26          Loop

27      End If

28      BDados.FechaTabela RS
29  End If

    Exit Sub
Trata:
End Sub

Private Sub DrawCombo(ByVal dwStyle As DrawCombo)
    On Error GoTo Trata

1   If bRegistrado Then
2       Dim rct As RECT
3       Dim cmbDC As Long
        ' get the combobox area
4       GetClientRect cboCombo.hWnd, rct
5       cmbDC = GetDC(cboCombo.hWnd)
        ' draw a rectangle over the combobox

6       Select Case dwStyle

            Case FC_DRAWDISABLED
7               DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
8               InflateRect rct, -1, -1
9               DrawRect cmbDC, rct, vb3DHighlight, vb3DHighlight

10          Case FC_DRAWNORMAL
11              DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
12              InflateRect rct, -1, -1
13              DrawRect cmbDC, rct, vbButtonFace, vbButtonFace

14          Case Else
15              DrawRect cmbDC, rct, vbButtonShadow, vb3DHighlight
16              InflateRect rct, -1, -1
17              DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
18      End Select

19      InflateRect rct, -1, -1
20      rct.Left = rct.Right - GetSystemMetrics(SM_CXHTHUMB)
21      DrawRect cmbDC, rct, vbButtonFace, vbButtonFace
22      InflateRect rct, -1, -1
23      DrawRect cmbDC, rct, vbButtonFace, vbButtonFace

24      Select Case dwStyle

            Case FC_DRAWNORMAL
25              rct.Top = rct.Top - 1
26              rct.Bottom = rct.Bottom + 1
27              DrawRect cmbDC, rct, vb3DHighlight, vb3DHighlight
28              rct.Left = rct.Left - 1
29              rct.Right = rct.Left
30              DrawRect cmbDC, rct, vbWindowBackground, &H0

31          Case FC_DRAWRAISED
32              rct.Top = rct.Top - 1
33              rct.Bottom = rct.Bottom + 1
34              rct.Right = rct.Right + 1
35              DrawRect cmbDC, rct, vb3DHighlight, vbButtonShadow

36          Case FC_DRAWPRESSED
37              rct.Left = rct.Left - 1
38              rct.Top = rct.Top - 2
39              OffsetRect rct, 1, 1
40              DrawRect cmbDC, rct, vbButtonShadow, vb3DHighlight
41      End Select

        ' release the memory
42      DeleteDC cmbDC
43  End If
    Exit Sub
Trata:
End Sub

' Draw the border for the new combobox
Private Function DrawRect(ByVal hdc As Long, ByRef rct As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR)
    On Error GoTo Trata

1   If bRegistrado Then
2       Dim hPen As Long
3       Dim hPenOld As Long
4       Dim tP As POINTAPI
        ' create and select a pen
5       hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
6       hPenOld = SelectObject(hdc, hPen)
        ' draw the lines
7       MoveToEx hdc, rct.Left, rct.Bottom - 1, tP
8       LineTo hdc, rct.Left, rct.Top
9       LineTo hdc, rct.Right - 1, rct.Top
        ' select the original pen and delete the pen created above
10      SelectObject hdc, hPenOld
11      DeleteObject hPen
        ' do the same for the Bottom-right border

12      If (rct.Left <> rct.Right) Then
13          hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
14          hPenOld = SelectObject(hdc, hPen)
15          LineTo hdc, rct.Right - 1, rct.Bottom - 1
16          LineTo hdc, rct.Left, rct.Bottom - 1
17          SelectObject hdc, hPenOld
18          DeleteObject hPen
19      End If

20  End If

    Exit Function
Trata:
End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    On Error GoTo Trata
1   If bRegistrado Then

2       If OleTranslateColor(clr, hPal, TranslateColor) Then
3           TranslateColor = -1
4       End If

5   End If
    Exit Function
Trata:
End Function

Public Sub SetarLinha(Codigo, Optional NumColuna = 0)
    On Error GoTo Trata
1   If bRegistrado Then
2       Dim i As Integer, col As Integer
3       col = IIf(IsMissing(NumColuna), 0, NumColuna)
        bTravaEventoClick = True
4       For i = 0 To cboCombo.ListCount - 1
5           cboCombo.ListIndex = i
            If Not colDados Is Nothing Then
                If Trim$(colDados("IDX" & i & "COL" & col).Valor) = Trim$(Codigo) Then Exit For
            End If
7       Next
        bTravaEventoClick = False
8       If i > cboCombo.ListCount - 1 Then
9           cboCombo.ListIndex = -1
10      End If

11  End If
    Exit Sub
Trata:
End Sub

Public Function ItemData(Index As Integer) As Long
    On Error GoTo Trata
1   If bRegistrado Then ItemData = cboCombo.ItemData(Index)
    Exit Function
Trata:
End Function

Public Function List(Index As Integer) As String
    On Error GoTo Trata
1   If bRegistrado Then List = cboCombo.List(Index)
    Exit Function
Trata:
End Function

Public Sub RemoveItem(Index As Integer)
    On Error GoTo Trata
1   If bRegistrado Then cboCombo.RemoveItem (Index)
    Exit Sub
Trata:
End Sub


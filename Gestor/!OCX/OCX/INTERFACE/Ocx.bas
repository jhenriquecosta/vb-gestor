Attribute VB_Name = "Ocx"
Option Explicit
Public Util As VSUtil
Public bRegistrado As Boolean 'Indica se o controle está registrado

Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As _
    Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal _
    hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' this variable is shared between the two routines below
Dim m_ClientIsInterpreted As Boolean

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type LV_FINDINFO
    flags As Long
    psz As String
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type
Public objFind As LV_FINDINFO

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
Public objItem As LV_ITEM

Public Const LVFI_PARAM As Long = &H1
Public Const LVIF_TEXT As Long = &H1

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Public Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

  
Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

 
   

'*********
'txtVISUAL
Public Const cteAmarelo As Long = &HC0FFFF
Public Const cteBranco As Long = &HFFFFFF
'******************************************

'********
'grdVISUAL
'Public Colunas As clsColunas
Public sOrder As Boolean
Public lngSubItem As Long
Public OpcoesImpressao As String
'******************************************

Public Function CompareText(ByVal lParam1 As Long, _
   ByVal lParam2 As Long, _
   ByVal hWnd As Long) As Long
' VTOcx.Ocx.Function CompareText
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : lParam1 (Long)
'              lParam2 (Long)
'              hwnd (Long)
'
' Ex:
'--------------------------------------------------------------------------------
     
    'CompareStrings: This is the sorting routine that gets passed to the
    'ListView control to provide the comparison test for String values.

    'Compare returns:
    ' 0 = Less Than
    ' 1 = Equal
    ' 2 = Greater Than

    Dim dString1 As String
    Dim dString2 As String
     
    'Obtain the item names and Strings corresponding to the
    'input parameters
    dString1 = ListView_GetItemString(hWnd, lParam1)
    dString2 = ListView_GetItemString(hWnd, lParam2)
     
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the Strings appropriately:

    Select Case sOrder

        Case True: 'sort descending
            
            If dString1 < dString2 Then

                CompareText = 0

            ElseIf dString1 = dString2 Then

                CompareText = 1
                Else: CompareText = 2

            End If
      
        Case Else: 'sort ascending
   
            If dString1 > dString2 Then

                CompareText = 0

            ElseIf dString1 = dString2 Then

                CompareText = 1
                Else: CompareText = 2

            End If
   
    End Select

End Function

Public Function CompareDates(ByVal lParam1 As Long, _
   ByVal lParam2 As Long, _
   ByVal hWnd As Long) As Long
' VTOcx.Ocx.Function CompareDates
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : lParam1 (Long)
'              lParam2 (Long)
'              hwnd (Long)
'
' Ex:
'--------------------------------------------------------------------------------
     
    'CompareDates: This is the sorting routine that gets passed to the
    'ListView control to provide the comparison test for date values.

    'Compare returns:
    ' 0 = Less Than
    ' 1 = Equal
    ' 2 = Greater Than

    Dim dDate1 As Date
    Dim dDate2 As Date
     
    'Obtain the item names and dates corresponding to the
    'input parameters
    dDate1 = ListView_GetItemDate(hWnd, lParam1)
    dDate2 = ListView_GetItemDate(hWnd, lParam2)
     
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the dates appropriately:

    Select Case sOrder

        Case True: 'sort descending
            
            If dDate1 < dDate2 Then

                CompareDates = 0

            ElseIf dDate1 = dDate2 Then

                CompareDates = 1
                Else: CompareDates = 2

            End If
      
        Case Else: 'sort ascending
   
            If dDate1 > dDate2 Then

                CompareDates = 0

            ElseIf dDate1 = dDate2 Then

                CompareDates = 1
                Else: CompareDates = 2

            End If
   
    End Select

End Function

Public Function CompareValues(ByVal lParam1 As Long, _
   ByVal lParam2 As Long, _
   ByVal hWnd As Long) As Long
' VTOcx.Ocx.Function CompareValues
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : lParam1 (Long)
'              lParam2 (Long)
'              hwnd (Long)
'
' Ex:
'--------------------------------------------------------------------------------
     
    'CompareValues: This is the sorting routine that gets passed to the
    'ListView control to provide the comparison test for numeric values.

    'Compare returns:
    ' 0 = Less Than
    ' 1 = Equal
    ' 2 = Greater Than
  
    Dim val1 As Double
    Dim val2 As Double
     
    'Obtain the item names and values corresponding
    'to the input parameters
    val1 = ListView_GetItemValueStr(hWnd, lParam1)
    val2 = ListView_GetItemValueStr(hWnd, lParam2)
     
    'based on the Public variable sOrder set in the
    'columnheader click sub, sort the values appropriately:

    Select Case sOrder

        Case True: 'sort descending
            
            If val1 < val2 Then

                CompareValues = 0

            ElseIf val1 = val2 Then

                CompareValues = 1
                Else: CompareValues = 2

            End If
      
        Case Else: 'sort ascending
   
            If val1 > val2 Then

                CompareValues = 0

            ElseIf val1 = val2 Then

                CompareValues = 1
                Else: CompareValues = 2

            End If
   
    End Select

End Function

Public Function ListView_GetItemDate(hWnd As Long, lParam As Long) As Date
' VTOcx.Ocx.Function ListView_GetItemDate
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : hwnd (Long)
'              lParam (Long)
'
' Ex:
'--------------------------------------------------------------------------------
  
    Dim hIndex As Long
    Dim r As Long
  
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = lngSubItem
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 1
    'and convert it into a date and exit
    r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)

    If r > 0 Then

        ListView_GetItemDate = CDate(Left$(objItem.pszText, r))

    End If
  
End Function

Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long) As Double
' VTOcx.Ocx.Function ListView_GetItemValueStr
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : hwnd (Long)
'              lParam (Long)
'
' Ex:
'--------------------------------------------------------------------------------

    Dim hIndex As Long
    Dim r As Long
  
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = lngSubItem
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 2
    'and convert it into a long
    r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
    Dim Util As New VSClass.VSUtil
    If r > 0 Then

        ListView_GetItemValueStr = CDbl(Util.Nvl(Trim$(Left$(objItem.pszText, r)), 0))

    End If

End Function

Public Function ListView_GetItemString(hWnd As Long, lParam As Long) As String
' VTOcx.Ocx.Function ListView_GetItemString
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : hwnd (Long)
'              lParam (Long)
'
' Ex:
'--------------------------------------------------------------------------------
  
    Dim hIndex As Long
    Dim r As Long
  
    'Convert the input parameter to an index in the list view
    objFind.flags = LVFI_PARAM
    objFind.lParam = lParam
    hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
     
    'Obtain the value of the specified list view item.
    'The objItem.iSubItem member is set to the index
    'of the column that is being retrieved.
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = lngSubItem
    objItem.pszText = Space$(32)
    objItem.cchTextMax = Len(objItem.pszText)
     
    'get the string at subitem 1
    'and convert it into a String and exit
    r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)

    If r > 0 Then

        ListView_GetItemString = CStr(Left$(objItem.pszText, r))

    End If
  
End Function

Public Function FARPROC(ByVal pfn As Long) As Long
' VTOcx.Ocx.Function FARPROC
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:22
'
' Descricao  : Descreva
'
' Parametros : pfn (Long)
'
' Ex:
'--------------------------------------------------------------------------------

    FARPROC = pfn

End Function

' return True if the client application of this DLL
' is an interpreted Visual Basic program running in the IDE
'
' NOTE: this code is meant to be inserted in a BAS module
'       inside an ActiveX DLL project

Public Function Programando() As Boolean
    EnumThreadWindows App.ThreadID, AddressOf EnumThreadWindows_CBK, 0
    Programando = m_ClientIsInterpreted
End Function

' this is a callback function that is executed for each
' window in the same thead as the DLL

Private Function EnumThreadWindows_CBK(ByVal hWnd As Long, _
    ByVal lParam As Long) As Boolean
    Dim buffer As String * 512
    Dim length As Long
    Dim windowClass As String

    ' get the class name of this window
    length = GetClassName(hWnd, buffer, Len(buffer))
    windowClass = Left$(buffer, length)
    
    If windowClass = "IDEOwner" Then
        ' this is the main VB IDE window, therefore
        ' the client application is interpreted
        m_ClientIsInterpreted = True
        ' return False to stop evaluation
        EnumThreadWindows_CBK = False
    Else
        ' return True to continue enumeration
        EnumThreadWindows_CBK = True
    End If
End Function

Public Sub ValidaComponente(Componente As String)
    
    bRegistrado = True
    Exit Sub
    
    On Error GoTo Trata
    
    'Faz validacao de registro se o componente estiver sendo
    'utilizado para desenvolvimento
    Dim Reg As Object
    Set Reg = CreateObject("VTRegistro.VTReg")
    If Programando Then
        bRegistrado = Reg.Verifica(Componente)
    Else
'        bRegistrado = True
        bRegistrado = Reg.Verifica_Cli("COMPONENTES")
    End If
    Exit Sub
Trata:
    bRegistrado = False
End Sub



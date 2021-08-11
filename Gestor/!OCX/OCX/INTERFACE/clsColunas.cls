VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' The private collection object
Private m_clsColunas As Collection

Private Sub Class_Initialize()
' VTOcx.clsColunas.Sub Class_Initialize
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    Set m_clsColunas = New Collection

End Sub

' This sub adds a new clsColuna item to the collection.
Sub Add(Item As clsColuna, Optional Key As Variant)
' VTOcx.clsColunas.Sub Add
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros : Item (clsColuna)
'              Key (Variant)
'
' Ex:
'--------------------------------------------------------------------------------

    'TODO: Initialize the new clsColuna item's properties here

    m_clsColunas.Add Item, Key

End Sub

' This sub remove an item from the collection.
Sub Remove(Index As Variant)
' VTOcx.clsColunas.Sub Remove
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros : Index (Variant)
'
' Ex:
'--------------------------------------------------------------------------------

    m_clsColunas.Remove Index

End Sub

' This function returns a clsColuna item from the collection. It's the default method.
Function Item(Index As Variant) As clsColuna
Attribute Item.VB_UserMemId = 0
' VTOcx.clsColunas.Function Item
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros : Index (Variant)
'
' Ex:
'--------------------------------------------------------------------------------

    Set Item = Nothing

    On Error Resume Next
    Set Item = m_clsColunas.Item(Index)
    On Error GoTo 0

End Function

' This function returns a boolean value indicating whether the item exists in the collection or not.
Function Exists(Index As Variant) As Boolean
' VTOcx.clsColunas.Function Exists
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros : Index (Variant)
'
' Ex:
'--------------------------------------------------------------------------------

    Dim objTest As clsColuna
    Set objTest = Nothing

    If Count > 0 Then

        On Error Resume Next
        Set objTest = Item(Index)
        On Error GoTo 0

    End If

    Exists = Not (objTest Is Nothing)

End Function

' This property returns the number of items in the collection.
Property Get Count() As Long
' VTOcx.clsColunas.Property Count
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    Count = m_clsColunas.Count

End Property

' This sub remove all items from the collection.
Sub Clear()
' VTOcx.clsColunas.Sub Clear
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    Set m_clsColunas = New Collection

End Sub

' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' VTOcx.clsColunas.Function NewEnum
'================================================================================
' Queiroz em VTDES_01
' 01/06/2002-14:03:26
'
' Descricao  : Descreva
'
' Parametros :
'
' Ex:
'--------------------------------------------------------------------------------

    Set NewEnum = m_clsColunas.[_NewEnum]

End Function


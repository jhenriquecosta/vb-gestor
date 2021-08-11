Attribute VB_Name = "Module2"
''''''''''''''''''''''''''''''''
' Visual Basic global constant '

' Show parameters
Global Const MODAL = 1
Global Const MODELESS = 0

' Key Codes
Global Const KEY_LBUTTON = &H1
Global Const KEY_RBUTTON = &H2
Global Const KEY_CANCEL = &H3
Global Const KEY_MBUTTON = &H4    ' NOT contiguous with L & RBUTTON
Global Const KEY_BACK = &H8
Global Const KEY_TAB = &H9
Global Const KEY_CLEAR = &HC
Global Const KEY_RETURN = &HD
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_MENU = &H12
Global Const KEY_PAUSE = &H13
Global Const KEY_CAPITAL = &H14
Global Const KEY_ESCAPE = &H1B
Global Const KEY_SPACE = &H20
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_END = &H23
Global Const KEY_HOME = &H24
Global Const KEY_LEFT = &H25
Global Const KEY_UP = &H26
Global Const KEY_RIGHT = &H27
Global Const KEY_DOWN = &H28
Global Const KEY_SELECT = &H29
Global Const KEY_PRINT = &H2A
Global Const KEY_EXECUTE = &H2B
Global Const KEY_SNAPSHOT = &H2C
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_HELP = &H2F

' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'

Global Const KEY_NUMPAD0 = &H60
Global Const KEY_NUMPAD1 = &H61
Global Const KEY_NUMPAD2 = &H62
Global Const KEY_NUMPAD3 = &H63
Global Const KEY_NUMPAD4 = &H64
Global Const KEY_NUMPAD5 = &H65
Global Const KEY_NUMPAD6 = &H66
Global Const KEY_NUMPAD7 = &H67
Global Const KEY_NUMPAD8 = &H68
Global Const KEY_NUMPAD9 = &H69
Global Const KEY_MULTIPLY = &H6A
Global Const KEY_ADD = &H6B
Global Const KEY_SEPARATOR = &H6C
Global Const KEY_SUBTRACT = &H6D
Global Const KEY_DECIMAL = &H6E
Global Const KEY_DIVIDE = &H6F
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_F13 = &H7C
Global Const KEY_F14 = &H7D
Global Const KEY_F15 = &H7E
Global Const KEY_F16 = &H7F

Global Const KEY_NUMLOCK = &H90

' Variant VarType tags
Global Const V_EMPTY = 0
Global Const V_NULL = 1
Global Const V_INTEGER = 2
Global Const V_LONG = 3
Global Const V_SINGLE = 4
Global Const V_DOUBLE = 5
Global Const V_CURRENCY = 6
Global Const V_DATE = 7
Global Const V_STRING = 8

' Event Parameters
' QueryUnload
Global Const FORM_CONTROLMENU = 0
Global Const FORM_CODE = 1
Global Const APP_WINDOWS = 2
Global Const APP_TASKMANAGER = 3
Global Const FORM_MDIFORM = 4

' Properties
' Colors
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF

' BorderStyle (form)
Global Const NONE = 0          ' 0 - None
Global Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Global Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Global Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

' MousePointer
Global Const DEFAULT = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

' WindowState
Global Const NORMAL = 0    ' 0 - Normal
Global Const MINIMIZED = 1 ' 1 - Minimized
Global Const MAXIMIZED = 2 ' 2 - Maximized

' Check Value
Global Const UNCHECKED = 0 ' 0 - Unchecked
Global Const CHECKED = 1   ' 1 - Checked
Global Const GRAYED = 2    ' 2 - Grayed

' Shift parameter masks
Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4
' Button parameter masks
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4

' Function Parameters
' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal
' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed
' SetAttr, Dir, GetAttr functions
Global Const ATTR_NORMAL = 0
Global Const ATTR_READONLY = 1
Global Const ATTR_HIDDEN = 2
Global Const ATTR_SYSTEM = 4
Global Const ATTR_VOLUME = 8
Global Const ATTR_DIRECTORY = 16
Global Const ATTR_ARCHIVE = 32
'Data control
'Editmode property values
Global Const DATA_EDITNONE = 0
Global Const DATA_EDITMODE = 1
Global Const DATA_EDITADD = 2
'Common Dialog Control
'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6
'
' Data Access constants
'
' Option argument values (CreateDynaset, etc)
Global Const DB_DENYWRITE = &H1
Global Const DB_DENYREAD = &H2
Global Const DB_READONLY = &H4
Global Const DB_APPENDONLY = &H8
Global Const DB_INCONSISTENT = &H10
Global Const DB_CONSISTENT = &H20
Global Const DB_SQLPASSTHROUGH = &H40
' Field Data Types
Global Const DB_BOOLEAN = 1
Global Const DB_BYTE = 2
Global Const DB_INTEGER = 3
Global Const DB_LONG = 4
Global Const DB_CURRENCY = 5
Global Const DB_SINGLE = 6
Global Const DB_DOUBLE = 7
Global Const DB_DATE = 8
Global Const DB_TEXT = 10
Global Const DB_LONGBINARY = 11
Global Const DB_MEMO = 12
' CreateDatabase and CompactDatabase Language constants
Global Const DB_LANG_GENERAL = ";LANGID=0x0809;CP=1252;COUNTRY=0"
' CreateDatabase and CompactDatabase options
Global Const DB_VERSION10 = 1        ' Microsoft Access Version 1.0
Global Const DB_ENCRYPT = 2          ' Make database encrypted.
Global Const DB_DECRYPT = 4          ' Decrypt database while compacting.


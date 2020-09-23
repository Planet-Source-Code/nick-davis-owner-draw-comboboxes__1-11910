Attribute VB_Name = "API"
Option Explicit

' ==================================================================
' Filename:     API.bas
' Description:  API function declarations and associated constants
' ------------------------------------------------------------------
' Created by:   Nicholas Davis      Date: 06-Sep-00
' Updated by:
' ------------------------------------------------------------------
' Notes:
'
'
'===================================================================


' **************************************************
'           API FUNCTION DECLARATIONS
' **************************************************

' -------------------------------------
' Sub Classing related functions
' -------------------------------------

' Adds a new entry or changes an existing entry in the property list of the specified window.
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
' Retrieves a data handle from the property list of the given window.
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
' Removes an entry from the property list of the specified window.
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

' --------------------------------------------
' Window manipulation related functions
' --------------------------------------------
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Boolean
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean

' -----------------------
' GDI Functions
' -----------------------
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hGDIObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hGDIObject As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal COLORREF As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Integer, ByVal COLORREF As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenstyle As Integer, ByVal nWidth As Integer, ByVal COLORREF As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Integer, ByVal Y As Integer, ByVal lpPoint As Long) As Boolean
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal nXEnd As Integer, ByVal nYEnd As Integer) As Boolean

Public Declare Function TextOutBStr Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpString As Any, ByVal nCount As Long) As Long

' Draw a highlighting rectangle
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Integer
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColour As Long) As Long

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Integer, ByVal nWidth As Integer, ByVal nEscapement As Integer, ByVal nOrientation As Integer, ByVal fnWeight As Integer, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPItchAndFamily As Long, ByVal lpszFace As Long) As Long

' ---------------------------
' General Functions
' ---------------------------
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

' ***************************************************
'               CONSTANT DECLARATIONS
' ***************************************************

' -----------------------------
' Window constants
' -----------------------------

' WindowProc Style flags
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

' Window Style Constants
Public Const WS_OVERLAPPED = &H0
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const WS_VISIBLE = &H10000000

Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4

' ------------------------------
' Windows Message constants
' ------------------------------
Public Const WM_DRAWITEM = &H2B
Public Const WM_COMMAND = &H111
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCHITTEST = &H84
Public Const WM_NCMOUSEMOVE = &HA0

' Hit test values
Public Const HTCAPTION = 2
Public Const HTBORDER = 18

' -------------------
' GDI Constants
' -------------------

' Windows Brush styles
Public Const HS_HORIZONTAL = 0
Public Const HS_VERTICAL = 1
Public Const HS_FDIAGONAL = 2
Public Const HS_BDIAGONAL = 3
Public Const HS_CROSS = 4
Public Const HS_DIAGCROSS = 5
' Not actually Windows Brush constants but we need to display
' a solid brush and a no fill
' in the combo box (we'll process this manually)
Public Const HS_NOFILL = 100
Public Const HS_SOLID = 101

' Windows Pen Styles
Public Const PS_SOLID = 0
Public Const PS_DASH = 1
Public Const PS_DOT = 2
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4

' -----------------------------
' Combo Box constants
' -----------------------------

' Combo styles
Public Const CBS_DROPDOWNLIST = &H3
Public Const CBS_OWNERDRAWFIXED = &H10
Public Const CBS_HASSTRINGS = &H200
Public Const CBS_AUTOHSCROLL = &H40

' Combo box message constants
Public Const CB_ADDSTRING = &H143
Public Const CB_GETITEMDATA = &H150
Public Const CB_SETITEMDATA = &H151
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160

' Notification (Sent with WM_COMMAND)
Public Const CBN_SELCHANGE = 1

' -----------------------------------------------
' WM_DRAWITEM message constants (sent in the
' DRAWITEMSTRUCT structure
' -----------------------------------------------

' Item action constants
Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_FOCUS = &H2
Public Const ODA_SELECT = &H4

' Item state constants
Public Const ODS_FOCUS = &H10
Public Const ODS_COMBOBOXEDIT = &H1000
' User defined constants
Public Const ODS_FOCUSONEDITBOX = &H1011
Public Const ODS_FOCUSITEM = &H11

' ------------------------
' Font constants
' ------------------------
Public Const ANSI_CHARSET = 0
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 800

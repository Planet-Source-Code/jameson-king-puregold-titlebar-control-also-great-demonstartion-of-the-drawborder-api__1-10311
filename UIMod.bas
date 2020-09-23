Attribute VB_Name = "UIMod"
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function CreateWindowEx Lib "user32" (DWORD As dwExStyle, LPCTSTR As lPClassName, LPCTSTR As lpWindowName, _
'DWORD As dwStyle, Inte As x, Inte As y, Inte As nWidth, Inte As nHeight, Hwnd As hWndParent, hMenu As hMenu, _
'HINSTACE As hInstance, LPVOID As lpParam)
'Define CreateWindowA(lPClassName, lpWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam)
'Public WithEvents menuPG As cPopupMenu
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const DC_ACTIVE = &H1
Public Const DC_SMALLCAP = &H2
Public Const DC_ICON = &H4
Public Const DC_TEXT = &H8
Public Const DC_INBUTTON = &H10
Public Const DC_GRADIENT = &H20
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function RegainCapture Lib "user32" () As Long
Declare Function Capture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2




Public Const GWL_EXSTYLE = (-20)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const WS_EX_LAYERED = &H80000




Public Type OLECOLOR
    RedOrSys As Byte
    Green As Byte
    Blue As Byte
    Type As Byte
End Type

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const DT_CENTER = &H1
Public Const DT_WORDBREAK = &H10
Public Const DT_CENTERCENTER = &H65
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long



Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_RAISED2 = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10


Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP _
        Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP _
        Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
        Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM _
        Or BF_RIGHT)
Public Enum EDEDBorderParts
    BF_MIDDLE = &H800
    BF_SOFT = &H1000
    BF_ADJUST = &H2000
    BF_FLAT = &H4000
    BF_MONO = &H8000
    BF_ALL = BF_RECT Or BF_MIDDLE Or BF_SOFT Or BF_ADJUST Or BF_FLAT Or BF_MONO
End Enum
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean



Public Sub MakeWindowLayered(hwnd As Long)
    Dim ExStyles As Long
    ExStyles = GetWindowLong(hwnd, GWL_EXSTYLE)
    ExStyles = ExStyles Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, ExStyles
End Sub

Public Sub SetColorKey(hwnd As Long, ByVal Color As Long)
    Color = WinColor(Color)
    SetLayeredWindowAttributes hwnd, Color, 0, LWA_COLORKEY
End Sub

Public Function WinColor(VBColor As Long) As Long
    Dim SysClr As OLECOLOR
    CopyMemory SysClr, VBColor, Len(SysClr)
    If SysClr.Type = &H80 Then
        WinColor = GetSysColor(SysClr.RedOrSys)
    Else
        WinColor = VBColor
    End If
End Function
Public Sub SetOpacity(hwnd As Long, Opacity As Byte)
    SetLayeredWindowAttributes hwnd, 0, Opacity, LWA_ALPHA
End Sub


Sub CapPaint(Control, current As Form)
Dim rt As RECT
   Dim rtn As Long
With rt
      .Left = 0
      .Top = 0
      .Right = Control.ScaleX(Control.ScaleWidth, Control.ScaleMode, vbPixels)
      .Bottom = 25
   End With
rtn = DrawCaption(current.hwnd, Control.hdc, rt, DC_ACTIVE Or DC_ICON Or DC_TEXT Or DC_GRADIENT)
Control.Refresh
End Sub
Sub CapPaintINA(Control, current As Form)
Dim rt As RECT
Dim rtn As Long
With rt
      .Left = 0
      .Top = 0
      .Right = Control.ScaleX(Control.ScaleWidth, Control.ScaleMode, vbPixels)
      .Bottom = 25
End With
rtn = DrawCaption(current.hwnd, Control.hdc, rt, DC_ICON Or DC_TEXT Or DC_GRADIENT)
Control.Refresh
End Sub
Public Sub MoveForm(Frm)
ReleaseCapture
x = SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
Public Sub MoveFormC(Frm)
ReleaseCapture
x = SendMessage(Frm, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
Function GetHDC(ByVal hwnd As Long)
GetHDC = GetDC(hwnd)
End Function




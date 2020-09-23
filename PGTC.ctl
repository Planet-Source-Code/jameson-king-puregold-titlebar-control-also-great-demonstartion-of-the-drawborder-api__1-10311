VERSION 5.00
Begin VB.UserControl PGTC 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   3135
   ToolboxBitmap   =   "PGTC.ctx":0000
End
Attribute VB_Name = "PGTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Default Property Values:
Const m_def_CaptionControl = 0
Const m_def_FocusTitle = False
'Const m_def_CaptionControl = False
Const m_def_AlphaValue = 128
Const m_def_Alpha = False
'Property Variables:
Dim m_CaptionControl As Boolean
Dim m_FocusTitle As Boolean
'Dim m_CaptionControl As Boolean
Dim m_AlphaValue As Byte
Dim m_Alpha As Boolean
Dim mH As Boolean
'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseEnter()
Event MouseExit()
'Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MoveFormC GetParentHwnd(UserControl.hwnd, "Form")
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
If m_FocusTitle <> True Then GoTo 5
Dim rt As RECT
Dim X1 As Single
Dim Y1 As Single
GetWindowRect UserControl.hwnd, rt
X1 = rt.Left
Y1 = rt.Top

If MoveM = False And mH = True Then
GoTo 5
Else:
End If

If MoveM <> True Then
        SetCapture (UserControl.hwnd)
        RaiseEvent MouseEnter
        mH = True
        UserControl_Paint
    Else:
        RaiseEvent MouseExit
        mH = False
        UserControl_Paint
        ReleaseCapture
End If
5
End Sub
Function MoveM()
' This is pretty self explanatory
Dim Pt As POINTAPI
Dim rt As RECT
GetCursorPos Pt
GetWindowRect UserControl.hwnd, rt
 If Pt.x >= rt.Left And Pt.x <= rt.Right And Pt.y >= rt.Top And Pt.y <= rt.Bottom Then
 SetCapture (UserControl.hwnd)
 MoveM = False
 Else:
 ReleaseCapture
 MoveM = True
 End If
End Function

Private Sub UserControl_Paint()
Dim rt As RECT
Dim rtn As Long
If Ambient.UserMode = False Then
rt.Left = 0
rt.Top = 0
rt.Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
rt.Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
DrawFocusRect UserControl.hdc, rt
Else:
        With rt
        .Left = 0
        .Top = 0
        .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
        .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
        End With
        rtn = DrawCaption(GetParentHwnd(UserControl.hwnd, "Form"), UserControl.hdc, rt, DC_ACTIVE Or DC_ICON _
        Or DC_TEXT Or DC_GRADIENT)
End If



If m_Alpha = "True" Then
MakeWindowLayered GetParentHwnd(UserControl.hwnd, "Form")
SetOpacity GetParentHwnd(UserControl.hwnd, "Form"), m_AlphaValue
Else:
End If


If mH = True Then
rt.Left = 0
rt.Top = 0
rt.Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
rt.Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
DrawEdge UserControl.hdc, rt, &H4, BF_RECT
Else:
End If
End Sub

Function GetClass(child)
Dim Buffer$
Dim getclas%
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Function GetParentHwnd(hwnd, str)
par = GetParent(hwnd)
DoEvents
Do
Cla1 = GetClass(par)
If InStr(1, Cla1, str, vbTextCompare) <> 0 Then
Result = Cla1
Debug.Print "Class Search Result:     " & par & "     " & "Class Name     " & Cla1
Else:
par = GetParent(par)
Debug.Print "Class Search Result:     " & par & "     " & "Class Name     " & Cla1
End If
Loop Until Result <> ""
GetParentHwnd = par
End Function


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    m_Alpha = PropBag.ReadProperty("Alpha", m_def_Alpha)
    m_AlphaValue = PropBag.ReadProperty("AlphaValue", m_def_AlphaValue)
'    m_CaptionControl = PropBag.ReadProperty("CaptionControl", m_def_CaptionControl)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    m_FocusTitle = PropBag.ReadProperty("FocusTitle", m_def_FocusTitle)
    m_CaptionControl = PropBag.ReadProperty("CaptionControl", m_def_CaptionControl)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("Alpha", m_Alpha, m_def_Alpha)
    Call PropBag.WriteProperty("AlphaValue", m_AlphaValue, m_def_AlphaValue)
'    Call PropBag.WriteProperty("CaptionControl", m_CaptionControl, m_def_CaptionControl)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, True)
    Call PropBag.WriteProperty("FocusTitle", m_FocusTitle, m_def_FocusTitle)
    Call PropBag.WriteProperty("CaptionControl", m_CaptionControl, m_def_CaptionControl)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Alpha() As Boolean
Attribute Alpha.VB_Description = "Sets the Caption or the Parent Form to Alpha Fade"
    Alpha = m_Alpha
End Property

Public Property Let Alpha(ByVal New_Alpha As Boolean)
    m_Alpha = New_Alpha
    PropertyChanged "Alpha"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Alpha = m_def_Alpha
    m_AlphaValue = m_def_AlphaValue
'    m_CaptionControl = m_def_CaptionControl
    UserControl.AutoRedraw = False
    m_FocusTitle = m_def_FocusTitle
    m_CaptionControl = m_def_CaptionControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get AlphaValue() As Long
Attribute AlphaValue.VB_Description = "Sets or Gives the value for the Alpha Fade"
    AlphaValue = m_AlphaValue
End Property

Public Property Let AlphaValue(ByVal New_AlphaValue As Long)
    m_AlphaValue = New_AlphaValue
    PropertyChanged "AlphaValue"
End Property
'
'Public Property Get CaptionControl() As Boolean
'    CaptionControl = m_CaptionControl
'End Property
'
'Public Property Let CaptionControl(ByVal New_CaptionControl As Boolean)
'    m_CaptionControl = New_CaptionControl
'    PropertyChanged "CaptionControl"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get FocusTitle() As Boolean
Attribute FocusTitle.VB_Description = "Sets Weather the titlebar Will ""PopUp"" To meet the mouse when the mouse is hovering over!"
    FocusTitle = m_FocusTitle
End Property

Public Property Let FocusTitle(ByVal New_FocusTitle As Boolean)
    m_FocusTitle = New_FocusTitle
    PropertyChanged "FocusTitle"
End Property

Public Property Get CaptionControl() As Boolean
Attribute CaptionControl.VB_Description = "Controls The tracking (active inactive ect.)"
Attribute CaptionControl.VB_MemberFlags = "400"
    CaptionControl = m_CaptionControl
End Property

Public Property Let CaptionControl(ByVal New_CaptionControl As Boolean)
    m_CaptionControl = New_CaptionControl
    PropertyChanged "CaptionControl"
End Property


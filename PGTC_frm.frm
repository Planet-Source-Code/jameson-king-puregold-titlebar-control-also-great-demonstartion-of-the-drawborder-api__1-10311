VERSION 5.00
Object = "{56F5C140-BB32-4175-9DFF-1CA82DF6B91B}#5.0#0"; "PureGoldTitleBarControl.ocx"
Begin VB.Form PGTC_frm 
   AutoRedraw      =   -1  'True
   Caption         =   "Even if the caption control Resides as a child of a frame or a picture box, the forms caption shows :)"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "PGTC_frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Parent Frame + Pic Box + Frame + Pic Box"
      Height          =   1365
      Left            =   30
      TabIndex        =   6
      Top             =   960
      Width           =   9945
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   60
         ScaleHeight     =   1035
         ScaleWidth      =   9735
         TabIndex        =   7
         Top             =   180
         Width           =   9795
         Begin VB.Frame Frame4 
            Caption         =   ":P"
            Height          =   1005
            Left            =   30
            TabIndex        =   8
            Top             =   30
            Width           =   9705
            Begin VB.PictureBox Picture2 
               Height          =   765
               Left            =   60
               ScaleHeight     =   705
               ScaleWidth      =   9495
               TabIndex        =   9
               Top             =   180
               Width           =   9555
               Begin PureGoldTitleBarControl.PGTC PGTC3 
                  Height          =   405
                  Left            =   90
                  TabIndex        =   10
                  Top             =   180
                  Width           =   9345
                  _ExtentX        =   16484
                  _ExtentY        =   714
                  AutoRedraw      =   0   'False
               End
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "With one Parent Frame"
      Height          =   585
      Left            =   30
      TabIndex        =   4
      Top             =   330
      Width           =   9945
      Begin PureGoldTitleBarControl.PGTC PGTC2 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   556
         AutoRedraw      =   0   'False
      End
   End
   Begin PureGoldTitleBarControl.PGTC PGTC1 
      Height          =   345
      Left            =   3990
      TabIndex        =   3
      Top             =   2880
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   609
      AutoRedraw      =   0   'False
      FocusTitle      =   -1  'True
   End
   Begin VB.Frame Frame3 
      Caption         =   "Extra examples of the draw caption API!"
      Height          =   1755
      Left            =   30
      TabIndex        =   0
      Top             =   2370
      Width           =   3825
      Begin VB.PictureBox cap1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   60
         ScaleHeight     =   405
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   390
         Width           =   3675
      End
      Begin VB.Label Label1 
         Caption         =   "Possible Look For a custom Start application Bar  ^  (Its Clickable)"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   3075
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Title Bar with TitleFocus Hover!  When the mouse is over the titlebar the title moves up to meet it"
      Height          =   375
      Left            =   3990
      TabIndex        =   11
      Top             =   2370
      Width           =   5385
   End
End
Attribute VB_Name = "PGTC_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Const MF_CHECKED = &H8&
Const MF_APPEND = &H100&
Const TPM_LEFTALIGN = &H0&
Const MF_DISABLED = &H2&
Const MF_GRAYED = &H1&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&
Private Type POINTAPI
    x As Long
    y As Long
End Type




Private Sub Command1_Click()

End Sub

Private Sub cap1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rt As RECT
Dim rtn As Long
        With rt
        .Left = 0
        .Top = 0
        .Right = cap1.ScaleX(cap1.ScaleWidth, cap1.ScaleMode, vbPixels)
        .Bottom = cap1.ScaleY(cap1.ScaleHeight, cap1.ScaleMode, vbPixels) '25
        End With
        DrawEdge cap1.hdc, rt, &HA, BF_RECT
        cap1.Refresh
End Sub

Private Sub cap1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim rt As RECT
Dim rtn As Long
        With rt
        .Left = 0
        .Top = 0
        .Right = cap1.ScaleX(cap1.ScaleWidth, cap1.ScaleMode, vbPixels)
        .Bottom = cap1.ScaleY(cap1.ScaleHeight, cap1.ScaleMode, vbPixels) '25
        End With
        DrawEdge cap1.hdc, rt, &H5, BF_RECT
        cap1.Refresh
End Sub

Private Sub Form_Load()

       
Dim rt As RECT

        With rt
        .Left = 0
        .Top = 0
        .Right = cap1.ScaleX(cap1.ScaleWidth, cap1.ScaleMode, vbPixels)
        .Bottom = cap1.ScaleY(cap1.ScaleHeight, cap1.ScaleMode, vbPixels) '25
        End With
        DrawCaption Me.hwnd, cap1.hdc, rt, &H14 Or &H18
        DrawEdge cap1.hdc, rt, &H5, BF_RECT
        


End Sub

Private Sub Form_Paint()
Me.Cls
Dim rt As RECT
Dim rtn As Long
        With rt
        .Left = 0
        .Top = 0
        .Right = Me.ScaleX(Me.ScaleWidth, Me.ScaleMode, vbPixels)
        .Bottom = Me.ScaleY(Me.ScaleHeight, Me.ScaleMode, vbPixels) '25
        End With
        DrawEdge Me.hdc, rt, &H2, BF_RECT
        
               With rt
        .Left = 0
        .Top = 0
        .Right = cap1.ScaleX(cap1.ScaleWidth, cap1.ScaleMode, vbPixels)
        .Bottom = cap1.ScaleY(cap1.ScaleHeight, cap1.ScaleMode, vbPixels) '25
        End With
        DrawCaption Me.hwnd, cap1.hdc, rt, &H14 Or &H18
        DrawEdge cap1.hdc, rt, &H5, BF_RECT
        

End Sub

Private Sub Form_Resize()
Form_Paint
End Sub


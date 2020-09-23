VERSION 5.00
Begin VB.UserControl UserControl1 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9075
   ControlContainer=   -1  'True
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3990
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   6
      Top             =   4860
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6495
      Top             =   2805
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   2040
      ScaleHeight     =   3405
      ScaleWidth      =   5535
      TabIndex        =   3
      Top             =   1275
      Width           =   5535
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1125
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   2
      ToolTipText     =   "Maximize/Restore"
      Top             =   195
      Width           =   225
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1395
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   195
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   735
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      ToolTipText     =   "Minimize"
      Top             =   240
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   240
      Stretch         =   -1  'True
      Top             =   285
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2325
      TabIndex        =   4
      Top             =   30
      Width           =   4320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2415
      TabIndex        =   5
      Top             =   285
      Width           =   4320
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const vbBorderColor = 4003095

Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const HTBOTTOMRIGHT = 17
Private Const HTBOTTOM = 15
Private Const HTBOTTOMLEFT = 16
Private Const HTLEFT = 10
Private Const HTRIGHT = 11
Private Const HTTOP = 12
Private Const HTTOPLEFT = 13
Private Const HTTOPRIGHT = 14



Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type FORMRECT
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim FORMRECT As FORMRECT
Dim Maximized As Boolean
Dim LastMouseOver As String





Private Function LoadLeft(X As Integer, Y As Integer)
    UserControl.Line (X + 4, Y + 0)-(X + 9, Y + 0), 3150094
    
    UserControl.Line (X + 0, Y + 5)-(X + 3, Y + 5), vbBorderColor
    UserControl.Line (X + 0, Y + 26)-(X + 9, Y + 26), 13482679
    UserControl.Line (X + 0, Y + 27)-(X + 9, Y + 27), 11444378
    UserControl.Line (X + 0, Y + 28)-(X + 9, Y + 28), vbBorderColor
    
    UserControl.Line (X + 3, Y + 1)-(X + 3, Y + 23), vbBorderColor
    UserControl.PSet (X + 3, Y + 23), 3413515
    UserControl.PSet (X + 3, Y + 24), 11441294
    UserControl.PSet (X + 3, Y + 25), 13551302
    
    UserControl.Line (X + 4, Y + 25)-(X + 9, Y + 25), 13484479
    UserControl.Line (X + 0, Y + 25)-(X + 3, Y + 25), 10921124
    
    UserControl.Line (X + 2, Y + 6)-(X + 2, Y + 26), 12563886
    
    UserControl.Line (X + 4, Y + 1)-(X + 4, Y + 24), 16442331
    
    UserControl.Line (X + 4, Y + 1)-(X + 9, Y + 1), 16768469
    
    UserControl.Line (X + 4, Y + 23)-(X + 9, Y + 23), 11238255
    
    UserControl.Line (X + 4, Y + 24)-(X + 9, Y + 24), vbBorderColor
    
    UserControl.Line (X + 0, Y + 6)-(X + 2, Y + 6), 16710907
    UserControl.Line (X + 0, Y + 7)-(X + 2, Y + 7), 14604497
    UserControl.Line (X + 0, Y + 8)-(X + 2, Y + 8), 14538704
    UserControl.Line (X + 0, Y + 9)-(X + 2, Y + 9), 14407118
    UserControl.Line (X + 0, Y + 10)-(X + 2, Y + 10), 14340811
    UserControl.Line (X + 0, Y + 11)-(X + 2, Y + 11), 14340297
    UserControl.Line (X + 0, Y + 12)-(X + 2, Y + 12), 14077125
    UserControl.Line (X + 0, Y + 13)-(X + 2, Y + 13), 13879233
    UserControl.Line (X + 0, Y + 14)-(X + 2, Y + 14), 13681854
    UserControl.Line (X + 0, Y + 15)-(X + 2, Y + 15), 13550781
    UserControl.Line (X + 0, Y + 16)-(X + 2, Y + 16), 13353402
    UserControl.Line (X + 0, Y + 17)-(X + 2, Y + 17), 13156023
    UserControl.Line (X + 0, Y + 18)-(X + 2, Y + 18), 12958644
    UserControl.Line (X + 0, Y + 19)-(X + 2, Y + 19), 12827058
    UserControl.Line (X + 0, Y + 20)-(X + 2, Y + 20), 12629679
    UserControl.Line (X + 0, Y + 21)-(X + 2, Y + 21), 12432300
    UserControl.Line (X + 0, Y + 22)-(X + 2, Y + 22), 12234921
    UserControl.Line (X + 0, Y + 23)-(X + 2, Y + 23), 12300711
    UserControl.Line (X + 0, Y + 24)-(X + 2, Y + 24), 12037031
    UserControl.Line (X + 0, Y + 25)-(X + 2, Y + 25), 10921124
    
    UserControl.Line (X + 5, Y + 2)-(X + 9, Y + 2), 12288888
    UserControl.Line (X + 5, Y + 3)-(X + 9, Y + 3), 12224891
    UserControl.Line (X + 5, Y + 4)-(X + 9, Y + 4), 12421500
    UserControl.Line (X + 5, Y + 5)-(X + 9, Y + 5), 12618879
    UserControl.Line (X + 5, Y + 6)-(X + 9, Y + 6), 12423298
    UserControl.Line (X + 5, Y + 7)-(X + 9, Y + 7), 12685956
    UserControl.Line (X + 5, Y + 8)-(X + 9, Y + 8), 12817542
    UserControl.Line (X + 5, Y + 9)-(X + 9, Y + 9), 12949128
    UserControl.Line (X + 5, Y + 10)-(X + 9, Y + 10), 13015435
    UserControl.Line (X + 5, Y + 11)-(X + 9, Y + 11), 13081228
    UserControl.Line (X + 5, Y + 12)-(X + 9, Y + 12), 13278607
    UserControl.Line (X + 5, Y + 13)-(X + 9, Y + 13), 13344913
    UserControl.Line (X + 5, Y + 14)-(X + 9, Y + 14), 13476499
    UserControl.Line (X + 5, Y + 15)-(X + 9, Y + 15), 13542551
    UserControl.Line (X + 5, Y + 16)-(X + 9, Y + 16), 13805723
    UserControl.Line (X + 5, Y + 17)-(X + 9, Y + 17), 14003102
    UserControl.Line (X + 5, Y + 18)-(X + 9, Y + 18), 13938079
    UserControl.Line (X + 5, Y + 19)-(X + 9, Y + 19), 14069665
    UserControl.Line (X + 5, Y + 20)-(X + 9, Y + 20), 14267044
    UserControl.Line (X + 5, Y + 21)-(X + 9, Y + 21), 14333351
    UserControl.Line (X + 5, Y + 22)-(X + 9, Y + 22), 14333351
    
End Function

Private Function LoadLeftFill(X As Integer, Y As Integer, Legnth As Integer)
    Dim Colors(28) As String
    Colors(0) = 0: Colors(1) = 0
    Colors(2) = 0: Colors(3) = 0
    Colors(4) = 0: Colors(5) = vbBorderColor
    Colors(6) = 16776953: Colors(7) = 14604497
    Colors(8) = 14538704: Colors(9) = 14407118
    Colors(10) = 14340811: Colors(11) = 14340297
    Colors(12) = 14077125: Colors(13) = 13879233
    Colors(14) = 13681854: Colors(15) = 13550781
    Colors(16) = 13353402: Colors(17) = 13156023
    Colors(18) = 12958644: Colors(19) = 12827058
    Colors(20) = 12629679: Colors(21) = 12432300
    Colors(22) = 12234921: Colors(23) = 12168874
    Colors(24) = 12103081: Colors(25) = 11839909
    Colors(26) = 13482679: Colors(27) = 11248026
    Colors(28) = vbBorderColor
    LineArray Colors, Legnth, X, Y
End Function

Private Function LoadLeftCorner(X As Integer, Y As Integer)
    UserControl.Line (X + 0, Y + 5)-(X + 7, Y + 5), vbBorderColor
    UserControl.Line (X + 0, Y + 28)-(X + 7, Y + 28), vbBorderColor
    
    UserControl.Line (X + 0, Y + 5)-(X + 0, Y + 28), vbBorderColor
    UserControl.Line (X + 1, Y + 6)-(X + 1, Y + 28), 16777211
    UserControl.Line (X + 6, Y + 6)-(X + 6, Y + 28), 13550781
    
    UserControl.PSet (X + 2, Y + 6), 16777211
    UserControl.PSet (X + 2, Y + 7), 14604497
    UserControl.PSet (X + 2, Y + 8), 14538706
    UserControl.PSet (X + 2, Y + 9), 14472913
    UserControl.PSet (X + 2, Y + 10), 14340811
    UserControl.PSet (X + 2, Y + 11), 13880776
    UserControl.PSet (X + 2, Y + 12), 13683397
    UserControl.PSet (X + 2, Y + 13), 13617604
    UserControl.PSet (X + 2, Y + 14), 13420225
    UserControl.PSet (X + 2, Y + 15), 13550781
    UserControl.PSet (X + 2, Y + 16), 13222076
    UserControl.PSet (X + 2, Y + 17), 13156283
    UserControl.PSet (X + 2, Y + 18), 12958644
    UserControl.PSet (X + 2, Y + 19), 12827058
    UserControl.PSet (X + 2, Y + 20), 12300974
    UserControl.PSet (X + 2, Y + 21), 12300974
    UserControl.PSet (X + 2, Y + 22), 12234921
    UserControl.PSet (X + 2, Y + 23), 12168874
    UserControl.PSet (X + 2, Y + 24), 11775400
    UserControl.PSet (X + 2, Y + 25), 11379872
    UserControl.PSet (X + 2, Y + 26), 13221816
    UserControl.PSet (X + 2, Y + 27), 11511195
    
    UserControl.PSet (X + 3, Y + 6), 16448758
    UserControl.PSet (X + 3, Y + 7), 14604497
    UserControl.PSet (X + 3, Y + 8), 2299667
    UserControl.PSet (X + 3, Y + 9), 11510687
    UserControl.PSet (X + 3, Y + 10), 14340811
    UserControl.PSet (X + 3, Y + 11), 14340297
    UserControl.PSet (X + 3, Y + 12), 2233874
    UserControl.PSet (X + 3, Y + 13), 11444894
    UserControl.PSet (X + 3, Y + 14), 13681854
    UserControl.PSet (X + 3, Y + 15), 13550781
    UserControl.PSet (X + 3, Y + 16), 2365460
    UserControl.PSet (X + 3, Y + 17), 11510687
    UserControl.PSet (X + 3, Y + 18), 12958644
    UserControl.PSet (X + 3, Y + 19), 12827058
    UserControl.PSet (X + 3, Y + 20), 2168081
    UserControl.PSet (X + 3, Y + 21), 11313308
    UserControl.PSet (X + 3, Y + 22), 12234921
    UserControl.PSet (X + 3, Y + 23), 12168874
    UserControl.PSet (X + 3, Y + 24), 2233874
    UserControl.PSet (X + 3, Y + 25), 11773343
    UserControl.PSet (X + 3, Y + 26), 13482679
    UserControl.PSet (X + 3, Y + 27), 11377556
    
    UserControl.PSet (X + 4, Y + 6), 16646139
    UserControl.PSet (X + 4, Y + 7), 14604497
    UserControl.PSet (X + 4, Y + 8), 11247772
    UserControl.PSet (X + 4, Y + 9), 16708591
    UserControl.PSet (X + 4, Y + 10), 14340811
    UserControl.PSet (X + 4, Y + 11), 14340297
    UserControl.PSet (X + 4, Y + 12), 11379358
    UserControl.PSet (X + 4, Y + 13), 16774641
    UserControl.PSet (X + 4, Y + 14), 13681854
    UserControl.PSet (X + 4, Y + 15), 13550781
    UserControl.PSet (X + 4, Y + 16), 11379358
    UserControl.PSet (X + 4, Y + 17), 16774641
    UserControl.PSet (X + 4, Y + 18), 12958644
    UserControl.PSet (X + 4, Y + 19), 12827058
    UserControl.PSet (X + 4, Y + 20), 11313565
    UserControl.PSet (X + 4, Y + 21), 16774384
    UserControl.PSet (X + 4, Y + 22), 12234921
    UserControl.PSet (X + 4, Y + 23), 12168874
    UserControl.PSet (X + 4, Y + 24), 10721942
    UserControl.PSet (X + 4, Y + 25), 16776695
    UserControl.PSet (X + 4, Y + 26), 13482679
    UserControl.PSet (X + 4, Y + 27), 11708832
    
    UserControl.PSet (X + 5, Y + 6), 16646137
    UserControl.PSet (X + 5, Y + 7), 14604497
    UserControl.PSet (X + 5, Y + 8), 14538704
    UserControl.PSet (X + 5, Y + 9), 14407118
    UserControl.PSet (X + 5, Y + 10), 14340811
    UserControl.PSet (X + 5, Y + 11), 14340297
    UserControl.PSet (X + 5, Y + 12), 14077125
    UserControl.PSet (X + 5, Y + 13), 13879233
    UserControl.PSet (X + 5, Y + 14), 13681854
    UserControl.PSet (X + 5, Y + 15), 13550781
    UserControl.PSet (X + 5, Y + 16), 13353402
    UserControl.PSet (X + 5, Y + 17), 13156023
    UserControl.PSet (X + 5, Y + 18), 12958644
    UserControl.PSet (X + 5, Y + 19), 12827058
    UserControl.PSet (X + 5, Y + 20), 12629679
    UserControl.PSet (X + 5, Y + 21), 12432300
    UserControl.PSet (X + 5, Y + 22), 12234921
    UserControl.PSet (X + 5, Y + 23), 12168874
    UserControl.PSet (X + 5, Y + 24), 12103081
    UserControl.PSet (X + 5, Y + 25), 11839909
    UserControl.PSet (X + 5, Y + 26), 13482679
    UserControl.PSet (X + 5, Y + 27), 11379099
    
    UserControl.PSet (X + 6, Y + 6), 16777211
    UserControl.PSet (X + 6, Y + 7), 14604497
    UserControl.PSet (X + 6, Y + 8), 14538704
    UserControl.PSet (X + 6, Y + 9), 14407118
    UserControl.PSet (X + 6, Y + 10), 14340811
    UserControl.PSet (X + 6, Y + 11), 14340297
    UserControl.PSet (X + 6, Y + 12), 14077125
    UserControl.PSet (X + 6, Y + 13), 13879233
    UserControl.PSet (X + 6, Y + 14), 13681854
    UserControl.PSet (X + 6, Y + 15), 13550781
    UserControl.PSet (X + 6, Y + 16), 13353402
    UserControl.PSet (X + 6, Y + 17), 13156023
    UserControl.PSet (X + 6, Y + 18), 12958644
    UserControl.PSet (X + 6, Y + 19), 12827058
    UserControl.PSet (X + 6, Y + 20), 12629679
    UserControl.PSet (X + 6, Y + 21), 12432300
    UserControl.PSet (X + 6, Y + 22), 12234921
    UserControl.PSet (X + 6, Y + 23), 12168874
    UserControl.PSet (X + 6, Y + 24), 12103081
    UserControl.PSet (X + 6, Y + 25), 11839909
    UserControl.PSet (X + 6, Y + 26), 13482679
    UserControl.PSet (X + 6, Y + 27), 11117468
End Function

Function LoadTitleFiller(X As Integer, Y As Integer, Legnth As Integer)
    Dim Colors(28) As String
    Colors(0) = 4067597: Colors(1) = 16769751
    Colors(2) = 12288888: Colors(3) = 12421500
    Colors(4) = 12421500: Colors(5) = 12618879
    Colors(6) = 12423298: Colors(7) = 12685956
    Colors(8) = 12817542: Colors(9) = 12949128
    Colors(10) = 13015435: Colors(11) = 13081228
    Colors(12) = 13278607: Colors(13) = 13344913
    Colors(14) = 13476499: Colors(15) = 13542551
    Colors(16) = 13805723: Colors(17) = 14003102
    Colors(18) = 13938079: Colors(19) = 14069665
    Colors(20) = 14267044: Colors(21) = 14333351
    Colors(22) = 14333351: Colors(23) = 11106669
    Colors(24) = 4264205: Colors(25) = 13682627
    Colors(26) = 13482679: Colors(27) = 11117723
    Colors(28) = 3540739
    LineArray Colors, Legnth, X, Y
End Function

Function LoadTitleEnd(X As Integer, Y As Integer)
    
    UserControl.Line (X + 16, Y + 8)-(X + 28, Y + 8), vbBorderColor
    
    UserControl.Line (X + 20, Y + 13)-(X + 28, Y + 13), vbBorderColor
    
    Dim i
    For i = 0 To 7
        UserControl.PSet (X + 13 + i, Y + 6 + i), vbBorderColor
    Next
    
    UserControl.Line (X + 7, Y + 5)-(X + 13, Y + 5), vbBorderColor
    
    UserControl.Line (X + 0, Y + 0)-(X + 6, Y + 0), vbBorderColor
    
    UserControl.Line (X + 5, Y + 0)-(X + 5, Y + 24), 11434605
    UserControl.Line (X + 0, Y + 23)-(X + 6, Y + 23), 11434605
    
    UserControl.PSet (X + 6, Y + 23), 3543817
    UserControl.Line (X + 6, Y + 1)-(X + 6, Y + 23), vbBorderColor
    
    UserControl.PSet (X + 7, Y + 24), 14271936
    UserControl.Line (X + 0, Y + 24)-(X + 7, Y + 24), vbBorderColor
    UserControl.Line (X + 6, Y + 24)-(X + 28, Y + 24), 12103081
    
    UserControl.Line (X + 0, Y + 25)-(X + 28, Y + 25), 10790821
    UserControl.Line (X + 0, Y + 26)-(X + 28, Y + 26), 13482679
    UserControl.Line (X + 0, Y + 27)-(X + 28, Y + 27), 11117723
    UserControl.Line (X + 0, Y + 28)-(X + 28, Y + 28), vbBorderColor
    
    UserControl.Line (X + 0, Y + 1)-(X + 5, Y + 1), 16638422
    UserControl.Line (X + 0, Y + 2)-(X + 5, Y + 2), 12288888
    UserControl.Line (X + 0, Y + 3)-(X + 5, Y + 3), 12224891
    UserControl.Line (X + 0, Y + 4)-(X + 5, Y + 4), 12421754
    UserControl.Line (X + 0, Y + 5)-(X + 5, Y + 5), 12619133
    UserControl.Line (X + 0, Y + 6)-(X + 5, Y + 6), 12357759
    UserControl.Line (X + 0, Y + 7)-(X + 5, Y + 7), 12685956
    UserControl.Line (X + 0, Y + 8)-(X + 5, Y + 8), 12817542
    UserControl.Line (X + 0, Y + 9)-(X + 5, Y + 9), 12949128
    UserControl.Line (X + 0, Y + 10)-(X + 5, Y + 10), 13015435
    UserControl.Line (X + 0, Y + 11)-(X + 5, Y + 11), 13081228
    UserControl.Line (X + 0, Y + 12)-(X + 5, Y + 12), 13278607
    UserControl.Line (X + 0, Y + 13)-(X + 5, Y + 13), 13344913
    UserControl.Line (X + 0, Y + 14)-(X + 5, Y + 14), 13476499
    UserControl.Line (X + 0, Y + 15)-(X + 5, Y + 15), 13542551
    UserControl.Line (X + 0, Y + 16)-(X + 5, Y + 16), 13805723
    UserControl.Line (X + 0, Y + 17)-(X + 5, Y + 17), 14003102
    UserControl.Line (X + 0, Y + 18)-(X + 5, Y + 18), 13938079
    UserControl.Line (X + 0, Y + 19)-(X + 5, Y + 19), 14069665
    UserControl.Line (X + 0, Y + 20)-(X + 5, Y + 20), 14267044
    UserControl.Line (X + 0, Y + 21)-(X + 5, Y + 21), 14333351
    UserControl.Line (X + 0, Y + 22)-(X + 5, Y + 22), 14333351
    UserControl.Line (X + 0, Y + 23)-(X + 5, Y + 23), 11106669
    
    UserControl.Line (X + 7, Y + 6)-(X + 7, Y + 24), 11572110
    
    UserControl.Line (X + 8, Y + 13)-(X + 20, Y + 13), 13945026
    UserControl.Line (X + 8, Y + 12)-(X + 19, Y + 12), 14208198
    UserControl.Line (X + 8, Y + 11)-(X + 18, Y + 11), 14274504
    UserControl.Line (X + 8, Y + 10)-(X + 17, Y + 10), 14340811
    UserControl.Line (X + 8, Y + 9)-(X + 16, Y + 9), 14472397
    UserControl.Line (X + 8, Y + 8)-(X + 15, Y + 8), 14538704
    UserControl.Line (X + 8, Y + 7)-(X + 14, Y + 7), 14670290
    UserControl.Line (X + 8, Y + 6)-(X + 13, Y + 6), 16777212
    
    UserControl.Line (X + 8, Y + 14)-(X + 28, Y + 14), 13813440
    UserControl.Line (X + 8, Y + 15)-(X + 28, Y + 15), 13484988
    UserControl.Line (X + 8, Y + 16)-(X + 28, Y + 16), 13353402
    UserControl.Line (X + 8, Y + 17)-(X + 28, Y + 17), 13221816
    UserControl.Line (X + 8, Y + 18)-(X + 28, Y + 18), 12893365
    UserControl.Line (X + 8, Y + 19)-(X + 28, Y + 19), 12695986
    UserControl.Line (X + 8, Y + 20)-(X + 28, Y + 20), 12564400
    UserControl.Line (X + 8, Y + 21)-(X + 28, Y + 21), 12235949
    UserControl.Line (X + 8, Y + 22)-(X + 28, Y + 22), 12170156
    UserControl.Line (X + 8, Y + 23)-(X + 28, Y + 23), 12693925
    
    UserControl.Line (X + 17, Y + 9)-(X + 28, Y + 9), 16777211
    UserControl.Line (X + 18, Y + 10)-(X + 28, Y + 10), 15320509
    UserControl.Line (X + 19, Y + 11)-(X + 28, Y + 11), 14926006
    UserControl.Line (X + 20, Y + 12)-(X + 28, Y + 12), 13412767
    
    For i = 0 To 4
        UserControl.PSet (X + 14 + i, Y + 9 + i), 16777211
    Next
End Function
Function LoadButtonTop(X As Integer, Y As Integer)
    UserControl.Line (X + 0, Y + 8)-(X + 15, Y + 8), vbBorderColor
    UserControl.Line (X + 0, Y + 9)-(X + 15, Y + 9), 16771032
End Function
Function LoadButton1(X As Integer, Y As Integer)
    Picture1.Line (X + 0, Y + 0)-(X + 15, Y + 0), vbBorderColor
    Picture1.Line (X + 0, Y + 0)-(X + 0, Y + 15), vbBorderColor
    Picture1.Line (X + 14, Y + 0)-(X + 14, Y + 15), vbBorderColor
    Picture1.Line (X + 0, Y + 14)-(X + 15, Y + 14), vbBorderColor
    Picture1.Line (X + 0, Y + 9)-(X + 15, Y + 9), vbBorderColor
    
    Picture1.Line (X + 1, Y + 10)-(X + 14, Y + 10), 16776959
    Picture1.Line (X + 1, Y + 10)-(X + 1, Y + 14), 16776959
    
    Picture1.Line (X + 2, Y + 11)-(X + 14, Y + 11), 15131368
    Picture1.Line (X + 2, Y + 12)-(X + 14, Y + 12), 15131368
    Picture1.Line (X + 2, Y + 13)-(X + 14, Y + 13), 15131368
    
    Picture1.Line (X + 2, Y + 1)-(X + 13, Y + 1), 10452597
    
    Picture1.Line (X + 2, Y + 1)-(X + 13, Y + 1), 10319470
    Picture1.Line (X + 2, Y + 2)-(X + 13, Y + 2), 13208453
    Picture1.Line (X + 2, Y + 3)-(X + 13, Y + 3), 12751239
    Picture1.Line (X + 2, Y + 4)-(X + 13, Y + 4), 13409169
    Picture1.Line (X + 2, Y + 5)-(X + 13, Y + 5), 13673884
    Picture1.Line (X + 2, Y + 6)-(X + 13, Y + 6), 14264733
    Picture1.Line (X + 2, Y + 7)-(X + 13, Y + 7), 14989226
    Picture1.Line (X + 2, Y + 8)-(X + 13, Y + 8), 13873833
    
    Picture1.PSet (X + 1, Y + 1), 10452597
    Picture1.PSet (X + 1, Y + 2), 12223358
    Picture1.PSet (X + 1, Y + 3), 12751759
    Picture1.PSet (X + 1, Y + 4), 12161934
    Picture1.PSet (X + 1, Y + 5), 12031890
    Picture1.PSet (X + 1, Y + 6), 12886938
    Picture1.PSet (X + 1, Y + 7), 13347744
    Picture1.PSet (X + 1, Y + 8), 13284012
    Picture1.PSet (X + 13, Y + 1), 9665391
    Picture1.PSet (X + 13, Y + 2), 12291976
    Picture1.PSet (X + 13, Y + 3), 13472915
    Picture1.PSet (X + 13, Y + 4), 12230555
    Picture1.PSet (X + 13, Y + 5), 12888221
    Picture1.PSet (X + 13, Y + 6), 13151651
    Picture1.PSet (X + 13, Y + 7), 13087402
    Picture1.PSet (X + 13, Y + 8), 12826031
End Function
Function LoadButtonBottom(X As Integer, Y As Integer)
    UserControl.Line (X + 0, Y + 0)-(X + 15, Y + 0), 16776954
    UserControl.Line (X + 0, Y + 1)-(X + 15, Y + 1), 13482679
    UserControl.Line (X + 0, Y + 2)-(X + 15, Y + 2), 11117723
    UserControl.Line (X + 0, Y + 3)-(X + 15, Y + 3), vbBorderColor
End Function
Function LoadSpace1(X As Integer, Y As Integer)
    UserControl.Line (X + 0, Y + 8)-(X + 7, Y + 8), vbBorderColor
    UserControl.Line (X + 0, Y + 9)-(X + 7, Y + 9), 16771032
    UserControl.Line (X + 0, Y + 10)-(X + 7, Y + 10), 15452347
    UserControl.Line (X + 0, Y + 11)-(X + 7, Y + 11), 14926006
    UserControl.Line (X + 0, Y + 12)-(X + 7, Y + 12), 13412767
    UserControl.Line (X + 0, Y + 13)-(X + 7, Y + 13), 4331284
    UserControl.Line (X + 0, Y + 14)-(X + 7, Y + 14), 16777211
    UserControl.Line (X + 0, Y + 15)-(X + 7, Y + 15), 13550527
    UserControl.Line (X + 0, Y + 16)-(X + 7, Y + 16), 13353148
    UserControl.Line (X + 0, Y + 17)-(X + 7, Y + 17), 13155769
    UserControl.Line (X + 0, Y + 18)-(X + 7, Y + 18), 12958390
    UserControl.Line (X + 0, Y + 19)-(X + 7, Y + 19), 12892597
    UserControl.Line (X + 0, Y + 20)-(X + 7, Y + 20), 12695218
    UserControl.Line (X + 0, Y + 21)-(X + 7, Y + 21), 12497839
    UserControl.Line (X + 0, Y + 22)-(X + 7, Y + 22), 12300460
    UserControl.Line (X + 0, Y + 23)-(X + 7, Y + 23), 12168874
    UserControl.PSet (X + 0, Y + 24), 16776186
    UserControl.PSet (X + 0, Y + 25), 16776186
    UserControl.Line (X + 1, Y + 24)-(X + 7, Y + 24), 12103081
    UserControl.Line (X + 1, Y + 25)-(X + 7, Y + 25), 11839909
    
    UserControl.Line (X + 0, Y + 26)-(X + 7, Y + 26), 13482679
    UserControl.Line (X + 0, Y + 27)-(X + 7, Y + 27), 11117723
    UserControl.Line (X + 0, Y + 28)-(X + 7, Y + 28), vbBorderColor
End Function
Function LoadSpace2(X As Integer, Y As Integer)
    UserControl.Line (X + 0, Y + 8)-(X + 7, Y + 8), vbBorderColor
    UserControl.Line (X + 0, Y + 28)-(X + 7, Y + 28), vbBorderColor
    Dim i: For i = 0 To 2
        UserControl.PSet (X + 0 + i, Y + 16 + i), vbBorderColor
    Next
    UserControl.Line (X + 2, Y + 18)-(X + 7, Y + 18), vbBorderColor
    UserControl.Line (X + 0, Y + 9)-(X + 7, Y + 9), 16771032
    UserControl.Line (X + 0, Y + 10)-(X + 7, Y + 10), 15452347
    UserControl.Line (X + 0, Y + 11)-(X + 7, Y + 11), 14926006
    UserControl.Line (X + 0, Y + 12)-(X + 7, Y + 12), 14070432
    
    UserControl.PSet (X + 0, Y + 13), 13477013
    UserControl.PSet (X + 0, Y + 14), 13739671
    UserControl.Line (X + 1, Y + 13)-(X + 7, Y + 13), 14397346
    UserControl.Line (X + 1, Y + 14)-(X + 7, Y + 14), 14068124
    
    UserControl.PSet (X + 0, Y + 15), 10647412
    UserControl.PSet (X + 1, Y + 15), 13738136
    UserControl.PSet (X + 1, Y + 16), 10975602
    UserControl.Line (X + 2, Y + 15)-(X + 7, Y + 15), 13474956
    UserControl.Line (X + 2, Y + 16)-(X + 7, Y + 16), 13079945
    UserControl.Line (X + 2, Y + 17)-(X + 7, Y + 17), 10844015
    
    UserControl.PSet (X + 0, Y + 17), 16508644
    UserControl.PSet (X + 0, Y + 18), 16776956
    UserControl.PSet (X + 1, Y + 18), 16639716
    UserControl.PSet (X + 2, Y + 17), 10844015
    
    UserControl.Line (X + 0, Y + 19)-(X + 7, Y + 19), 16777209
    UserControl.Line (X + 0, Y + 20)-(X + 7, Y + 20), 12695218
    UserControl.Line (X + 0, Y + 21)-(X + 7, Y + 21), 12497839
    UserControl.Line (X + 0, Y + 22)-(X + 7, Y + 22), 12300460
    UserControl.Line (X + 0, Y + 23)-(X + 7, Y + 23), 12168874
    UserControl.Line (X + 0, Y + 24)-(X + 7, Y + 24), 12103081
    UserControl.Line (X + 0, Y + 25)-(X + 7, Y + 25), 11839909
    UserControl.Line (X + 0, Y + 26)-(X + 7, Y + 26), 13482679
    UserControl.Line (X + 0, Y + 27)-(X + 7, Y + 27), 11117723
    
    UserControl.PSet (X + 0, Y + 24), 16579583
    UserControl.PSet (X + 0, Y + 25), 16776703
    UserControl.PSet (X + 0, Y + 23), 16777215
End Function
Function LoadButton3Top(X As Integer, Y As Integer)
    UserControl.PSet (X + 0, Y + 8), 3937814
    UserControl.PSet (X + 0, Y + 9), 16771032
    UserControl.PSet (X + 1, Y + 8), 4002572
    UserControl.PSet (X + 1, Y + 9), 16771032
    UserControl.PSet (X + 2, Y + 8), 3936779
    UserControl.PSet (X + 2, Y + 9), 16771032
    UserControl.PSet (X + 3, Y + 8), 3870986
    UserControl.PSet (X + 3, Y + 9), 16771032
    UserControl.PSet (X + 4, Y + 8), 3870986
    UserControl.PSet (X + 4, Y + 9), 16771032
    UserControl.PSet (X + 5, Y + 8), 3870986
    UserControl.PSet (X + 5, Y + 9), 16771032
    UserControl.PSet (X + 6, Y + 8), 3805193
    UserControl.PSet (X + 6, Y + 9), 16771032
    UserControl.PSet (X + 7, Y + 8), 3870986
    UserControl.PSet (X + 7, Y + 9), 16771032
    UserControl.PSet (X + 8, Y + 8), 3936779
    UserControl.PSet (X + 8, Y + 9), 16771032
    UserControl.PSet (X + 9, Y + 8), 3412746
    UserControl.PSet (X + 9, Y + 9), 16771032
    UserControl.PSet (X + 10, Y + 8), 3609871
    UserControl.PSet (X + 10, Y + 9), 16771032
    UserControl.PSet (X + 11, Y + 8), 2821129
    UserControl.PSet (X + 11, Y + 9), 16771032
    UserControl.PSet (X + 12, Y + 8), 3347732
    UserControl.PSet (X + 12, Y + 9), 10188653
    UserControl.PSet (X + 13, Y + 9), 3609355
End Function
Function LoadRightCorner(X As Integer, Y As Integer)
    Dim i: For i = 0 To 7
    UserControl.PSet (X + 0 + i, Y + 11 + i), vbBorderColor
    UserControl.PSet (X + 0 + i, Y + 12 + i), 10188653
    Next
    

    UserControl.PSet (X + 0, Y + 13), 14398629
    UserControl.Line (X + 0, Y + 17)-(X + 5, Y + 17), 9926767
    UserControl.Line (X + 0, Y + 16)-(X + 4, Y + 16), 12160645
    UserControl.Line (X + 0, Y + 15)-(X + 3, Y + 15), 13803927
    UserControl.Line (X + 0, Y + 14)-(X + 2, Y + 14), 14135197
    UserControl.Line (X + 10, Y + 20)-(X + 10, Y + 29), 12296862
    UserControl.Line (X + 0, Y + 19)-(X + 10, Y + 19), 16777209
    UserControl.Line (X + 0, Y + 20)-(X + 8, Y + 20), 12695218
    UserControl.Line (X + 0, Y + 21)-(X + 8, Y + 21), 12497839
    UserControl.Line (X + 1, Y + 22)-(X + 8, Y + 22), 12300460
    UserControl.Line (X + 1, Y + 23)-(X + 10, Y + 23), 12168874
    UserControl.Line (X + 1, Y + 24)-(X + 10, Y + 24), 12103081
    UserControl.Line (X + 1, Y + 25)-(X + 8, Y + 25), 11839909
    
    UserControl.PSet (X + 0, Y + 22), 16776959: UserControl.PSet (X + 0, Y + 23), 16777212
    UserControl.PSet (X + 0, Y + 24), 16776958: UserControl.PSet (X + 0, Y + 25), 16776703
    
    UserControl.Line (X + 0, Y + 26)-(X + 8, Y + 26), 13482679: UserControl.PSet (X + 8, Y + 26), 11182747
    UserControl.Line (X + 0, Y + 27)-(X + 9, Y + 27), 11117723
    UserControl.Line (X + 0, Y + 28)-(X + 9, Y + 28), vbBorderColor
    UserControl.Line (X + 0, Y + 18)-(X + 11, Y + 18), vbBorderColor
    
    UserControl.PSet (X + 8, Y + 21), 2036240
    UserControl.PSet (X + 8, Y + 20), 12695218
    UserControl.PSet (X + 8, Y + 22), 11510685
    UserControl.PSet (X + 8, Y + 25), 1905934
    UserControl.PSet (X + 8, Y + 26), 11182747
    
    UserControl.PSet (X + 9, Y + 20), 12695218
    UserControl.PSet (X + 9, Y + 21), 11707811
    UserControl.PSet (X + 9, Y + 22), 16774125
    UserControl.PSet (X + 9, Y + 25), 11051161
    UserControl.PSet (X + 9, Y + 26), 16775152
    UserControl.PSet (X + 9, Y + 26), 16775152
    UserControl.PSet (X + 9, Y + 27), 12958644
    UserControl.PSet (X + 9, Y + 28), 14857141

    UserControl.Line (X + 11, Y + 18)-(X + 11, Y + 29), vbBorderColor
End Function

Function LoadButton2(X As Integer, Y As Integer)
    Picture3.Line (X + 0, Y + 0)-(X + 15, Y + 0), vbBorderColor
    Picture3.Line (X + 0, Y + 0)-(X + 0, Y + 15), vbBorderColor
    Picture3.Line (X + 14, Y + 0)-(X + 14, Y + 15), vbBorderColor
    Picture3.Line (X + 0, Y + 14)-(X + 15, Y + 14), vbBorderColor
    Picture3.Line (X + 0, Y + 5)-(X + 15, Y + 5), vbBorderColor
    
    Picture3.Line (X + 1, Y + 1)-(X + 14, Y + 1), 16776959
    Picture3.Line (X + 1, Y + 1)-(X + 1, Y + 5), 16776959
    
    Picture3.Line (X + 2, Y + 2)-(X + 14, Y + 2), 15131368
    Picture3.Line (X + 2, Y + 3)-(X + 14, Y + 3), 15131368
    Picture3.Line (X + 2, Y + 4)-(X + 14, Y + 4), 15131368
    
    Picture3.Line (X + 2, Y + 6)-(X + 13, Y + 6), 10452597
    
    Picture3.Line (X + 2, Y + 6)-(X + 13, Y + 6), 10319470
    Picture3.Line (X + 2, Y + 7)-(X + 13, Y + 7), 13208453
    Picture3.Line (X + 2, Y + 8)-(X + 13, Y + 8), 12751239
    Picture3.Line (X + 2, Y + 9)-(X + 13, Y + 9), 13409169
    Picture3.Line (X + 2, Y + 10)-(X + 13, Y + 10), 13673884
    Picture3.Line (X + 2, Y + 11)-(X + 13, Y + 11), 14264733
    Picture3.Line (X + 2, Y + 12)-(X + 13, Y + 12), 14989226
    Picture3.Line (X + 2, Y + 13)-(X + 13, Y + 13), 13873833
    
    Picture3.PSet (X + 1, Y + 6), 10452597
    Picture3.PSet (X + 1, Y + 7), 12223358
    Picture3.PSet (X + 1, Y + 8), 12751759
    Picture3.PSet (X + 1, Y + 9), 12161934
    Picture3.PSet (X + 1, Y + 10), 12031890
    Picture3.PSet (X + 1, Y + 11), 12886938
    Picture3.PSet (X + 1, Y + 12), 13347744
    Picture3.PSet (X + 1, Y + 13), 13284012
    
    Picture3.PSet (X + 13, Y + 6), 9665391
    Picture3.PSet (X + 13, Y + 7), 12291976
    Picture3.PSet (X + 13, Y + 8), 13472915
    Picture3.PSet (X + 13, Y + 9), 12230555
    Picture3.PSet (X + 13, Y + 10), 12888221
    Picture3.PSet (X + 13, Y + 11), 13151651
    Picture3.PSet (X + 13, Y + 12), 13087402
    Picture3.PSet (X + 13, Y + 13), 12826031
End Function
Function LoadSideBarLeft(X As Integer, Y As Integer, Legnth As Integer)
    UserControl.Line (X + 0, Y + 0)-(X + 0, Y + 0 + Legnth), vbBorderColor
    UserControl.Line (X + 1, Y + 0)-(X + 1, Y + 0 + Legnth), 16771547
    UserControl.Line (X + 2, Y + 0)-(X + 2, Y + 0 + Legnth), 9857121
    UserControl.Line (X + 3, Y + 0)-(X + 3, Y + 0 + Legnth), vbBorderColor
    
    Dim i: For i = 0 To 2
    UserControl.Line (X + 4 + i, Legnth - 4 + Y)-(X + 45 + i, Legnth - 4 + Y), vbBorderColor
    UserControl.Line (X + 1 + i, Legnth - 3 + Y)-(X + 45 + i, Legnth - 3 + Y), 16771035
    UserControl.Line (X + 1 + i, Legnth - 2 + Y)-(X + 45 + i, Legnth - 2 + Y), 9857121
    UserControl.Line (X + 1 + i, Legnth - 1 + Y)-(X + 45 + i, Legnth - 1 + Y), vbBorderColor
    Next
    
End Function
Function LoadSideBarRight(X As Integer, Y As Integer, Legnth As Integer, Width As Integer)
    UserControl.Line (X + 0, Y + 0)-(X + 0, Y + 0 + Legnth), vbBorderColor
    UserControl.Line (X + 1, Y + 0)-(X + 1, Y + 0 + Legnth), 16250349
    UserControl.Line (X + 2, Y + 0)-(X + 2, Y + 0 + Legnth), 11510685
    UserControl.Line (X + 3, Y + 0)-(X + 3, Y + 0 + Legnth), vbBorderColor
    
    Dim i: For i = 0 To 2
    UserControl.Line (X - 3 + i, Legnth - 4 + Y)-(X - Width + i, Legnth - 4 + Y), vbBorderColor
    UserControl.Line (X - 1 + i, Legnth - 3 + Y)-(X - Width + i, Legnth - 3 + Y), 16250349
    UserControl.Line (X + 0 + i, Legnth - 2 + Y)-(X - Width + i, Legnth - 2 + Y), 11510685
    UserControl.Line (X + 1 + i, Legnth - 1 + Y)-(X - Width + i, Legnth - 1 + Y), vbBorderColor
    Next
End Function
Function LoadBorderCap(X As Integer, Y As Integer)
    UserControl.Line (X + 0, Y + 0)-(X + 5, Y + 0), 4527377
    UserControl.PSet (X + 0, Y + 1), 16703954
    UserControl.PSet (X + 0, Y + 2), 9331285
    UserControl.Line (X + 0, Y + 3)-(X + 5, Y + 3), 4658963

    UserControl.PSet (X + 1, Y + 1), 3413768
    UserControl.PSet (X + 1, Y + 2), 9462871

    UserControl.PSet (X + 2, Y + 1), 16774637
    UserControl.PSet (X + 2, Y + 2), 8085851
    UserControl.PSet (X + 2, Y + 3), 4656653

    UserControl.PSet (X + 3, Y + 1), 16511721
    UserControl.PSet (X + 3, Y + 2), 3086097

    UserControl.PSet (X + 4, Y + 1), 16709614
    UserControl.PSet (X + 4, Y + 2), 16773610
End Function
Function LoadDragDot(X As Integer, Y As Integer)
    Picture5.PSet (X + 0, Y + 0), 16383997
    Picture5.PSet (X + 0, Y + 1), 12634052
    Picture5.PSet (X + 1, Y + 0), 13025728
    Picture5.PSet (X + 1, Y + 1), 393472
    Picture5.PSet (X + 1, Y + 2), 16117999
    Picture5.PSet (X + 2, Y + 1), 16777214
    Picture5.PSet (X + 2, Y + 2), 16776186
End Function
Function LoadRest(X As Integer, Y As Integer)

End Function

Function LineArray(ColorArray() As String, Legnth As Integer, X As Integer, Y As Integer)
    Dim P
    For P = LBound(ColorArray) To UBound(ColorArray)
        If ColorArray(P) <> 0 Then UserControl.Line (X + 0, Y + P)-(X + Legnth, Y + P), ColorArray(P)
    Next
End Function

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag UserControl.Parent
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 1
End Sub

Private Sub Picture1_Click()
    UserControl.Parent.WindowState = 1
End Sub

Private Sub Picture2_Click()
    Unload UserControl.Parent
End Sub

Private Sub Picture3_Click()
    If Maximized <> True Then
        FORMRECT.Top = UserControl.Parent.Top
        FORMRECT.Left = UserControl.Parent.Left
        FORMRECT.Width = UserControl.Parent.Width
        FORMRECT.Height = UserControl.Parent.Height
        UserControl.Parent.Top = 30
        UserControl.Parent.Left = 0
        UserControl.Parent.Width = Screen.Width
        UserControl.Parent.Height = Screen.Height - GetTaskbarHeight - 36
        DoEvents
        LoadGui
        Maximized = True
    Else
        UserControl.Parent.Top = FORMRECT.Top
        UserControl.Parent.Left = FORMRECT.Left
        UserControl.Parent.Width = FORMRECT.Width
        UserControl.Parent.Height = FORMRECT.Height
        DoEvents
        LoadGui
        Maximized = False
    End If
End Sub



Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 8
If Button = 1 Then
        ReleaseCapture
        SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
        LoadGui
End If
End Sub

Private Sub Timer1_Timer()
    UserControl.Width = Parent.Width
    UserControl.Height = Parent.Height
    Label1.Caption = UserControl.Parent.Caption
    Label2.Caption = UserControl.Parent.Caption
    Image1.Picture = UserControl.Parent.Icon
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If UCase(LastMouseOver) = "TOP" Then FormDrag UserControl.Parent
End Sub

Private Sub FormDrag(frm As Form)
    If Maximized <> True Then
        ReleaseCapture
        Call SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Function LoadGui()
    On Error Resume Next
    UserControl.Cls
    Label1.Caption = UserControl.Parent.Caption
    Label2.Caption = UserControl.Parent.Caption
    Image1.Picture = UserControl.Parent.Icon
    Image1.Top = 6
    Image1.Left = 9
    
    Parent.ScaleMode = 3
    Parent.BackColor = &HFF00FF
    UserControl.Parent.BorderStyle = 0
    UserControl.Width = Parent.Width
    UserControl.Height = Parent.Height
    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Top = 0
    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Left = 0
    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Align = 1
    UserControl.BackColor = &HFF00FF
    Dim Ret As Long
    Ret = GetWindowLong(Parent.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Parent.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Parent.hWnd, &HFF00FF, 0, LWA_COLORKEY
    
    Dim He As Integer, Wi As Integer: He = UserControl.ScaleHeight: Wi = UserControl.ScaleWidth

    Wi = Wi - 99
    LoadLeftCorner 0, 0
    LoadLeftFill 7, 0, 20
    LoadLeft 27, 0
    
    LoadTitleFiller 35, 0, Wi - 35
    
    LoadTitleEnd Wi, 0
    LoadButtonTop Wi + 28, 0
    Picture1.Top = 10
    Picture1.Left = Wi + 28
    LoadButton1 0, 0
    LoadButtonBottom Wi + 28, 25
    LoadSpace1 Wi + 28 + 15, 0
    LoadButtonTop Wi + 28 + 15 + 7, 0
    LoadButtonBottom Wi + 28 + 15 + 7, 25
    LoadButton2 0, 0
    Picture3.Top = 10
    Picture3.Left = Wi + 28 + 15 + 7
    LoadSpace2 Wi + 28 + 15 + 7 + 15, 0
    LoadButtonBottom Wi - 16 - 11 + 99, 25
    LoadButton3Top Wi - 16 - 11 + 99, 0
    Picture2.Top = 10
    Picture2.Left = Wi - 16 - 11 + 99
    LoadRightCorner Wi - 12 + 99, 0
    
    LoadSideBarLeft 0, 29, He - 29
    LoadSideBarRight Wi + 99 - 4, 29, He - 29, Wi + 99 - 51
    LoadBorderCap 46, He - 4
    
    Picture4.Top = 29
    Picture4.Left = 4
    Picture4.Width = Wi + 99 - 8
    Picture4.Height = He - 29 - 4
    
    Label1.Top = 2
    Label1.Left = 0
    Label1.Width = Wi
    Label1.Refresh
    Label2.Top = 3
    Label2.Left = 0 + 1
    Label2.Width = Wi
    Label2.Refresh
    
    Picture5.Width = 15
    Picture5.Height = 15
    LoadDragDot Picture5.Width - 10, Picture5.Height - 10
    LoadDragDot Picture5.Width - 15, Picture5.Height - 5
    LoadDragDot Picture5.Width - 10, Picture5.Height - 5
    LoadDragDot Picture5.Width - 5, Picture5.Height - 5
    LoadDragDot Picture5.Width - 5, Picture5.Height - 10
    LoadDragDot Picture5.Width - 5, Picture5.Height - 15
    
    Picture5.Left = Wi + 99 - 15 - 4
    Picture5.Top = He - 15 - 4
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.ScaleHeight - Y < 5 And UserControl.ScaleWidth - X < 5 Then
        LastMouseOver = "Corner"
        Screen.MousePointer = 8
        If Button = 1 Then
        ReleaseCapture
        SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
        LoadGui
        End If
    ElseIf UserControl.ScaleHeight - Y < 5 Then
        LastMouseOver = "Bottom"
        Screen.MousePointer = 7
        If Button = 1 Then
        ReleaseCapture
        SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0&
        LoadGui
        End If
    ElseIf UserControl.ScaleWidth - X < 5 Then
        LastMouseOver = "Right"
        Screen.MousePointer = 9
        If Button = 1 Then
        ReleaseCapture
        SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0&
        LoadGui
        End If
    Else
        Screen.MousePointer = 1
        LastMouseOver = "Top"
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Parent.Cls
    UserControl.Parent.ScaleMode = 3
    UserControl.Parent.BackColor = &HFF00FF
    UserControl.Parent.BorderStyle = 0
    UserControl.Width = Parent.Width
    UserControl.Height = Parent.Height
    'UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Top = 0
    'UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Left = 0
    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Align = 1
End Sub

Private Sub UserControl_Show()
    LoadGui
    Timer1.Enabled = True
End Sub

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function



VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码查看器"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   12
      Text            =   "VB精彩编程百例"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox PasswordText 
      Height          =   270
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox WndClassText 
      Height          =   270
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox hWndText 
      Height          =   270
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox PointText 
      Height          =   270
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CheckBox Check1 
      Caption         =   "总在最上面"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "密码"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "请拖动放大镜图标到显示 '***' 的密码窗口"
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "窗口文本Password："
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "鼠标所在窗口句柄："
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客户坐标值："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "获得窗口类型："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsDragging As Boolean

Private Sub SetOnTop(ByVal IsOnTop As Integer)
Dim rtn As Long
    If IsOnTop = 1 Then
        '将窗口置于最上面
        rtn = SetWindowPos(frmMain.hwnd, -1, 0, 0, 0, 0, 3)
    Else
        rtn = SetWindowPos(frmMain.hwnd, -2, 0, 0, 0, 0, 3)
    End If
End Sub

Private Sub Check1_Click()
    SetOnTop (Check1.Value)
End Sub

Private Sub Form_Load()
    Check1.Value = 1
    SetOnTop (Check1.Value)
    IsDragging = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsDragging = True Then
    Dim rtn As Long, curwnd As Long
    Dim tempstr As String
    Dim strlong As Long
    Dim point As POINTAPI
    point.x = x
    point.y = y
    '将客户坐标转化为屏幕坐标并显示在PointText文本框中
    If ClientToScreen(frmMain.hwnd, point) = 0 Then Exit Sub
    PointText.Text = Str(point.x) + "," + Str(point.y)
    '获得鼠标所在的窗口句柄并显示在hWndText文本框中
    curwnd = WindowFromPoint(point.x, point.y)
    hWndText.Text = Str(curwnd)
    '获得该窗口的类型并显示在WndClassText文本框中
    tempstr = Space(255)
    strlong = Len(tempstr)
    rtn = GetClassName(curwnd, tempstr, strlong)
    If rtn = 0 Then Exit Sub
    tempstr = Trim(tempstr)
    WndClassText.Text = tempstr
    '向该窗口发送一个WM_GETTEXT消息，以获得该窗口的文本，并显示在PasswordText文本框中
    tempstr = Space(255)
    strlong = Len(tempstr)
    rtn = SendMessage(curwnd, WM_GETTEXT, strlong, tempstr)
    tempstr = Trim(tempstr)
    PasswordText.Text = tempstr
End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsDragging = True Then
    Screen.MousePointer = vbDefault
    IsDragging = False
    '释放鼠标消息抓取
    ReleaseCapture
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsDragging = False Then
    IsDragging = True
    Screen.MouseIcon = LoadPicture(App.Path + "\pass.ico")
    Screen.MousePointer = vbCustom
    '将以后的鼠标输入消息都发送到本程序窗口
    SetCapture (frmMain.hwnd)
End If
   
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "隐藏标题栏"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function TitleBar(ByVal bState As Boolean)
    Dim lStyle As Long
    Dim tR As RECT

    GetWindowRect Me.hWnd, tR
    '得到窗体的区域保存在tr中
    lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    '得到窗体目前的风格设置
    If (bState) Then
    '如果显示标题栏
        Me.Caption = Me.Tag
        '设置Caption属性
        If Me.ControlBox Then
            lStyle = lStyle Or WS_SYSMENU
            '设置显示系统菜单
        End If
        If Me.MaxButton Then
            lStyle = lStyle Or WS_MAXIMIZEBOX
            '设置显示最大化按钮
        End If
        If Me.MinButton Then
            lStyle = lStyle Or WS_MINIMIZEBOX
            '设置显示最小化按钮
        End If
        If Me.Caption <> "" Then
            lStyle = lStyle Or WS_CAPTION
            '显示窗口的标题
        End If
    Else
    '如果隐藏标题栏
        Me.Tag = Me.Caption
        '将窗口标题保存到窗口的tag属性中
        Me.Caption = ""
        '将窗口标题设置为空字符串
        lStyle = lStyle And Not WS_SYSMENU
        '隐藏系统菜单
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        '隐藏最大化按钮
        lStyle = lStyle And Not WS_MINIMIZEBOX
        '隐藏最小化按钮
        lStyle = lStyle And Not WS_CAPTION
        '隐藏标题
    End If
    SetWindowLong Me.hWnd, GWL_STYLE, lStyle
    '设置窗体风格
    SetWindowPos Me.hWnd, 0, tR.Left, tR.Top, _
                    tR.Right - tR.Left, tR.Bottom - tR.Top, _
                    SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
    '保证窗口具有相同的大小
End Function

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        TitleBar False
    Else
        TitleBar True
    End If
End Sub


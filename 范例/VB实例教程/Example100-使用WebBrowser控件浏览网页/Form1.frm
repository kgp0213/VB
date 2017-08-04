VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "浏览器"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7500
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton StopButton 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton ForwardButton 
      Caption         =   "Forward"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton GoButton 
      Caption         =   "Go"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BackButton 
      Caption         =   "Back"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   3495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   8705
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "地址："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackButton_Click()
'返回上一个页面
    WebBrowser1.GoBack
End Sub

Private Sub Form_Load()
'程序装入后进入IE设定的起始页
    WebBrowser1.GoHome
End Sub

Private Sub Form_Resize()
'改变窗口大小后同时改变控件的大小
    WebBrowser1.Width = Form1.ScaleWidth
    WebBrowser1.Height = Form1.ScaleHeight - 900
    
End Sub

Private Sub ForwardButton_Click()
'进入下一个页面
    WebBrowser1.GoForward
End Sub

Private Sub GoButton_Click()
'浏览输入的页面
    WebBrowser1.Navigate (Text1.Text)
End Sub

Private Sub StopButton_Click()
'停止浏览
    WebBrowser1.Stop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'浏览输入的页面
    If KeyAscii = 13 Then
        WebBrowser1.Navigate (Text1.Text)
    End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'将当前显示的页面的URL地址显示在Text1上
    Text1.Text = URL
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
'窗口的标题栏中显示当前页面装入情况
    Me.Caption = Text
End Sub

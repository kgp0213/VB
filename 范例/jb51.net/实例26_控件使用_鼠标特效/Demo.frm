VERSION 5.00
Object = "{5754A831-F79C-11D3-A259-0080C8588E1D}#3.0#0"; "Mouse.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mouse"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Mouse_Events.Mouse Mouse1 
      Left            =   30
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "右键点击"
      Height          =   375
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1530
      Width           =   1875
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "左键点击"
      Height          =   375
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1110
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "鼠标位置初始化"
      Height          =   375
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   690
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "获取鼠标的坐标"
      Height          =   375
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   270
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Not yet clicked"
      Height          =   2025
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   1605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X&, Y&

Private Sub Command1_Click()
  Mouse1.MouseCoordinates X, Y
  Label1.Caption = "X:" & X & " Y:" & Y
End Sub

Private Sub Command2_Click()
  X = 0: Y = 0
  Mouse1.MouseMove X, Y
End Sub

Private Sub Command3_Click()
  X = (Form1.Left + Label1.Left + 1500) / Screen.TwipsPerPixelX: Y = (Form1.Top + Label1.Top + 1000) / Screen.TwipsPerPixelY
  Mouse1.MouseMove X, Y
  Mouse1.LeftClick
End Sub

Private Sub Command4_Click()
  X = (Form1.Left + Label1.Left + 1500) / Screen.TwipsPerPixelX: Y = (Form1.Top + Label1.Top + 1000) / Screen.TwipsPerPixelY
  Mouse1.MouseMove X, Y
  Mouse1.RightClick
End Sub


Private Sub Form_Load()

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then Label1.Caption = "Left Button Click"
  If Button = 2 Then Label1.Caption = "Right Button Click"
  If Button = 4 Then Label1.Caption = "Middle Button Click"
End Sub

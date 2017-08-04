VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Link"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3090
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label2 
      Caption         =   "mailto:yinlimin@sina.com"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "http://www.hlgnet.com"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      MouseIcon       =   "Form1.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
                Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) _
                As Long

Private Sub Label1_Click()
    Call ShellExecute(Me.hwnd, "Open", Me.Label1.Caption, "", App.Path, 1)
End Sub

Private Sub Label2_Click()
    Call ShellExecute(Me.hwnd, "Open", Me.Label2.Caption, "", App.Path, 1)
End Sub

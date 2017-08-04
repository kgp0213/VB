VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "批量文本文件转为HTML文件"
   ClientHeight    =   3225
   ClientLeft      =   3990
   ClientTop       =   3195
   ClientWidth     =   4965
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "退出(&X)"
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   3300
      TabIndex        =   11
      Text            =   "VB创作效果百例"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   3300
      TabIndex        =   6
      Text            =   "#AAD3F2"
      Top             =   2355
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   3300
      TabIndex        =   5
      Text            =   "#000000"
      Top             =   1655
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   3300
      TabIndex        =   4
      Text            =   "2"
      Top             =   955
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   240
      MultiSelect     =   2  'Extended
      Pattern         =   "*.txt"
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   570
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转换(&L)"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "共有0个文件需要转换"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "窗口标题"
      Height          =   180
      Index           =   4
      Left            =   2400
      TabIndex        =   10
      Top             =   300
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "背景颜色"
      Height          =   180
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "字体颜色"
      Height          =   180
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "字体大小"
      Height          =   180
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Top             =   1005
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
       Case 0
mypath = Dir1.Path
te = Right(mypath, 1)
If te <> "\" Then mypath = mypath + "\"
i = File1.ListCount - 1
For j = 0 To i
abcd = (File1.List(j))
file = Left(abcd, Len(abcd) - 4)

filet = mypath + abcd
fileh = mypath + file + ".html"

Open fileh For Output As #1
Close #1

Open filet For Input As #2

Me.Caption = "文件正在转换，请等待。。。。"
Command1(0).Enabled = False
Open fileh For Append As #1
    a1 = "<Title>" & Text1(0).Text & "</Title> <META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=gb2312'> <body bgcolor=" & Text1(3).Text & ">  </body> <font color=" & Text1(2).Text & " Size=" & Text1(1).Text & " > "
    Print #1, a1
    Close #1
Do While Not EOF(2) '
    Line Input #2, textline
    Open fileh For Append As #1
    a2 = textline + "<br>"
    Print #1, a2
    Close #1
Loop
Open fileh For Append As #1
    a3 = "</font>"
    Print #1, a3
    Close #1

Close #2
Next
Me.Caption = "批量文本文件转为HTML文件"
Command1(0).Enabled = True


Case 3
     End
End Select
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
Label1(3).Caption = "共有" & File1.ListCount & "个文件需要转换"
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
Label1(3).Caption = "共有" & File1.ListCount & "个文件需要转换"
End Sub


Private Sub Form_Load()
Label1(3).Caption = "共有" & File1.ListCount & "个文件需要转换"
End Sub


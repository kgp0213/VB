VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ListBox"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   495
      Index           =   4
      Left            =   1920
      TabIndex        =   3
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   465
      Index           =   3
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "已选项"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "可选项"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    Screen.MousePointer = 11
    Select Case Index
        Case 1  '将选择的列表项从List1移到List2
            If List1.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            '如果List1中没有列表项则退出
            If List1.ListIndex = -1 Then
                List1.SetFocus
                List1.Selected(0) = True
            End If
            '如果List1中没有选中的列表项则选择第一个列表项
            DoEvents
            List2.AddItem List1.Text
            List1.RemoveItem List1.ListIndex
            '将选择的列表项从List1移到List2
        Case 2    '将List1的所有列表项移到List2中
            If List1.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            '如果List1中没有列表项则退出
            If List1.ListIndex = -1 Then
                List1.SetFocus
                List1.Selected(0) = True
            End If
            '如果List1中没有选中的列表项则选择第一个列表项
            DoEvents
            For i = (List1.ListCount - 1) To 0 Step -1
                List2.AddItem List1.List(i)
                DoEvents
            Next i
            '将List1的所有列表项添加到List2中
            List1.Clear
            '删除List1中的所有列表项

        Case 3  '将选择的列表项从List2移到List1
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            '如果List2中没有列表项则退出
            If List2.ListIndex = -1 Then
                List2.SetFocus
                List2.Selected(0) = True
            End If
            '如果List2中没有选中的列表项则选择第一个列表项

            List1.AddItem List2.Text
            List2.RemoveItem List2.ListIndex
            '将选择的列表项从List2移到List1

        Case 4 '将List2的所有列表项移到List1中
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            If List2.ListCount = 0 Then Exit Sub
            If List2.ListIndex = -1 Then
                List2.SetFocus
                List2.Selected(0) = True
            End If
            For i = (List2.ListCount - 1) To 0 Step -1
                List1.AddItem List2.List(i)
                DoEvents
            Next i
            List2.Clear
    End Select
    Screen.MousePointer = 0
End Sub
 
Private Sub Form_Load()
    List1.AddItem "姓名", 0
    List1.AddItem "性别", 1
    List1.AddItem "年龄", 2
    List1.AddItem "籍贯", 3
    List1.AddItem "文化程度", 4
End Sub

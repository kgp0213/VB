VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "输入检查"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4395
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "临时标签，用来获得焦点，使组合框能够失去焦点，并在失去焦点时进行有效性检查"
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Validate(Cancel As Boolean)
   Dim temp As Boolean
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = Combo1.Text Then
            temp = True
        Else: temp = False
        End If
     Next
        
    If temp = False Then
        MsgBox "输入的数据不在列表中，请重新输入！", vbExclamation + vbOKOnly, "数据错误"
        Combo1.SetFocus
        Combo1.SelStart = 0
        Combo1.SelLength = Len(Combo1.Text)
    End If
End Sub

Private Sub Form_Load()
'添加列表
    Combo1.AddItem "姓名", 0
    Combo1.AddItem "性别", 1
    Combo1.AddItem "年级", 2
    Combo1.AddItem "籍贯", 3
    Combo1.AddItem "政治面貌", 4
    Combo1.AddItem "出生年月", 5
End Sub

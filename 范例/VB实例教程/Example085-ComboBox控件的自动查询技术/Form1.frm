VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "自动查询"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4455
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   245
      TabIndex        =   1
      Top             =   645
      Width           =   3970
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
    Dim sString As String
    Dim start As Integer
    start = Combo1.SelStart
    sString = Left(Combo1.Text, start)
    For i = 0 To Combo1.ListCount - 1 Step 1
        Dim sitem As String
        sitem = Combo1.List(i)
        sitem = Left(sitem, start)
        If sitem = sString Then
            List1.ListIndex = i
            List1.Visible = True
            Exit For
        End If
    Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.ListIndex = List1.ListIndex
        List1.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "abc"
    Combo1.AddItem "acb"
    Combo1.AddItem "edf"
    Combo1.AddItem "ffff"
    '向Combo1添加列表项
    
    Dim i As Integer
    For i = 0 To Combo1.ListCount - 1 Step 1
        List1.AddItem Combo1.List(i), i
    Next
    List1.Visible = False
    '向List1添加与Combo1相同的列表项
End Sub

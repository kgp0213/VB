VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Form1 
   Caption         =   "MaskEdBox"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "取消掩码"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
      _Version        =   393216
      BackColor       =   16777215
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'取消或者设置掩码
    If Check1.Value = 1 Then
        MaskEdBox1.Mask = ""
    Else
        MaskEdBox1.Mask = "##/##/##"
    End If
End Sub

Private Sub Form_Load()
    With MaskEdBox1
        .Mask = "##/##/##"
        .Text = "12/01/04"
        .Font.Name = "Arial"
        .Font.Bold = True
        .Font.Size = 10
        .Height = 500
        .Width = 1500
    End With
    '该控件显示日期信息
 End Sub
Private Sub MaskEdBox1_GotFocus()
'当输入数据时背景为紫色文字颜色为黄色
    MaskEdBox1.BackColor = &H800080
    MaskEdBox1.ForeColor = &HFFFF&
End Sub

Private Sub MaskEdBox1_LostFocus()
'输入完毕恢复默认的颜色
    MaskEdBox1.ForeColor = &H0
    MaskEdBox1.BackColor = &HFFFFFF
End Sub

Private Sub MaskEdBox1_ValidationError( _
            InvalidText As String, StartPosition As Integer)
'错误处理
    MsgBox "数据错误"
    MaskEdBox1.Text = "12/01/04"
    MaskEdBox1.SelStart = 0
    MaskEdBox1.SelLength = Len(MaskEdBox1.FormattedText)
End Sub

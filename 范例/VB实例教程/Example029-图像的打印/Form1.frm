VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "打印"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印图"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打印窗体"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Err_Handle
    '发生错误则转移到Err_handle
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter '显示打印对话框
    Form1.PrintForm           '打印窗体
    Exit Sub                  '退出
Err_Handle:                   '如果发生错误执行下面程序段
    MsgBox Err.Description    '显示错误信息
    Exit Sub                  '然后退出该过程
End Sub

Private Sub Command2_Click()
    Printer.PaintPicture Picture1.Picture, 0, 0, _
                         Picture1.Width, Picture1.Width
    Printer.EndDoc
End Sub

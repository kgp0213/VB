VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form2 
   Caption         =   "段落"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3780
   LinkTopic       =   "Form2"
   ScaleHeight     =   2835
   ScaleWidth      =   3780
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin MSComCtl2.UpDown UpDownR 
      Height          =   375
      Left            =   3016
      TabIndex        =   8
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      BuddyControl    =   "TextR"
      BuddyDispid     =   196616
      OrigLeft        =   3240
      OrigTop         =   1680
      OrigRight       =   3480
      OrigBottom      =   2055
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TextR 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin MSComCtl2.UpDown UpDownL 
      Height          =   375
      Left            =   3016
      TabIndex        =   6
      Top             =   960
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      BuddyControl    =   "TextL"
      BuddyDispid     =   196614
      OrigLeft        =   3240
      OrigTop         =   960
      OrigRight       =   3480
      OrigBottom      =   1335
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TextL 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Text            =   "0"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TextF 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   1095
   End
   Begin MSComCtl2.UpDown UpDownF 
      Height          =   375
      Left            =   3016
      TabIndex        =   3
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      BuddyControl    =   "TextF"
      BuddyDispid     =   196613
      OrigLeft        =   3240
      OrigTop         =   360
      OrigRight       =   3480
      OrigBottom      =   735
      Max             =   1000
      Min             =   -1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "右缩进量："
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "左缩进量："
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "首行悬挂缩进："
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Form1.RichTextBox1.SelHangingIndent = Val(TextF.Text)
    Form1.RichTextBox1.SelIndent = Val(TextL.Text)
    Form1.RichTextBox1.SelRightIndent = Val(TextR.Text)
    Unload Me
End Sub

Private Sub Form_Load()
    TextF.Text = Str(Form1.RichTextBox1.SelHangingIndent)
    TextL.Text = Str(Form1.RichTextBox1.SelIndent)
    TextR.Text = Str(Form1.RichTextBox1.SelRightIndent)
End Sub

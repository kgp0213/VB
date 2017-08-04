VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   11295
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "试验2"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "试验1"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin 工程1.muchMenu muchMenu1 
      Height          =   1890
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3334
      StartColor      =   12632256
      CeaseColor      =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RepeatCount     =   5
      RepeatCurrent   =   4
      ItemSum1        =   14
      ItemSum2        =   8
      ItemSum3        =   9
      ItemSum4        =   7
      ItemSum5        =   2
      sCaption101     =   "新建          Ctrl+N"
      sCaption102     =   "打开          Ctrl+O"
      sCaption103     =   "保存          Ctrl+S"
      sCaption104     =   "另存为"
      sCaption105     =   "---------------------"
      sCaption106     =   "文本重排          F1"
      sCaption107     =   "文本合并          F2"
      sCaption108     =   "文本分割          F3"
      sCaption109     =   "文本朗读          F4"
      sCaption1010    =   "文本寄发          F5"
      sCaption1011    =   "---------------------"
      sCaption1012    =   "打印          Ctrl+P"
      sCaption1013    =   "加密"
      sCaption1014    =   "解密"
      sCaption201     =   "全选          Ctrl+A"
      sCaption202     =   "字块          Ctrl+B"
      sCaption203     =   "--------------------"
      sCaption204     =   "撤消          Ctrl+Z"
      sCaption205     =   "剪切          Ctrl+X"
      sCaption206     =   "复制          Ctrl+C"
      sCaption207     =   "粘贴          Ctrl+v"
      sCaption208     =   "删除             Del"
      sCaption301     =   "字体"
      sCaption302     =   "--------------------"
      sCaption303     =   "用户设置      Ctrl+I"
      sCaption304     =   "定时提醒      Ctrl+Y"
      sCaption305     =   "系统设置"
      sCaption306     =   "系统信息"
      sCaption307     =   "--------------------"
      sCaption308     =   "普通滚屏      Ctrl+G"
      sCaption309     =   "字数统计      Ctrl+L"
      sCaption401     =   "国繁>台繁    Ctrl+F1"
      sCaption402     =   "台繁>国繁    Ctrl+F2"
      sCaption403     =   "国繁>国简    Ctrl+F3"
      sCaption404     =   "国简>国繁    Ctrl+F4"
      sCaption405     =   "--------------------"
      sCaption406     =   "全角>半角    Ctrl+F7"
      sCaption407     =   "半角>全角    Ctrl+F8"
      sCaption501     =   "帮助主题"
      sCaption502     =   "关于...."
   End
   Begin VB.CommandButton Command 
      Caption         =   "帮助"
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "转换"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin 工程1.MyMenu MyMenu1 
      Height          =   3780
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6668
      StartColor      =   16777215
      CeaseColor      =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemSum         =   14
      sCaption1       =   "新建          Ctrl+N"
      sCaption2       =   "打开          Ctrl+O"
      sCaption3       =   "保存          Ctrl+S"
      sCaption4       =   "另存为"
      sCaption5       =   "---------------------"
      sCaption6       =   "文本重排          F1"
      sCaption7       =   "文本合并          F2"
      sCaption8       =   "文本分割          F3"
      sCaption9       =   "文本朗读          F4"
      sCaption10      =   "文本寄发          F5"
      sCaption11      =   "---------------------"
      sCaption12      =   "打印          Ctrl+P"
      sCaption13      =   "加密"
      sCaption14      =   "解密"
   End
   Begin 工程1.MenuControl MenuControl1 
      Height          =   1605
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2831
      ItemSum         =   5
      CeaseColor      =   8421631
   End
   Begin VB.CommandButton Command 
      Caption         =   "选项"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "编辑"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "文件"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Download by http://down.liehuo.net
Private Sub Form_Click()
MyMenu1.Visible = False
muchMenu1.Visible = False
MenuControl1.Visible = False
End Sub

Private Sub Command1_Click()
MenuControl1.Move Command1.left + Command1.Width, Command1.top
MenuControl1.Visible = True
End Sub

Private Sub Command2_Click()
MyMenu1.Move Command2.left + Command2.Width, Command2.top
MyMenu1.Visible = True
End Sub

Private Sub Command_Click(Index As Integer)
muchMenu1.Move Command(Index).left + Command(Index).Width, Command(Index).top
muchMenu1.RepeatCurrent = Index
muchMenu1.Visible = True
End Sub

Private Sub muchMenu1_Click(SelectedItem As Integer)
muchMenu1.Visible = False
MsgBox "第" & SelectedItem & "个菜单项被选中"
End Sub

Private Sub MyMenu1_Click(SelectedItem As Integer)
MyMenu1.Visible = False
MsgBox "第" & SelectedItem & "个菜单项被选中"
End Sub

Private Sub menuControl1_Click(SelectedItem As Integer)
MenuControl1.Visible = False
MsgBox "第" & SelectedItem & "个菜单项被选中"
End Sub



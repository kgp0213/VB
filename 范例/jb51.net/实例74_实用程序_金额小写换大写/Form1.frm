VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "转换(&C)"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "大写（￥）"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "小写(数字)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = rmb(Text1.Text)
End Sub

Public Function rmb(s As Currency) As String
    s1$ = LTrim(Str$(Abs(s)))
    L% = Len(s1)
    Select Case L - InStrRev(s1, ".")
    '双引号内是小数点
       Case L
         s2$ = s1 + ".00"
       Case 1
         s2$ = s1 + "0"
       Case 2
         s2$ = s1
    End Select
    L = Len(s2)
    DX$ = ""
    C1$ = "零壹贰叁肆伍陆柒捌玖"
    C2$ = "分角 元拾佰仟万拾佰仟亿拾佰"
    '角和元之间留一个空格
     Do While L >= 1
     x$ = Mid(s2, Len(s2) - L + 1, 1)
     DX = DX + IIf(x <> ".", Mid(C1, Val(x) + 1, 1) + Trim(Mid(C2, (L - 1) + 1, 1)), "")
     L = L - 1
     Loop
     rmb = DX + "整"
End Function
  

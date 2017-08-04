VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Begin VB.Form Form1 
   Caption         =   "切割"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "切   割"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox TextH 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "20"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TextW 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "20"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox TextY 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TextX 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1215
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   360
      Top             =   3000
      _ExtentX        =   9525
      _ExtentY        =   1905
      _Version        =   393216
      Rows            =   3
      Cols            =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "左上角Y坐标："
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "左上角X坐标："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "高："
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "宽："
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Image1.Width = TextW.Text
    Image1.Height = TextH.Text
    '根据剪切区域的大小设置Image控件大小
    PictureClip1.ClipX = TextX.Text
    PictureClip1.ClipY = TextY.Text
    PictureClip1.ClipWidth = TextW.Text
    PictureClip1.ClipHeight = TextH.Text
    Image1.Picture = PictureClip1.Clip
End Sub

Private Sub Form_Load()
    PictureClip1.Picture = LoadPicture(App.Path + "\1.bmp")
    '向PictureClip1添加图片
    TextX.Text = 0
    TextY.Text = 0
    TextW.Text = PictureClip1.Width
    TextH.Text = PictureClip1.Height
    '根据图片的大小设置各TextBox控件的Text属性值

End Sub

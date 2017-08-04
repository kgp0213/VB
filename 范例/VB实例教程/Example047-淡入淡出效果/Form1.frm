VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "淡入淡出效果"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "淡出效果"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "淡入效果"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   3240
      ScaleHeight     =   3675
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AlphaBlend Lib "Msimg32.dll" ( _
        ByVal hdcDest As Long, _
        ByVal nXOriginDest As Long, _
        ByVal nYOriginDest As Long, _
        ByVal nWidthDest As Long, _
        ByVal nHeightDest As Long, _
        ByVal hdcSrc As Long, _
        ByVal nXOriginSrc As Long, _
        ByVal nYOriginSrc As Long, _
        ByVal nWidthSrc As Long, _
        ByVal nHeightSrc As Long, _
        ByVal BLENDFUNCTION As Long) As Boolean
Const AC_SRC_OVER = &H0
Const AC_SRC_ALPHA = &H1
Private Type BLENDFUNCTION
        BlendOP As Byte
        BlendFlags As Byte
        SourceConstantAlpha As Byte
        AlphaFormat As Byte
End Type
Private Declare Sub Sleep Lib "kernel32" _
( _
    ByVal dwMilliseconds As Long _
    )
'Sleep为延时函数以毫秒为单位指定等待的时间
Dim sBlendFunction As BLENDFUNCTION
Private Sub Form_Load()
    sBlendFunction.BlendOP = AC_SRC_OVER
    sBlendFunction.BlendFlags = 0
    sBlendFunction.AlphaFormat = 0
    Form1.ScaleMode = 3
    Picture1.ScaleMode = 3
    Picture2.ScaleMode = 3
    '设置Form、Picture1和Picture2的标志单位为像素
    Picture1.AutoRedraw = False
    Picture2.AutoRedraw = False
    Picture1.Picture = LoadPicture(App.Path + "\1.bmp")
End Sub
Private Sub Command1_Click()
'淡入效果
    Dim LnBlendPtr As Long
    Picture2.Cls
    For i = 0 To 50
        sBlendFunction.SourceConstantAlpha = i * 5
        CopyMemory LnBlendPtr, sBlendFunction, 4
        AlphaBlend Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, _
                   Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, _
                   LnBlendPtr
        Sleep (50)
        DoEvents
    Next
End Sub

Private Sub Command2_Click()
'淡出效果
    Dim LnBlendPtr As Long
    Picture2.Cls
    For i = 0 To 10
        sBlendFunction.SourceConstantAlpha = 250 - i * 25
        CopyMemory LnBlendPtr, sBlendFunction, 4
        AlphaBlend Picture2.hDC, 0, 0, Picture2.Width, _
                   Picture2.Height, Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, _
                   LnBlendPtr
        Sleep (50)
        Picture2.Refresh
        DoEvents
    Next
End Sub


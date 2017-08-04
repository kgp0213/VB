VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "×¥ÆÁÈí¼þ"
   ClientHeight    =   3270
   ClientLeft      =   2010
   ClientTop       =   1545
   ClientWidth     =   3360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Print Screen"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
Dim wScreen As Long
Dim hScreen As Long
Dim w As Long
Dim h As Long
Picture1.Cls

wScreen = Screen.Width \ Screen.TwipsPerPixelX
hScreen = Screen.Height \ Screen.TwipsPerPixelY

Picture1.ScaleMode = vbPixels
w = Picture1.ScaleWidth
h = Picture1.ScaleHeight


hdcScreen = GetDC(0)

r = StretchBlt(Picture1.hdc, 0, 0, w, h, hdcScreen, 0, 0, wScreen, hScreen, vbSrcCopy)

End Sub


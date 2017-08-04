VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Getpixel sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   1830
   ClientLeft      =   1665
   ClientTop       =   1545
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1830
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4680
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   3780
      ScaleHeight     =   1635
      ScaleWidth      =   2055
      TabIndex        =   10
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   9
      Top             =   1500
      Width           =   1755
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   8
      Top             =   1140
      Width           =   1755
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   420
      Width           =   1755
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RGB Color:"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Client Pixel Pos:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1140
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Absolute Pixel Pos:"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1710
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Window hDC:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Window Handle:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Sub Timer1_Timer()
    Static lX As Long, lY As Long
    On Local Error Resume Next
    Dim P As POINTAPI, h As Long, hD As Long, r As Long
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    lblData(0).Caption = lX & "," & lY
    h = WindowFromPoint(lX, lY)
    lblData(1).Caption = h
    hD = GetDC(h)
    lblData(2).Caption = hD
    ScreenToClient h, P
    lblData(3).Caption = P.x & "," & P.y
    r = GetPixel(hD, P.x, P.y)
    If r = -1 Then
        BitBlt Picture1.hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
        r = Picture1.Point(0, 0)
    Else
        Picture1.PSet (0, 0), r
    End If
    lblData(4).Caption = Hex$(r)
    Picture1.BackColor = r
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "精彩万花筒"
   ClientHeight    =   3720
   ClientLeft      =   3975
   ClientTop       =   2820
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "BMP Files(*.bmp)|*.bmp"
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "退出"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "保存"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdpaint 
      Caption         =   "画图"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdduo 
      Caption         =   "多色"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmddan 
      Caption         =   "单色"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picpaint 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public duo As Boolean
Public icolor As Long
Private Sub cmddan_Click()
cdl1.ShowColor
icolor = cdl1.Color
duo = False
End Sub

Private Sub cmdduo_Click()
duo = True
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdpaint_Click()
Const pi = 3.1415926
Dim temp As Double
Dim per As Integer
picpaint.Cls
a = 95
ifunction = Int(4 * Rnd)
cx = 120: cy = 110
d = 2 * Rnd
per = Int(Rnd * 5) + 5
For bt = 0 To pi * (Rnd + 1) Step pi / per
  bt1 = Cos(bt): bt2 = Sin(bt)
  For g = 1 To 2
    For l = -1 To 1 Step 2
      For z = -90 To 90 Step 5
          x = z: al = (z + 90) * 2 * pi / 180
        Select Case ifuncion
         Case 0
          y = l * a * Sin(al) * Cos(d * al)
         Case 1
           y = l * a * Sin(al) * Sin(d * al)
         Case 2
           y = l * a * Cos(al) * Cos(d * al)
         Case 3
             y = l * a * Cos(al) * Sin(d * al)
        End Select
            If g = 2 Then
              temp = x: x = y: y = temp
            End If
          X1 = x * bt1 - y * bt2
          Y1 = x * bt2 + y * bt1
          X2 = cx - X1: Y2 = cy + Y1
            If z = -90 Then
            bx = X2: By = Y2
            picpaint.PSet (bx, By), QBColor(13)
             ElseIf duo Then
              Randomize
              rr = Int(225 * Rnd): gg = Int(225 * Rnd): bb = Int(225 * Rnd)
               picpaint.Line -(X2, Y2), RGB(rr, gg, bb)
            Else
              picpaint.Line -(X2, Y2), icolor
           End If
       Next z: Next l: Next g: Next bt
            
End Sub

Private Sub cmdsave_Click()
Dim filename As String
cdl1.DialogTitle = "保存"
cdl1.ShowSave


filename = cdl1.filename
If filename <> "" Then
SavePicture picpaint.Image, filename
End If

End Sub

Private Sub Form_Load()
cdl1.Flags = cdlOFNOverwritePrompt + cdlOFNFileMustExist + cdlOFNCreatePrompt + cdlOFNHideReadOnly
End Sub

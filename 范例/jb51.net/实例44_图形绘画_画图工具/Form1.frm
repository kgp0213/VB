VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open Face Off"
   ClientHeight    =   6585
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox masktemp 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   7440
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   14
      Top             =   6600
      Width           =   3975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Image Mask"
      Height          =   6495
      Left            =   7800
      TabIndex        =   7
      Top             =   0
      Width           =   4935
      Begin VB.PictureBox mask 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ClipControls    =   0   'False
         DrawWidth       =   50
         ForeColor       =   &H00000000&
         Height          =   6135
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   6075
         ScaleWidth      =   4635
         TabIndex        =   8
         Top             =   240
         Width           =   4695
         Begin VB.Shape dc 
            Height          =   750
            Left            =   1080
            Shape           =   3  'Circle
            Top             =   2280
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox tempbuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   6600
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   -240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".jpg"
      FileName        =   "*.gif;*.jpg;*.bmp"
      Filter          =   "Image Files"
      InitDir         =   "."
   End
   Begin VB.Frame Frame2 
      Caption         =   "Main"
      Height          =   6495
      Left            =   5160
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      Begin VB.PictureBox colorshow 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   21
         Top             =   2160
         Width           =   495
      End
      Begin ComctlLib.Slider b 
         Height          =   1455
         Left            =   330
         TabIndex        =   19
         Top             =   2880
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2566
         _Version        =   327682
         Orientation     =   1
         Max             =   255
      End
      Begin ComctlLib.Slider g 
         Height          =   1455
         Left            =   225
         TabIndex        =   18
         Top             =   2880
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2566
         _Version        =   327682
         Orientation     =   1
         Max             =   255
      End
      Begin ComctlLib.Slider r 
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2566
         _Version        =   327682
         Orientation     =   1
         Max             =   255
      End
      Begin VB.Frame Frame7 
         Caption         =   "<--  <--  <--  <--  <--  <--  <--  <--"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   4320
         Width           =   2295
         Begin VB.CommandButton Command3 
            Caption         =   "<<Generate"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Undo"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "-->  -->  -->  -->  -->  -->  -->  -->"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   2295
         Begin VB.CommandButton Command1 
            Caption         =   "Generate Onion Skin>>"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Drawsize - 50"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2295
         Begin ComctlLib.Slider bsize 
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   450
            _Version        =   327682
            LargeChange     =   10
            Min             =   1
            Max             =   200
            SelStart        =   50
            Value           =   50
         End
      End
      Begin VB.CommandButton undobutton 
         Caption         =   "Undo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Effects"
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   4920
         Width           =   2295
         Begin VB.ListBox effects 
            Height          =   1035
            ItemData        =   "Form1.frx":0000
            Left            =   120
            List            =   "Form1.frx":0010
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Label Label1 
         Caption         =   "R G B"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Image"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   6135
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   6075
         ScaleWidth      =   4635
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         Begin VB.Line statusline 
            Visible         =   0   'False
            X1              =   0
            X2              =   4680
            Y1              =   0
            Y2              =   0
         End
      End
   End
   Begin VB.Menu mfile 
      Caption         =   "File"
      Begin VB.Menu sopenimage 
         Caption         =   "Open Image"
      End
      Begin VB.Menu squit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu moptions 
      Caption         =   "Options"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1 As Integer
Dim y1 As Integer
Dim t As Integer
Dim s As Integer
Dim busy As Boolean
Dim curcol As Variant
Dim returner As Variant
Dim color As Long

Dim md As Boolean

Dim red As Integer
Dim green As Integer
Dim blue As Integer



Private Sub coloraverage()


For y1 = 2 To pic.Height / 15
    For x1 = 2 To pic.Width / 15
        If GetPixel(mask.hdc, x1, y1) = RGB(0, 0, 0) Then
        
            csort (GetPixel(pic.hdc, x1, y1))
            returner = SetPixel(pic.hdc, x1, y1, RGB((red * 2 + r) / 3, (green * 2 + g) / 3, (blue * 2 + b) / 3))
        
        End If
    Next x1
    clearstack
Next y1

finish



End Sub


Private Sub csort(ByVal color As Long)

red = color Mod &H100
green = (color \ &H100) Mod &H100
blue = (color \ &H10000) Mod &H100

End Sub

Private Sub clearstack()

statusline.y1 = y1 * 15
statusline.Y2 = y1 * 15
DoEvents

End Sub



Private Sub finish()

Form1.Enabled = True
busy = False
statusline.Visible = False
undobutton.Enabled = True

End Sub


Private Sub onion()
statusline.Visible = True
busy = True

For y1 = 1 To pic.Height / 15
    For x1 = 1 To pic.Width / 15
        csort (GetPixel(pic.hdc, x1, y1))
        returner = SetPixel(mask.hdc, x1, y1, RGB(red + (250 - red) / 2, green + (250 - green) / 2, blue + (250 - blue) / 2))
    Next x1
    clearstack
Next y1

busy = False
statusline.Visible = False
mask.Refresh

End Sub

Private Sub skintrace()

For y1 = 1 To pic.Height / 15
    For x1 = 1 To pic.Width / 15
        If GetPixel(mask.hdc, x1, y1) = RGB(0, 0, 0) Then
                csort GetPixel(pic.hdc, x1, y1)
                If red > green And green > blue Then
                    If green + 30 > red And blue + 60 > green Then
                        If red + blue + green > 150 Then
                            returner = SetPixel(pic.hdc, x1, y1, RGB(100, 0, 0))
                        End If
                    End If
                End If
                
        End If
    Next x1
    clearstack
Next y1

finish
End Sub

Sub subnonface()

For y1 = 1 To Form1.Height / 15
    For x1 = 1 To Form1.Width / 15
        If GetPixel(mask.hdc, x1, y1) = RGB(0, 0, 0) Then
        
            curcol = GetPixel(pic.hdc, x1, y1) - RGB(10, 15, 35)
            If curcol > RGB(10, 10, 10) Then returner = SetPixel(pic.hdc, x1, y1, curcol)
        
        End If
    Next x1
    clearstack
Next y1

finish

End Sub

Private Sub pointize()

For y1 = 2 To pic.Height / 15 Step 2
    For x1 = 2 To pic.Width / 15 Step 2
        If GetPixel(mask.hdc, x1, y1) = RGB(0, 0, 0) Then
        
            curcol = GetPixel(pic.hdc, x1, y1)
            returner = SetPixel(pic.hdc, x1 - 1, y1, curcol)
            returner = SetPixel(pic.hdc, x1 + 1, y1, curcol)
            returner = SetPixel(pic.hdc, x1, y1 + 1, curcol)
            returner = SetPixel(pic.hdc, x1, y1 - 1, curcol)
        
        End If
        Next x1
        clearstack
Next y1

finish

End Sub












Private Sub b_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
colorshow.BackColor = RGB(r, g, b)
End Sub


Private Sub bsize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame4.Caption = "Drawsize -" + Str$(bsize.Value)
mask.DrawWidth = bsize
dc.Width = bsize * 15
dc.Height = bsize * 15
End Sub




Private Sub Command1_Click()

If Not busy Then onion

End Sub



Private Sub Command2_Click()

mask = masktemp.Picture

End Sub

Private Sub Command3_Click()

Form1.Enabled = False

statusline.Visible = True

tempbuffer = pic.Image
busy = True

Select Case effects.ListIndex
    Case 0
        pointize
    Case 1
        subnonface
    Case 2
        skintrace
    Case 3
        coloraverage
    Case Else
        finish
End Select


End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
colorshow.BackColor = RGB(r, g, b)
End Sub


Private Sub mask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
masktemp = mask.Image
md = True
End Sub

Private Sub mask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not busy Then



dc.Left = X - dc.Width / 2
dc.Top = Y - dc.Height / 2




End If



If md Then mask.PSet (X, Y)


End Sub


Private Sub mask_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
md = False
End Sub

Private Sub r_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
colorshow.BackColor = RGB(r, g, b)
End Sub


Private Sub sopenimage_Click()

cmd1.ShowOpen
On Error GoTo err
tempbuffer.Picture = LoadPicture(cmd1.FileName)
pic.PaintPicture tempbuffer.Picture, 0, 0, pic.Width, pic.Height, 0, 0, tempbuffer.Width, tempbuffer.Height, vbSrcCopy
cmd1.FileName = "*.gif;*.jpg;*.bmp"
Exit Sub

err:
MsgBox "Error Loading Image File", vbInformation, "Face off"

Exit Sub

End Sub

Private Sub squit_Click()
End
End Sub




Private Sub undobutton_Click()
pic.Picture = tempbuffer.Picture
End Sub



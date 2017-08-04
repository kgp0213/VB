VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   6255
      Left            =   2880
      ScaleHeight     =   6195
      ScaleWidth      =   8475
      TabIndex        =   4
      Top             =   600
      Width           =   8535
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   240
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   4440
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2400
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "÷ªœ‘ æjpgÕº∆¨"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ºÚµ•µƒÕº∆¨‰Ø¿¿"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
Form1.File1.Path = Form1.Dir1.Path
End Sub

Private Sub Drive1_Change()
    Form1.Dir1.Path = Form1.Drive1.Drive
End Sub

Private Sub File1_Click()
    Form1.Picture1.Picture = LoadPicture(Form1.File1.Path + "\" + Form1.File1.FileName)
End Sub


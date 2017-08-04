VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   1545
   ClientTop       =   1545
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   3180
   Begin VB.TextBox txtalldrives 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdgetdrives 
      Caption         =   "&Get Drives and Type"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdgetdrives_Click()

Dim DriveNum As Integer
Dim DriveSpecifics As String
Dim TempDrive As String


txtalldrives.Text = ""

' The letters a - z
'97  --  122
For DriveNum = 97 To 122 Step 1

TempDrive = GetDriveType(Chr(DriveNum) & ":\")
  
  If TempDrive = "1" Then
    GoTo noDrive
  ElseIf TempDrive = "2" Then
    DriveSpecifics = Chr(DriveNum) & ":\" & "  is a Removeable Drive"
  ElseIf TempDrive = "3" Then
    DriveSpecifics = Chr(DriveNum) & ":\" & "  is a Local Drive"
  ElseIf TempDrive = "4" Then
    DriveSpecifics = Chr(DriveNum) & ":\" & "  is a Network Drive"
  ElseIf TempDrive = "5" Then
    DriveSpecifics = Chr(DriveNum) & ":\" & "  is a CD-ROM Drive"
  ElseIf TempDrive = "6" Then
    DriveSpecifics = Chr(DriveNum) & ":\" & "  is a RamDrive"
  Else
    GoTo noDrive
  End If
  
  ''
  ''  Display the info
  txtalldrives.Text = txtalldrives.Text & DriveSpecifics & vbCrLf
  
noDrive:
Next DriveNum
  

End Sub

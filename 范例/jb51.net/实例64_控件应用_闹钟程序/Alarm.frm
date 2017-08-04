VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Playwhat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "daN's Alarmclock"
   ClientHeight    =   3255
   ClientLeft      =   5010
   ClientTop       =   4245
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Alarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "&Open"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraSetAlarm 
      BackColor       =   &H00000000&
      Caption         =   "Set alarm"
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
      Begin VB.CheckBox chkAlarmBox 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtAlarmTime 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Text            =   "Enter alarm time"
         Top             =   295
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   600
   End
   Begin MediaPlayerCtl.MediaPlayer mpMP3WAV 
      Height          =   615
      Left            =   135
      TabIndex        =   6
      Top             =   2400
      Width           =   4215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   30
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1470
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   5685
   End
End
Attribute VB_Name = "Playwhat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AlarmTime, AlarmSounded, i, CurTime
Dim Playing As Boolean
Const conMinimized = 1

Private Sub chkAlarmBox_Click()
AlarmTime = txtAlarmTime.Text
    If AlarmTime = "" Then Exit Sub
    If Not IsDate(AlarmTime) Then
        chkAlarmBox.Value = False
        MsgBox "The time you entered was not valid."
    Else
        AlarmTime = CDate(AlarmTime)
    End If
End Sub

Private Sub cmdExit_Click()
mpMP3WAV.Stop
mpMP3WAV.SelectionEnd = True
End
End Sub

Private Sub Form_Load()
    AlarmTime = ""
    mpMP3WAV.Visible = False
    cmdOpen.Visible = True
End Sub

Private Sub Form_Resize()
    If WindowState = conMinimized And txtAlarmTime.Text = "Enter alarm time" Then      ' If form is minimized, display the time in a caption.
        SetCaptionTime
    ElseIf WindowState = conMinimized And chkAlarmBox.Value = 1 And mpMP3WAV.FileName <> "" = True Then
        Caption = "Alarm has been set!"
    Else
        Caption = "Alarm Clock"
    End If
End Sub

Private Sub SetCaptionTime()
    Caption = Format(Time, "Medium Time")   ' Display time using medium time format.
End Sub

Private Sub OpenFiles()
    CommonDialog1.CancelError = True
    'User hits Cancel
    On Error GoTo ErrHandler
    CommonDialog1.Flags = cdlOFNHideReadOnly
    'Specifies what kind of files to open
    CommonDialog1.Filter = "MP3 Files (*.mp3)|*.mp3|MP3 Playlists (*.m3u)|*.m3u|WAV Files (*.wav)|*.wav"
    'Default file to open
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    mpMP3WAV.FileName = CommonDialog1.FileName
    Playwhat.Caption = mpMP3WAV.FileName
    If chkAlarmBox.Value = 1 Then
        mpMP3WAV.Stop
    End If

    Exit Sub
    'User has hit Cancel
ErrHandler:
    If mpMP3WAV.FileName = "" Then
    MsgBox "no mp3 or wav"
    End If
    Exit Sub
End Sub

Private Sub cmdOpen_Click()
  OpenFiles
End Sub




Private Sub Timer1_Timer()
    
    'CurTime = TimeValue(Time)
    'If lblTime.Caption <> CStr(Time) Then

        
    If AlarmTime >= Time Then
       
        If Time >= AlarmTime And Not AlarmSounded Then
            mpMP3WAV.Play
            AlarmSounded = True
        ElseIf Time < AlarmTime Then
            AlarmSounded = False
        End If
        
    End If
    
        If WindowState = conMinimized And txtAlarmTime.Text = "Enter alarm time" Then
            If Minute(CDate(Caption)) <> Minute(Time) Then SetCaptionTime
        Else
            lblTime.Caption = Format$(Time, "h:mm am/pm")
            'lblTime.Caption = Format$(Time, "Long Time")
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
mpMP3WAV.Stop
mpMP3WAV.SelectionEnd = True
End Sub

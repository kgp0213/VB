VERSION 5.00
Object = "{D82C0000-9B4E-412B-851E-1DCB662ADC4F}#1.0#0"; "DJMeter.ocx"
Begin VB.Form TestMeter 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "¿ªÊ¼"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   100
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin ProgressMeter.DJMeter DJMeter1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      Caption         =   "DJMeter1"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   45
   End
End
Attribute VB_Name = "TestMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim lngX As Long
    Dim lngY As Long
    
    For lngY = 0 To 100
         DJMeter1.Percent = lngY
         For lngX = 1 To 1000
            DoEvents
        Next lngX
    Next lngY
    
End Sub


Private Sub HScroll1_Change()

    DJMeter1.Percent = HScroll1.Value
    
End Sub

Private Sub DJMeter1_Change()

    lblPercent.Caption = DJMeter1.Percent & "%"
    
End Sub



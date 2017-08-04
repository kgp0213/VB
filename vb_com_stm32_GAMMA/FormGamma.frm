VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F1EB48E5-7E03-41F5-A4D0-CA86119EF992}#73.0#0"; "CaControl.ocx"
Object = "{F0971ADD-CEF2-46B3-8D7F-C075DE0316B1}#18.0#0"; "MinoltaColorSpaceControl.ocx"
Begin VB.Form FormGamma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gamma V1.02 2017.1.20"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormGamma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   13860
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox ColorRgb2Bgr 
      Caption         =   "R/B交换"
      Height          =   255
      Left            =   12000
      TabIndex        =   46
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   960
      ScaleHeight     =   1035
      ScaleWidth      =   7275
      TabIndex        =   44
      Top             =   4800
      Width           =   7335
      Begin VB.Label Label3 
         Caption         =   "请先连接CA310，大约需要30S"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   45
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.CommandButton CommandCloseConnect 
      Caption         =   "断开"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   43
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox CMD1Text 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   41
      Text            =   "阶数回显"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TextView 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   39
      Text            =   "FormGamma.frx":58C3A
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "联机"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      TabIndex        =   35
      Top             =   2640
      Width           =   2295
      Begin VB.CommandButton comSett 
         Caption         =   "COM"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton CommandConnect 
         Caption         =   "连接CA310"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command0Cal 
         Caption         =   "校准"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   675
      End
      Begin VB.Image imgledoff 
         Height          =   270
         Left            =   840
         Picture         =   "FormGamma.frx":58C43
         Top             =   720
         Width           =   285
      End
      Begin VB.Image imgledon 
         Height          =   300
         Left            =   840
         Picture         =   "FormGamma.frx":590BD
         Top             =   840
         Width           =   300
      End
   End
   Begin VB.CheckBox gammaTestMode 
      Caption         =   "四色模式"
      Height          =   255
      Left            =   12000
      TabIndex        =   34
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "自定义Gamma画面"
      Height          =   2295
      Left            =   6960
      TabIndex        =   29
      Top             =   120
      Width           =   6855
      Begin VB.CheckBox CheckgrayValueText 
         Caption         =   "锁定"
         Height          =   255
         Left            =   5760
         TabIndex        =   42
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton CommandmanualMeasure 
         Caption         =   "ManualMeasure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox grayValueText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         MaxLength       =   1000
         TabIndex        =   30
         Text            =   "00,16,32,48,64,80,96,112,128,144,160,176,192,208,224,240,255"
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   4200
         TabIndex        =   48
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Readme"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Label laberGrayinfo 
         Caption         =   "当前阶数："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   4095
      End
   End
   Begin MinoltaCaControl.CaControl CaControl1 
      Height          =   555
      Left            =   7080
      TabIndex        =   27
      Top             =   3120
      Width           =   2595
      _ExtentX        =   3625
      _ExtentY        =   873
   End
   Begin VB.TextBox TextIntervalSec 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   7680
      TabIndex        =   24
      Text            =   "5"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton CommandStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12000
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame FrameRefData 
      Caption         =   "Ref. xyLv"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      TabIndex        =   5
      Top             =   1980
      Width           =   2295
      Begin VB.Label LabelRefData 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LabelRefData 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LabelRefData 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1500
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FrameCurrentData 
      Caption         =   "Current Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   2295
      Begin MSComCtl2.UpDown UpDownCurrentData 
         Height          =   195
         Left            =   2040
         TabIndex        =   21
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   344
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelData 
         Caption         =   "x:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   315
      End
      Begin VB.Label LabelData 
         Caption         =   "y:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   780
         Width           =   315
      End
      Begin VB.Label LabelData 
         Caption         =   "Lv:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label LabelDataVal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label LabelDataVal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   600
         TabIndex        =   11
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label LabelDataVal 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   1260
         Width           =   1575
      End
   End
   Begin MSComCtl2.UpDown UpDownGraph 
      Height          =   195
      Left            =   6600
      TabIndex        =   2
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   344
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Frame FrameMsrSpec 
      Caption         =   "Measurement Spec."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   9720
      TabIndex        =   16
      Top             =   2520
      Width           =   1995
      Begin VB.OptionButton Option16 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option256 
         Caption         =   "256"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option128 
         Caption         =   "128"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   675
      End
      Begin VB.OptionButton Option32 
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option64 
         Caption         =   "64"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton CommandMeasure 
      Caption         =   "Measure"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10080
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "Save..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9240
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmdDiag 
      Left            =   9120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton xlsClear 
      Caption         =   "清除"
      Height          =   375
      Left            =   9240
      TabIndex        =   33
      Top             =   5400
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdDataList 
      Height          =   2715
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   1041
      Cols            =   13
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   2400
      ScaleHeight     =   3765
      ScaleWidth      =   4410
      TabIndex        =   3
      Top             =   0
      Width           =   4440
   End
   Begin MinoltaxyControl.xyControl xyControl1 
      Height          =   4275
      Left            =   2160
      TabIndex        =   28
      Top             =   -120
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   7541
   End
   Begin VB.Label Label5 
      Caption         =   "YJ2017"
      Height          =   375
      Left            =   13200
      TabIndex        =   49
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label LabelCom 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   47
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "x100ms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   26
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label LabelWait 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   25
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "FormGamma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public cmdBusyFlag As Boolean
Dim WithEvents mycomm  As MSComm
Attribute mycomm.VB_VarHelpID = -1
Option Explicit

Public WithEvents objCaControl As CaControl
Attribute objCaControl.VB_VarHelpID = -1
'Public WithEvents objVGControl As VGControl

Dim typMeasurementData(1040) As TypeMeasurementData
Dim lCount As Long

Dim lMeasureMode As Long
Dim lDisplayMode As Long
Dim lDataType As Long
Dim lGraghType As Long

Const COLOR_MODE As Long = 0
Const FMA_MODE As Long = 1
Const JEITA_MODE As Long = 2

Const NO_DATA As Long = -9999
Dim ListNo As Integer               'DataNo.
Dim SelectDataName As String        'ModeName
Dim bStop As Boolean                'Stop Flag

Dim Mydata(300, 4) As Single
Dim sRedGamma(256) As Single
Dim sGreenGamma(256) As Single
Dim sBlueGamma(256) As Single
Dim sWhiteGamma(256) As Single
Dim lMeasureSpec As Long
Dim lSelectedRow0 As Long
Dim lSelectedRow1 As Long


Const MSR_16 As Long = 16
Const MSR_32 As Long = 32
Const MSR_64 As Long = 64
Const MSR_128 As Long = 128
Const MSR_256 As Long = 256
Dim WaitTime As Double

Sub SetGraphData()
    Dim i As Integer, j As Integer
    
    If lMeasureSpec = MSR_256 Then
        j = lMeasureSpec - 1
    Else
        j = lMeasureSpec
    End If
    For i = 0 To j
        Mydata(i + 1, 1) = sRedGamma(i)
        Mydata(i + 1, 2) = sGreenGamma(i)
        Mydata(i + 1, 3) = sBlueGamma(i)
        Mydata(i + 1, 4) = sWhiteGamma(i)
    Next i
    
    SetGraph
    
End Sub
Sub GridInit()
    
    Dim i As Integer
    
    xyControl1.ClearData
    picGraph.Cls
    grdDataList.Clear
    
    grdDataList.Cols = 13
    xyControl1.Visible = True
    
    grdDataList.FocusRect = flexFocusHeavy
    grdDataList.HighLight = flexHighlightAlways
    grdDataList.Row = 0
    grdDataList.Col = 0: grdDataList.Text = "No."
    grdDataList.Col = 1: grdDataList.Text = "X"
    grdDataList.Col = 2: grdDataList.Text = "Y"
    grdDataList.Col = 3: grdDataList.Text = "Z"
    grdDataList.Col = 4: grdDataList.Text = "x"
    grdDataList.Col = 5: grdDataList.Text = "y"
    
    ' 021225
    ' grdDataList.Col = 6: grdDataList.Text = "Lv"
    grdDataList.Col = 6: grdDataList.Text = gstrLvOrEv
    
    grdDataList.Col = 7: grdDataList.Text = "ud"
    grdDataList.Col = 8: grdDataList.Text = "vd"
    grdDataList.Col = 9: grdDataList.Text = "T"
    grdDataList.Col = 10: grdDataList.Text = "duv"
    grdDataList.Col = 11: grdDataList.Text = "Date"
    grdDataList.Col = 12: grdDataList.Text = "Time"
    
    grdDataList.ColWidth(0) = 420 '380   'No
    
    grdDataList.ColWidth(1) = 600   'X
    grdDataList.ColWidth(2) = 600   'Y
    grdDataList.ColWidth(3) = 600   'Z
    
    grdDataList.ColWidth(4) = 600   'x
    grdDataList.ColWidth(5) = 600   'y
    grdDataList.ColWidth(6) = 950   'Lv
    
    grdDataList.ColWidth(7) = 600   'ud
    grdDataList.ColWidth(8) = 600   'vd
    
    grdDataList.ColWidth(9) = 600   'T
    grdDataList.ColWidth(10) = 580   'duv
    grdDataList.ColWidth(11) = 950  'Date
    grdDataList.ColWidth(12) = 700   'Time
    
    grdDataList.Col = 0
    For i = 1 To 1040
        grdDataList.Row = i
        grdDataList.Text = Format(i)
    Next i
    
    ListNo = 1

    grdDataList.TopRow = 1

End Sub
Sub SaveData()

    Dim dd(1040, 12) As String
    Dim i As Integer, j As Integer
    Dim fm As String, fnum As Integer, fname As String

    For i = 1 To ListNo - 1
        grdDataList.Row = i
        For j = 1 To 12
            grdDataList.Col = j
            dd(i, j) = grdDataList.Text
        Next j
    Next i
    
    On Error Resume Next
    cmdDiag.CancelError = True
    cmdDiag.FileName = ""
    cmdDiag.Filter = "Data Files (*.csv)|*.csv"
    cmdDiag.FilterIndex = 2
    cmdDiag.Action = 2
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        fm = cmdDiag.FileName
        fm = Mid$(fm, 1, InStr(1, fm, ".")) + "csv"
    End If
    
    fname = Dir$(fm, vbNormal Or vbReadOnly)
    If fname <> "" Then
        If MsgBox(fname & " Overwrite. OK ?", vbOKCancel) = vbCancel Then Exit Sub
    Else
    End If
    
    fnum = FreeFile
    Open fm For Output Access Write Shared As fnum
    Print #fnum, "Gamma"
    
    ' 021225
    ' Write #fnum, "No.", "X", "Y", "Z", "x", "y", "Lv", "ud", "vd", "T", "duv", "Date", "Time"
    Write #fnum, "No.", "X", "Y", "Z", "x", "y", gstrLvOrEv, "ud", "vd", "T", "duv", "Date", "Time"
    
    For i = 1 To ListNo - 1
        Print #fnum, Format(i, "000"); ",";
        For j = 1 To 11
            Print #fnum, dd(i, j); ",";
        Next j
        Print #fnum, dd(i, 12)
    Next i
    
    Write #fnum, "No.", "Red", "Green", "Blue", "White"
    
    If lMeasureSpec = MSR_256 Then
        j = lMeasureSpec - 1
    Else
        j = lMeasureSpec
    End If
    
    For i = 0 To j
        Print #fnum, Format(i, "000"); ",";
        Print #fnum, sRedGamma(i); ","; sGreenGamma(i); ","; sBlueGamma(i); ","; sWhiteGamma(i)
    Next i
    Close fnum
End Sub

Private Sub CaControl1_Update()
    objMemory.GetReferenceColor objProbe.ID, typCurrentRefereceData.sRefx, typCurrentRefereceData.sRefy, typCurrentRefereceData.sRefLv
    LabelRefData(0).Caption = Round(typCurrentRefereceData.sRefx, 4)
    LabelRefData(1).Caption = Round(typCurrentRefereceData.sRefy, 4)
    LabelRefData(2).Caption = Round(typCurrentRefereceData.sRefLv, 4)
    lDisplayMode = objCa.DisplayMode
    Select Case lDisplayMode
        Case DSP_LXY
            SelectDataName = "COLOR"
        Case DSP_DUV
            SelectDataName = "COLOR"
        Case DSP_ANL
            SelectDataName = "COLOR"
        Case DSP_ANLG
            SelectDataName = "COLOR"
        Case DSP_ANLR
            SelectDataName = "COLOR"
        Case DSP_PUV
            SelectDataName = "COLOR"
 '       Case DSP_FMA
 '           SelectDataName = "FMA"
        Case DSP_XYZ
            SelectDataName = "COLOR"
        Case Else
'            SelectDataName = "JEITA"
            SelectDataName = "COLOR"
    End Select
    GridInit
    DoEvents
    
End Sub


Private Sub Check1_Click()

End Sub

Private Sub CheckgrayValueText_Click()
grayValueText.Locked = Not grayValueText.Locked
End Sub

Private Sub Colorrgb2bgr_Click()
If ColorRgb2Bgr.Value = 1 Then
rgbChange = True
Else: rgbChange = False
End If
End Sub

Private Sub Command0Cal_Click()
    On Error GoTo E
    objCa.CalZero
    Exit Sub
E:
    '===================================
    ' Error Trap
    '===================================
    Dim strERR As String
    Dim iReturn As Integer
    
    strERR = "Error from " + Err.Source + Chr$(10) + Chr$(13)
    strERR = strERR + Err.Description + Chr$(10) + Chr$(13)
    strERR = strERR + "HRESULT " + CStr(Err.Number - vbObjectError)
    iReturn = MsgBox(strERR, vbAbortRetryIgnore)
    Select Case iReturn
        Case vbAbort: End
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub CommandCloseConnect_Click()
 objCa.RemoteMode = 0
       ' CommandConnect.BackColor = vbRed
Command0Cal.Enabled = False
CommandMeasure.Enabled = False
'CheckgrayValueText.Value = Checked
CommandmanualMeasure.Enabled = False
TextView.Text = "确认数据保存后再关闭程序"
End Sub


Private Sub sCommandConnect_Click()
'SCommandConnect_Click
'FormCa310Connect.Label1.Caption = "正在连接CA310, 请等待联机完成..."
'FormCa310Connect.Show 'vbModal
End Sub


Public Sub CommandConnect_Click()
'Me.Hide
'FormCa310Connect.Show vbModal

'FormCa310Connect.Label1.Caption = "正在连接CA310, 请等待联机完成..."
'Picture1.Visible = True   ' 提示： 正在连接CA310, 请等待联机完成..."
cmdBusyFlag = True
Label3.Caption = "CA310联机中，请静候连接完成..."
DoEvents

StartMain
CommandConnect.BackColor = &H8000000F
CommandmanualMeasure.Enabled = True
Command0Cal.Enabled = True
CommandCloseConnect.Enabled = True
'Unload FormCa310Connect
Picture1.Visible = False   ' 结束提示： 正在连接CA310, 请等待联机完成..."
CommandMeasure.Enabled = True
cmdBusyFlag = False
End Sub



Private Sub CommandMeasure_Click()
    Dim bResult As Boolean
    CMD1Text.FontSize = 12
    
    gammaTestMode.Enabled = False
    ' 030407
    If objCa.DisplayMode <> COLOR_MODE Then
        objCa.DisplayMode = DSP_LXY
    End If
    
    cmdBusyFlag = True
    
   ' cmdWin.Enabled = False
    bStop = False
    CommandMeasure.Enabled = False
    CommandStop.Enabled = True
    FrameMsrSpec.Enabled = False
    CommandmanualMeasure.Enabled = False
    DoEvents
    
    MeasureGamma
    CommandmanualMeasure.Enabled = True
    CommandMeasure.Enabled = True
    CommandStop.Enabled = False
    'cmdWin.Enabled = True
    FrameMsrSpec.Enabled = True
    gammaTestMode.Enabled = True
    cmdBusyFlag = False

End Sub

Private Sub CommandSave_Click()
    
    SaveData

End Sub


Private Sub CommandStop_Click()

    bStop = True

End Sub

Private Sub comSett_Click()
frmSet.Show  'vbModal
End Sub

Private Sub Form_Activate()

    'typCurrentMeasurementData.lColorStatus = NO_DATA
    'typCurrentMeasurementData.lColorStatus = NO_DATA
    'typCurrentMeasurementData.lColorStatus = NO_DATA
    'SetCurrentData

    LabelRefData(0).Caption = Round(typCurrentRefereceData.sRefx, 4)
    LabelRefData(1).Caption = Round(typCurrentRefereceData.sRefy, 4)
    LabelRefData(2).Caption = Round(typCurrentRefereceData.sRefLv, 4)
    
'    lDisplayMode = objCa.DisplayMode
    SelectDataName = "COLOR"
    
    'GridInit
    xyControl1.Visible = True
    
    'CommandMeasure.Enabled = True
    CommandStop.Enabled = False
    
    If mycomm.PortOpen = True Then
    imgledoff.Visible = False
    imgledon.Visible = True
    Else
    imgledoff.Visible = True
    imgledon.Visible = False
    End If

End Sub
Private Sub Form_Load()
dubuggFlag = 0
m = 0
grayNumflag = 0
cmdBusyFlag = True
'Call grayValueText_Change
'Unload frmSet
Set mycomm = frmSet.MSComm1
If (mycomm.PortOpen = True) Then
    
      imgledon.Visible = True
      imgledoff.Visible = False
      
      
     Else
          'imgledon.Visible = False
          'imgledoff.Visible = True
        On Error Resume Next
         frmSet.MSComm1.PortOpen = True
         If (frmSet.MSComm1.PortOpen = False) Then
         imgledon.Visible = False
         imgledoff.Visible = True
         
         MsgBox "请先确认本机串口可以正常工作", vbExclamation + vbOKOnly, "友情提醒"
       End If
End If

Label2.Caption = "上述数据以0开始、以255结尾，数据间以逗号分隔"
'Me.AcceptButton = Me.CommandCconnect
'Me.CommandConnect.SetFocus
CommandConnect.BackColor = vbRed
Command0Cal.Enabled = False
CommandMeasure.Enabled = False
CheckgrayValueText.Value = Checked
CommandmanualMeasure.Enabled = False
CommandCloseConnect.Enabled = False
'Picture1.Visible = False
 ColorRgb2Bgr.Enabled = False
 rgbChange = False
 
Label5.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Sub xForm_Initialize()
   ' Dim ComPortNo As Long, ErrCheck As Boolean
    
    On Error GoTo E
    
    Set objCaControl = CaControl1
   ' Set objVGControl = VGControl1
    
   ' ComPortNo = 1
   ' objVGControl.ControlInitialize ComPortNo, ErrCheck
        
    typCurrentMeasurementData.lColorStatus = NO_DATA
    typCurrentMeasurementData.lColorStatus = NO_DATA
    typCurrentMeasurementData.lColorStatus = NO_DATA
    SetCurrentData

    lDisplayMode = objCa.DisplayMode
    SelectDataName = "COLOR"
    
    GridInit
    xyControl1.Visible = True
    
    'CommandMeasure.Enabled = True
    CommandStop.Enabled = False
    
    'If ErrCheck = False Then
    ' objVGControl.Visible = False
   ' End If
    'Exit Sub

E:
    'Dim strERR As String
    'Dim iReturn As Integer
    
    'strERR = "Error from " + Err.Source + Chr$(10) + Chr$(13)
    'strERR = strERR + Err.Description + Chr$(10) + Chr$(13)
    'strERR = strERR + "HRESULT " + CStr(Err.Number - vbObjectError)
    'iReturn = MsgBox(strERR, vbOKOnly)
   ' objVGControl.Visible = False
    Resume Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
   Unload frmSet
   Unload FormStart
    
    If Me.Tag <> "END" Then
        'If lDisplayMode <> objCa.DisplayMode Then
       '     objCa.DisplayMode = objCa.DisplayMode
       ' End If
      '  FormGamma.Hide
      '  Cancel = True  'cancel 参数为 True 可防止窗体被卸载
    Else
        objCa.RemoteMode = 0
    End If
   ' If FormVisibleFlg = True Then
       ' Unload FormWinPtn
   ' End If
   Cancel = 0
   Unload Me
 '  ExitProcess 0
    End
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub



Public Sub SetCurrentData()
    Select Case lDataType
        Case DSP_LXY
            LabelData(0).Caption = "x:"
            LabelData(1).Caption = "y:"
            
            '021225
            ' LabelData(2).Caption = "Lv:"
            LabelData(2).Caption = gstrLvOrEv + ":"
            
            If typCurrentMeasurementData.lColorStatus = NO_DATA Then
                LabelDataVal(0).Caption = "-----"
                LabelDataVal(1).Caption = "-----"
                LabelDataVal(2).Caption = "-----"
            Else
                LabelDataVal(0).Caption = Format(typCurrentMeasurementData.ssx, FORMAT_SXY)
                LabelDataVal(1).Caption = Format(typCurrentMeasurementData.ssy, FORMAT_SXY)
                LabelDataVal(2).Caption = Format(typCurrentMeasurementData.sLv, FORMAT_LV)
            End If
        Case DSP_DUV
            LabelData(0).Caption = "T:"
            LabelData(1).Caption = "duv:"
            
            ' 021225
            ' LabelData(2).Caption = "Lv"
            LabelData(2).Caption = gstrLvOrEv + ":"
            
            If typCurrentMeasurementData.lColorStatus = NO_DATA Then
                LabelDataVal(0).Caption = "-----"
                LabelDataVal(1).Caption = "-----"
                LabelDataVal(2).Caption = "-----"
            Else
                LabelDataVal(0).Caption = Format(typCurrentMeasurementData.LT, "00000")
                LabelDataVal(1).Caption = Format(typCurrentMeasurementData.sduv, "+.000;-.000") 'Format(typCurrentMeasurementData.sduv, "0.0000")
                LabelDataVal(2).Caption = Format(typCurrentMeasurementData.sLv, FORMAT_LV)
            End If
        Case DSP_PUV
            LabelData(0).Caption = "ud:"
            LabelData(1).Caption = "vd:"
            
            ' 021225
            ' LabelData(2).Caption = "Lv"
            LabelData(2).Caption = gstrLvOrEv + ":"
            
            If typCurrentMeasurementData.lColorStatus = NO_DATA Then
                LabelDataVal(0).Caption = "-----"
                LabelDataVal(1).Caption = "-----"
                LabelDataVal(2).Caption = "-----"
            Else
                LabelDataVal(0).Caption = Format(typCurrentMeasurementData.sud, FORMAT_SXY)
                LabelDataVal(1).Caption = Format(typCurrentMeasurementData.svd, FORMAT_SXY)
                LabelDataVal(2).Caption = Format(typCurrentMeasurementData.sLv, FORMAT_LV)
            End If
'        Case DSP_FMA
'            LabelData(0).Caption = "FMA:"
'            LabelData(1).Caption = ""
'            LabelData(2).Caption = ""
'            If typCurrentMeasurementData.lFMAStatus = NO_DATA Then
'                LabelDataVal(0).Caption = "-----"
'                LabelDataVal(1).Caption = ""
'                LabelDataVal(2).Caption = ""
'            Else
'                LabelDataVal(0).Caption = Format(typCurrentMeasurementData.sFMA, "0.0")
'                LabelDataVal(1).Caption = ""
'                LabelDataVal(2).Caption = ""
'            End If
        Case DSP_XYZ
            LabelData(0).Caption = "X:"
            LabelData(1).Caption = "Y:"
            LabelData(2).Caption = "Z:"
            If typCurrentMeasurementData.lColorStatus = NO_DATA Then
                LabelDataVal(0).Caption = "-----"
                LabelDataVal(1).Caption = "-----"
                LabelDataVal(2).Caption = "-----"
            Else
                LabelDataVal(0).Caption = Format(typCurrentMeasurementData.Sx, "0.00")
                LabelDataVal(1).Caption = Format(typCurrentMeasurementData.Sy, "0.00")
                LabelDataVal(2).Caption = Format(typCurrentMeasurementData.Sz, "0.00")
            End If
'        Case DSP_JEITA
'            LabelData(0).Caption = "JEITA:"
'            LabelData(1).Caption = ""
'            LabelData(2).Caption = ""
'            If typCurrentMeasurementData.lJEITAStatus = NO_DATA Then
'                LabelDataVal(0).Caption = "-----"
'                LabelDataVal(1).Caption = ""
'                LabelDataVal(2).Caption = ""
'            Else
'                LabelDataVal(0).Caption = Format(typCurrentMeasurementData.sJEITA, "0.0")
'                LabelDataVal(1).Caption = ""
'                LabelDataVal(2).Caption = ""
'            End If
    End Select
    
End Sub




Private Sub gammaTestMode_Click()
If gammaTestMode.Value = 1 Then
ColorRgb2Bgr.Enabled = True
Else: ColorRgb2Bgr.Enabled = False
End If
End Sub

Private Sub grayValueText_Change()
Dim i, j As Integer
Dim a
Dim b$()
Dim c$()

 
  a = grayValueText.Text
  a = Replace(a, " ", "")   '去空格
  
  
  If a = "" Then   '防止空数据
    grayValueText.Text = "0,"
    grayValueText.SelStart = Len(grayValueText.Text)
     a = "0"
  End If
  
  b = Split(a, ",", -1, 1)  '数组a的数据去逗号后填入b
  
  'ReDim gstr(UBound(b))
  
  m = UBound(b)   '获取数组b的最大下标
    
  ReDim c(m)
   
  j = 0
  i = 0
For i = 0 To m  '轮询把不是空的数据传到数组c
 
     If (b(i) <> "") Then
  
     c(j) = b(i)
     j = j + 1
     Else
     End If
 
Next
    
    
  Do While (c(m) = "")  '把数组c不为空的最大下标找出来
  m = m - 1
  If m = 0 Then
 ' m = 1
  Exit Do
  End If
  Loop
  
  
  
  
  ReDim Preserve c(m)  '缩小数组大小，把数组结尾的空数组舍弃掉
  ReDim gstr(m)        ',确定数组gstr大小

 
For i = 0 To m   ' 把处理后的数据送到gstr
                            '判断是否为10进制数据
     If (c(i) Like "*[!0-9, ]*") Then
    ' MsgBox "数据格式非法", vbExclamation + vbOKOnly, "友情提醒"
    Label4.Caption = "数据格式非法"
     CommandmanualMeasure.Enabled = False
     Exit Sub
     End If
     
     gstr(i) = Val(c(i))
     If gstr(i) > 255 Then
     'MsgBox "数据超范围，请修改使其处于范围（0～255）", vbExclamation + vbOKOnly, "友情提醒"
     Label4.Caption = "数据超范围"
     CommandmanualMeasure.Enabled = False
     Exit Sub
     End If
         
Next

If Not cmdBusyFlag Then
CommandmanualMeasure.Enabled = True
Else: CommandmanualMeasure.Enabled = False
End If
Label4.Caption = ""
End Sub

Private Sub grayShow_Click()
 '判断数据长度
 If m < 9 Then
 MsgBox "数据过少", vbExclamation + vbOKOnly, "友情提醒"
 'Exit Function
 Exit Sub
 End If
laberGrayinfo.Caption = "当前阶数 " + CStr(gstr(grayNumflag)) + "/255" 'CStr


Dim tep As String
tep = "5A55120400" + CStr(Replace((Format(Hex(gstr(grayNumflag)), "@@")), " ", "0")) + "00AA"
FormGamma.CMD1Text.Text = gstr(grayNumflag)
Call frmSet.comSend(tep)

''0~9转换成16进制时，转成00 01~09的格式，以符合协议要求

If grayNumflag < m Then
grayNumflag = grayNumflag + 1
Else
grayNumflag = 0
End If
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lSelectedRow0 = grdDataList.MouseRow
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim L As Long
    
    xyControl1.SetVisibleAll False
    lSelectedRow1 = grdDataList.MouseRow
    If lSelectedRow0 > lSelectedRow1 Then
        For L = lSelectedRow1 To lSelectedRow0
            If L <= ListNo Then
                xyControl1.SetVisible L
            Else
                Exit Sub
            End If
        Next L
    Else
        For L = lSelectedRow0 To lSelectedRow1
            If L <= ListNo Then
                xyControl1.SetVisible L
            Else
                Exit Sub
            End If
        Next L
    End If
End Sub


Private Sub LabelView_Click()

End Sub

Private Sub TextIntervalSec_KeyPress(KeyAscii As Integer)
    Dim tl As Integer
    Select Case Chr$(KeyAscii)
        Case "0" To "9":                  ' 0 1 2 3 4 5 6 7 8 9
        Case Chr$(vbKeyBack):              ' Back Space
            Exit Sub
        Case Else:                        '
        KeyAscii = 0
    End Select
    tl = TextIntervalSec.SelLength
    If tl > 0 And tl <= 3 Then
    ElseIf Len(TextIntervalSec.Text) > 2 Then
        KeyAscii = 0
    End If

End Sub

Private Sub TextIntervalSec_LostFocus()
    
    If TextIntervalSec.Text = "" Then Exit Sub
    If CLng(TextIntervalSec.Text) > 600 Or CLng(TextIntervalSec.Text) < 0 Then
        MsgBox "Interval Second 0-600", vbOKOnly
        TextIntervalSec.Text = "0"
    End If
End Sub
Private Sub SetWait()
Dim t1 As Double
t1 = Timer
Do While Timer < t1 + WaitTime
        DoEvents
        If bStop = True Then Exit Sub
Loop

End Sub

Private Sub UpDownCurrentData_DownClick()
    
    Select Case lDataType
        Case DSP_LXY
            lDataType = DSP_DUV
        Case DSP_DUV
            lDataType = DSP_PUV
        Case DSP_PUV
            lDataType = DSP_XYZ
        Case DSP_XYZ
            lDataType = DSP_LXY
        Case Else
            lDataType = DSP_LXY
    End Select
    SetCurrentData

End Sub

Private Sub UpDownCurrentData_UpClick()
    
    Select Case lDataType
        Case DSP_LXY
            lDataType = DSP_XYZ
        Case DSP_DUV
            lDataType = DSP_LXY
        Case DSP_PUV
            lDataType = DSP_DUV
        Case DSP_XYZ
            lDataType = DSP_PUV
        Case Else
            lDataType = DSP_LXY
    End Select
    SetCurrentData

End Sub

Sub SetGammaData(ByVal LisNo As Integer, lClr As Long)
    
    grdDataList.Row = LisNo
    
    grdDataList.Col = 1
    grdDataList.Text = Format(typCurrentMeasurementData.Sx, "0.00")
    grdDataList.Col = 2
    grdDataList.Text = Format(typCurrentMeasurementData.Sy, "0.00")
    grdDataList.Col = 3
    grdDataList.Text = Format(typCurrentMeasurementData.Sz, "0.00")
    grdDataList.Col = 4
    grdDataList.Text = Format(typCurrentMeasurementData.ssx, FORMAT_SXY)
    grdDataList.Col = 5
    grdDataList.Text = Format(typCurrentMeasurementData.ssy, FORMAT_SXY)
    grdDataList.Col = 6
    grdDataList.Text = Format(typCurrentMeasurementData.sLv, FORMAT_LV)
    grdDataList.Col = 7
    grdDataList.Text = Format(typCurrentMeasurementData.sud, FORMAT_SXY)
    grdDataList.Col = 8
    grdDataList.Text = Format(typCurrentMeasurementData.svd, FORMAT_SXY)
    grdDataList.Col = 9
    If typCurrentMeasurementData.LT = -1 Then
        grdDataList.Text = "-"
    Else
        grdDataList.Text = Format(typCurrentMeasurementData.LT, "00000")
    End If
    grdDataList.Col = 10
    If typCurrentMeasurementData.LT = -1 Then
        grdDataList.Text = "-"
    Else
        grdDataList.Text = Format(typCurrentMeasurementData.sduv, "+.000;-.000")
    End If
    grdDataList.Col = 11
    grdDataList.Text = Format(typCurrentMeasurementData.dateColorData, "yyyy/mm/dd")
    grdDataList.Col = 12
    grdDataList.Text = Format(typCurrentMeasurementData.timeColorData, "hh:mm:ss")

    xyControl1.AddXYGraphData CLng(ListNo)
End Sub


Public Sub SetGraph()
    
    Dim i As Integer, j As Integer, K As Integer
    
    Dim max(4) As Double
    
    Dim X0 As Double, X1 As Double
    Dim Y0 As Double, Y1 As Double
    Dim dv As Single
    Dim MaxDisp As Integer
    
    Dim Col(4) As Long
    Dim xoff As Integer, xw As Integer
    Dim xjn As Integer, yjn As Integer
    Dim keta As String, habax As Single
    Dim xstp As Integer, icofx As Integer
    Dim MaxDataNo As Integer
    
    Dim stt As Integer, enn As Integer
    Dim Moji As String, ofsx As Integer
    Dim divx As Integer, divy As Integer
    
    divy = 10

    xoff = 60
    xw = 1000 - 60 - 60
        
    picGraph.ScaleHeight = 1000
    picGraph.ScaleWidth = 1000
    picGraph.Cls
    picGraph.DrawWidth = 1
    
    ' Display Item Count
    If lMeasureSpec = MSR_16 Then
        MaxDataNo = MSR_16 + 1      '17 data
    ElseIf lMeasureSpec = MSR_32 Then
        MaxDataNo = MSR_32 + 1      '33 data
    ElseIf lMeasureSpec = MSR_64 Then
        MaxDataNo = MSR_64 + 1      '65data
    ElseIf lMeasureSpec = 128 Then
        MaxDataNo = MSR_128 + 1     '129data
    Else
        MaxDataNo = MSR_256         '256data
    End If

    ' Draw Box
    picGraph.Line (xoff, 50)-(xw + xoff, 950), , B
    If MaxDataNo <= 50 Then
        xjn = MaxDataNo
        MaxDisp = MaxDataNo
        xstp = 1
        If MaxDataNo > 25 Then
            xstp = 2
        End If
    Else
        xjn = MaxDataNo
        MaxDisp = MaxDataNo
        xstp = Int((MaxDataNo - 50) / 20 + 0.5) + 2
    End If
    
    For i = 0 To 16
        ' x Scale
        picGraph.Line (xoff + xw / 16 * i, 50)-(xoff + xw / 16 * i, 950 + 5)
    Next i
    yjn = 10
    For i = 0 To yjn
        picGraph.Line (xoff - 5, 50 + 90 * i)-(xw + xoff, 50 + 90 * i)
    Next i
    
    ' get Max
    max(0) = 1#
    max(1) = 1#
    max(2) = 1#
    max(3) = 1#
    
    
    '===========================
    ' Draw Data Graph
    '===========================
    Col(0) = RGB(255, 0, 0)
    Col(1) = RGB(0, 255, 0)
    Col(2) = RGB(0, 0, 255)
    Col(3) = RGB(0, 0, 0)
    
    stt = 0: enn = 3   'Graph 3 Kinds
    
    picGraph.DrawWidth = 2
    For j = stt To enn
        X0 = 0 / MaxDataNo * xw + xoff
        X1 = 1 / MaxDataNo * xw + xoff
        Y0 = 950 - Mydata(1, j + 1) / max(j) * 900  '950 is 0 Line. width is 900.
        Y1 = 950 - Mydata(2, j + 1) / max(j) * 900
        picGraph.Line (X0, Y0)-(X1, Y1), Col(j)
        
        For i = 2 To MaxDataNo
            X0 = i / MaxDataNo * xw + xoff
            Y0 = 950 - Mydata(i, j + 1) / max(j) * 900
            picGraph.Line -(X0, Y0), Col(j)
        Next i
    Next j
    
    '===========================
    ' Print Graph Info.
    '===========================
    picGraph.FontSize = 6
    picGraph.FontBold = False
    
    picGraph.Line (100, 25)-(200, 25), Col(0)
    picGraph.Line (300, 25)-(400, 25), Col(1)
    picGraph.Line (500, 25)-(600, 25), Col(2)
    picGraph.Line (700, 25)-(800, 25), Col(3)
    picGraph.CurrentX = 70: picGraph.CurrentY = 5: picGraph.Print " R"
    picGraph.CurrentX = 270: picGraph.CurrentY = 5: picGraph.Print " G"
    picGraph.CurrentX = 470: picGraph.CurrentY = 5: picGraph.Print " B"
    picGraph.CurrentX = 670: picGraph.CurrentY = 5: picGraph.Print " W"
    
    ' y(axis) Label
    dv = max(0) / divy
    If (max(0) - dv * i) - Int(max(0) - dv * i) <> 0 Then
        keta = "0.0"
    Else
        keta = "0"
    End If
    For i = 0 To divy
        picGraph.CurrentX = 5
        picGraph.CurrentY = 50 + 900 / divy * i - 15
        picGraph.Print Format(max(0) - dv * i, keta)
    Next i
    
    
    
    ' x Label
    For i = 0 To 16
        habax = xw / 16
        If i = 0 Then
            icofx = 0
        ElseIf Len(Format(256 / 16 * i)) = 2 Then
            icofx = 0
        Else '3 keta
            icofx = 8
        End If
        picGraph.CurrentX = i * habax + 50 - icofx - 5
        picGraph.CurrentY = 900 + 65
        If i = 16 Then
            picGraph.Print "255"
        Else
            picGraph.Print 256 / 16 * i
        End If
    Next i
Exit Sub

ER:
    Resume Next
    Return
End Sub

Private Sub CommandmanualMeasure_Click()
grayValueText_Change
If Label4.Caption <> "" Then   '若Label4.Caption不为空则说明数据异常
Exit Sub
End If

   grayValueText.Enabled = False
     bStop = False
    CommandMeasure.Enabled = False
    CommandStop.Enabled = True
    FrameMsrSpec.Enabled = False
    DoEvents
    CommandmanualMeasure.Enabled = False
    DoEvents
    
Call ManualMeasureGamma

    grayValueText.Enabled = True
    CommandMeasure.Enabled = True
    CommandStop.Enabled = False
    FrameMsrSpec.Enabled = True
    CommandmanualMeasure.Enabled = True
   
End Sub
Private Sub ManualMeasureGamma()

    Dim lVLocation As Long
    Dim lHLocation As Long
    Dim lVLocationMax As Long
    Dim lHLocationMax As Long
    Dim strMsg As String
    Dim mm As Integer
    
    On Error Resume Next
    
    WaitTime = Val(TextIntervalSec.Text) / 10#
   
    InitializeData
    GridInit
    
    xyControl1.Visible = True
    picGraph.Visible = False
    
    If SelectDataName <> "COLOR" Then
        objCa.DisplayMode = DSP_LXY
        lDisplayMode = DSP_LXY
        SelectDataName = "COLOR"
    End If
    
   
    
   ' If FormVisibleFlg = True Then
       ' FormWinPtn.SetWin 255, 255, 255
  '  Else
        ' Set Window Pattern
      '  objVGControl.Pattern = 1
        ' Set Video Levle
     '   objVGControl.SetGVideoLevel 255, 255, 255
    'End If
    
    If m < 8 Then
 MsgBox "数据过少", vbExclamation + vbOKOnly, "友情提醒"
 'Exit Function
 Exit Sub
 End If

    
    
    Dim lLevelStep As Long
    Dim lStep As Integer
    Dim tepG As Integer
     mm = UBound(gstr)
     lMeasureSpec = mm
    lLevelStep = 256 / lMeasureSpec
    
'----------------------------------------------------------
 ' if gammaTestMode.Value = False Then
  '   GoTo MeasureWhite
    ' End If
    '=====================
    ' Measure White
    '=====================
    CommandStop.Enabled = True
  
MeasureWhite:
    
    'If FormVisibleFlg = True Then
        SetWin 255, 255, 255
        Call frmSet.comsendGrayNum(255, CLR_WHITE, rgbChange)
        
        If ComState <> "" Then
        LabelCom.Caption = ComState
        Else: LabelCom = ""
        End If
   '
        'Call frmSet.comsendGrayNum(255, CLR_WHITE)
   ' End If
    MsgBox "White Measure!", vbOKOnly
    xyControl1.Visible = True
    picGraph.Visible = False
    DoEvents
    
    'objCa.Measure
    '---------------------
    
    
    With typCurrentMeasurementData
        For lStep = 0 To lMeasureSpec    '量测<255部分的阶数
        
                       
            '    objVGControl.SetGVideoLevel lLevelStep * lStep, lLevelStep * lStep, lLevelStep * lStep
            tepG = gstr(lStep)
            SetWin tepG, tepG, tepG
            DoEvents
            Call frmSet.comsendGrayNum(tepG, CLR_WHITE, rgbChange)
            DoEvents
            laberGrayinfo.Caption = "当前阶数 " + CStr(tepG) + "/255" 'CSt
            
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sWhiteGamma(lStep) = .Sy
            
            LabelDataVal(0).Caption = Format(.ssx, FORMAT_SXY)
            LabelDataVal(1).Caption = Format(.ssy, FORMAT_SXY)
            LabelDataVal(2).Caption = Format(.sLv, FORMAT_LV)
            DoEvents
            Call SetGammaData(ListNo, CLR_WHITE)
            If ListNo > 8 Then
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
            
            If bStop = True Then
                Exit Sub
            End If
        Next lStep
        
          
        xyControl1.Visible = False
        picGraph.Visible = True
        DoEvents
        
    End With
    '----------------------------------------------------------
   If lMeasureSpec <> 16 And lMeasureSpec <> 32 And lMeasureSpec <> 64 Then
   Exit Sub
   End If
   
        For lStep = 0 To lMeasureSpec
            sWhiteGamma(lStep) = sWhiteGamma(lStep) / sWhiteGamma(lMeasureSpec)
        Next lStep
  
    SetGraphData
    
End Sub
Private Sub MeasureGamma()

    Dim lVLocation As Long
    Dim lHLocation As Long
    Dim lVLocationMax As Long
    Dim lHLocationMax As Long
    Dim strMsg As String
    
    
    On Error Resume Next
    
    WaitTime = Val(TextIntervalSec.Text) / 10#
   
    InitializeData
    GridInit
    
    xyControl1.Visible = True
    picGraph.Visible = False
    
    If SelectDataName <> "COLOR" Then
        objCa.DisplayMode = DSP_LXY
        lDisplayMode = DSP_LXY
        SelectDataName = "COLOR"
    End If
    
    If FormVisibleFlg = True Then
        SetWin 255, 255, 255
    Else
        ' Set Window Pattern
      '  objVGControl.Pattern = 1
        ' Set Video Levle
       ' objVGControl.SetGVideoLevel 255, 255, 255
    End If
    
    
    Dim lLevelStep As Long
    Dim lStep As Long
    
    lLevelStep = 256 / lMeasureSpec
    
'----------------------------------------------------------
  If gammaTestMode.Value = False Then
  '如果不是四色模式
     GoTo MeasureWhite
     '直接测试黑白
     End If
    '=====================
    ' Measure Red
    '=====================
    'If FormVisibleFlg = True Then
       ' SetWin 255, 0, 0
        ' Call frmSet.comsendGrayNum(255, CLR_RED)
    'Else
      '  objVGControl.RedSW = True
       ' objVGControl.GreenSW = False
       ' objVGControl.BlueSW = False
   ' End If
    SetWin 255 * 1, 0, 0
    
   'If Not rgbChange Then
            
             '   DoEvents
             Call frmSet.comsendGrayNum(255 * 1, CLR_RED, rgbChange)
             
             '判断数据是否能送出
             If ComState <> "" Then
             LabelCom.Caption = ComState
             Else: LabelCom = ""
             End If
    'Else
           '  SetWin 0, 0, 255
            '    DoEvents
          '  Call frmSet.comsendGrayNum(255, CLR_BLUE, rgbChange)
     'End If
    MsgBox "Red Measure!", vbOKOnly
    'objCa.Measure
    
    ListNo = 1
    With typCurrentMeasurementData
        For lStep = 0 To lMeasureSpec - 1
            'If FormVisibleFlg = True Then
               ' SetWin lLevelStep * lStep, 0, 0
              '  DoEvents
           ' Else
              '  objVGControl.SetGVideoLevel , 0, 0
              ' Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_RED)
           ' End If
          ' If Not rgbChange Then
                 SetWin lLevelStep * lStep, 0, 0
                 DoEvents
                 Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_RED, rgbChange)
          ' Else
            '    SetWin lLevelStep * lStep, 0, 0
             '   DoEvents
             '    Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_BLUE, rgbChange)
          ' End If
            
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sRedGamma(lStep) = .Sx
            SetCurrentData
            DoEvents
            Call SetGammaData(ListNo, CLR_RED)
            If ListNo > 8 Then           'Grid Scroll
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
            
            If bStop = True Then
                Exit Sub
            End If
        Next lStep
                
        If lMeasureSpec <> MSR_256 Then
           ' If FormVisibleFlg = True Then
             '   SetWin 255, 0, 0
             '   DoEvents
           ' Else
              '  objVGControl.SetGVideoLevel 255, 0, 0
             ' Call frmSet.comsendGrayNum(255, CLR_RED)
           ' End If
           
            SetWin 255, 0, 0
          ' If Not rgbChange Then
               '  SetWin 255 * 1, 0, 0
                DoEvents
                Call frmSet.comsendGrayNum(255 * 1, CLR_RED, rgbChange)
          ' Else
                
           '      DoEvents
              '   Call frmSet.comsendGrayNum(255, CLR_BLUE, rgbChange)
           ' End If
           
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sRedGamma(lMeasureSpec) = .Sx
             
            Call SetGammaData(ListNo, CLR_RED)
            SetCurrentData
            If ListNo > 8 Then
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
        End If
        xyControl1.Visible = False
        picGraph.Visible = True
        DoEvents
    
    End With
    
    If lMeasureSpec = MSR_256 Then
        For lStep = 0 To lMeasureSpec - 1
            sRedGamma(lStep) = sRedGamma(lStep) / sRedGamma(lMeasureSpec - 1)
        Next lStep
    Else
        For lStep = 0 To lMeasureSpec
            sRedGamma(lStep) = sRedGamma(lStep) / sRedGamma(lMeasureSpec)
        Next lStep
    End If
    SetGraphData
    
    '=====================
    ' Measure Green
    '=====================
  '  If FormVisibleFlg = True Then
        SetWin 0, 255, 0
   ' Else
       ' objVGControl.RedSW = False
      ''  objVGControl.GreenSW = True
       ' objVGControl.BlueSW = False
       Call frmSet.comsendGrayNum(255, CLR_GREEN, rgbChange)
   ' End If
   
   
    MsgBox "Green Measure!", vbOKOnly
    xyControl1.Visible = True
    picGraph.Visible = False
    DoEvents
    
    'objCa.Measure
    
    With typCurrentMeasurementData
        For lStep = 0 To lMeasureSpec - 1
          '  If FormVisibleFlg = True Then
                SetWin 0, lLevelStep * lStep, 0
                DoEvents
                
          '  Else
              '  objVGControl.SetGVideoLevel 0, lLevelStep * lStep, 0
              Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_GREEN, rgbChange)
          '  End If
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sGreenGamma(lStep) = .Sy
            
            LabelDataVal(0).Caption = Format(.ssx, FORMAT_SXY)
            LabelDataVal(1).Caption = Format(.ssy, FORMAT_SXY)
            LabelDataVal(2).Caption = Format(.sLv, FORMAT_LV)
            DoEvents
            Call SetGammaData(ListNo, CLR_GREEN)
            If ListNo > 8 Then           'Grid Scroll
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
            
            If bStop = True Then
                Exit Sub
            End If
        Next lStep
        If lMeasureSpec <> MSR_256 Then
           ' If FormVisibleFlg = True Then
                SetWin 0, 255, 0
                DoEvents
                
           ' Else
               ' objVGControl.SetGVideoLevel 0, 255, 0
               Call frmSet.comsendGrayNum(255, CLR_GREEN, rgbChange)
          '  End If
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sGreenGamma(lMeasureSpec) = .Sy
            
            Call SetGammaData(ListNo, CLR_GREEN)
            SetCurrentData  '020311
            If ListNo > 8 Then           'Grid Scroll
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
        End If
        
        xyControl1.Visible = False
        picGraph.Visible = True
        DoEvents
        
    End With
    If lMeasureSpec = MSR_256 Then
        For lStep = 0 To lMeasureSpec - 1
            sGreenGamma(lStep) = sGreenGamma(lStep) / sGreenGamma(lMeasureSpec - 1)
        Next lStep
    Else
        For lStep = 0 To lMeasureSpec
            sGreenGamma(lStep) = sGreenGamma(lStep) / sGreenGamma(lMeasureSpec)
        Next lStep
    End If
    
    SetGraphData
    
    '=====================
    ' Measure Blue
    '=====================
   ' If FormVisibleFlg = True Then
    '    SetWin 0, 0, 255
   ' Else
      '  objVGControl.RedSW = False
      '  objVGControl.GreenSW = False
      '  objVGControl.BlueSW = True
   '   Call frmSet.comsendGrayNum(255, CLR_BLUE)
   ' End If
    SetWin 0, 0, 255
   ' If Not rgbChange Then
               
                DoEvents
                Call frmSet.comsendGrayNum(255 * 1, CLR_BLUE, rgbChange)
   '  Else
                ' SetWin 255, 0, 0
           '      DoEvents
               '  Call frmSet.comsendGrayNum(255, CLR_RED, rgbChange)
     'End If
   
    MsgBox "Blue Measure!", vbOKOnly
    xyControl1.Visible = True
    picGraph.Visible = False
    DoEvents
    
    'objCa.Measure
    
    With typCurrentMeasurementData
        For lStep = 0 To lMeasureSpec - 1
          '  If FormVisibleFlg = True Then
             '   SetWin 0, 0, lLevelStep * lStep
            '    DoEvents
                
           ' Else
               ' objVGControl.SetGVideoLevel 0, 0,
             '  Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_BLUE)
               
           ' End If
            SetWin 0, 0, lLevelStep * lStep
          ' If Not rgbChange Then
               
                DoEvents
                Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_BLUE, rgbChange)
           '  Else
                ' SetWin lLevelStep * lStep, 0, 0
             '    DoEvents
               '  Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_RED, rgbChange)
           '  End If
           
           
           
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sBlueGamma(lStep) = .Sz
            
            LabelDataVal(0).Caption = Format(.ssx, FORMAT_SXY)
            LabelDataVal(1).Caption = Format(.ssy, FORMAT_SXY)
            LabelDataVal(2).Caption = Format(.sLv, FORMAT_LV)
            DoEvents
            Call SetGammaData(ListNo, CLR_BLUE)
            If ListNo > 8 Then           'Grid Scroll
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
            
            If bStop = True Then
                Exit Sub
            End If
        Next lStep
        If lMeasureSpec <> MSR_256 Then
           ' If FormVisibleFlg = True Then
              '  SetWin 0, 0, 255
             '   DoEvents
           ' Else
              '  objVGControl.SetGVideoLevel 0, 0, 255
            '  Call frmSet.comsendGrayNum(255, CLR_BLUE)
           ' End If
           SetWin 0, 0, 255
          ' If Not rgbChange Then
              '
                DoEvents
                Call frmSet.comsendGrayNum(255 * 1, CLR_BLUE, rgbChange)
           ' Else
                 
                ' DoEvents
                ' Call frmSet.comsendGrayNum(255, CLR_RED, rgbChange)
            'End If
           
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sBlueGamma(lMeasureSpec) = .Sz
            
            Call SetGammaData(ListNo, CLR_BLUE)
            SetCurrentData  '020311
            If ListNo > 8 Then           'Grid Scroll
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
        End If
        
        xyControl1.Visible = False
        picGraph.Visible = True
        DoEvents
        
    End With
    If lMeasureSpec = MSR_256 Then
        For lStep = 0 To lMeasureSpec - 1
            sBlueGamma(lStep) = sBlueGamma(lStep) / sBlueGamma(lMeasureSpec - 1)
        Next lStep
    Else
        For lStep = 0 To lMeasureSpec
            sBlueGamma(lStep) = sBlueGamma(lStep) / sBlueGamma(lMeasureSpec)
        Next lStep
    End If
    SetGraphData
    
    '=====================
    ' Measure White
    '=====================
MeasureWhite:
    
    'If FormVisibleFlg = True Then
        SetWin 255, 255, 255
        Call frmSet.comsendGrayNum(255, CLR_WHITE, rgbChange)
        
        If ComState <> "" Then
        LabelCom.Caption = ComState
        Else: LabelCom = ""
        End If
        
   ' Else
     '   Call frmSet.comsendGrayNum(255, CLR_WHITE, rgbChange)
   ' End If
    MsgBox "White Measure!", vbOKOnly
    xyControl1.Visible = True
    picGraph.Visible = False
    DoEvents
    
    'objCa.Measure
    
    With typCurrentMeasurementData
        For lStep = 0 To lMeasureSpec - 1   '量测<255部分的阶数
        
            SetWin lLevelStep * lStep, lLevelStep * lStep, lLevelStep * lStep
            
            '    objVGControl.SetGVideoLevel lLevelStep * lStep, lLevelStep * lStep, lLevelStep * lStep
            Call frmSet.comsendGrayNum(lLevelStep * lStep, CLR_WHITE, rgbChange)
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sWhiteGamma(lStep) = .Sy
            
            LabelDataVal(0).Caption = Format(.ssx, FORMAT_SXY)
            LabelDataVal(1).Caption = Format(.ssy, FORMAT_SXY)
            LabelDataVal(2).Caption = Format(.sLv, FORMAT_LV)
            DoEvents
            Call SetGammaData(ListNo, CLR_WHITE)
            If ListNo > 8 Then
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
            
            If bStop = True Then
                Exit Sub
            End If
        Next lStep
        If lMeasureSpec <> MSR_256 Then   '针对不是256阶情况加测白画面（256阶时候不需要测）
             DoEvents
             
             '   objVGControl.SetGVideoLevel 255, 255, 255
              SetWin 255, 255, 255
              Call frmSet.comsendGrayNum(255, CLR_WHITE, rgbChange)
              
            If WaitTime = 0 Then
            Else
                SetWait
            End If
            
            objCa.Measure
            .dateColorData = Date
            .timeColorData = Time
            .lColorStatus = objProbe.RD
            .ssx = objProbe.Sx
            .ssy = objProbe.Sy
            .sLv = objProbe.Lv
            .sLvfL = objProbe.LvfL
            .Sx = objProbe.X
            .Sy = objProbe.Y
            .Sz = objProbe.Z
            .sud = objProbe.ud
            .svd = objProbe.vd
            .sduv = objProbe.duv
            .LT = objProbe.T
            sWhiteGamma(lMeasureSpec) = .Sy
            
            Call SetGammaData(ListNo, CLR_WHITE)
            SetCurrentData  '020311
            If ListNo > 8 Then
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            ListNo = ListNo + 1
        End If
        
        xyControl1.Visible = False
        picGraph.Visible = True
        DoEvents
        
    End With
    If lMeasureSpec = MSR_256 Then
        For lStep = 0 To lMeasureSpec - 1
            sWhiteGamma(lStep) = sWhiteGamma(lStep) / sWhiteGamma(lMeasureSpec - 1)
        Next lStep
    Else
        For lStep = 0 To lMeasureSpec
            sWhiteGamma(lStep) = sWhiteGamma(lStep) / sWhiteGamma(lMeasureSpec)
        Next lStep
    End If
    SetGraphData
End Sub
Public Sub InitializeData()
    Dim i As Integer
    
    For i = 0 To 255
        sRedGamma(i) = 0#
        sGreenGamma(i) = 0#
        sBlueGamma(i) = 0#
        sWhiteGamma(i) = 0#
    Next i
    
    If Option16.Value = True Then
        lMeasureSpec = MSR_16
    ElseIf Option32.Value = True Then
        lMeasureSpec = MSR_32
    ElseIf Option64.Value = True Then
        lMeasureSpec = MSR_64
    ElseIf Option128.Value = True Then
        lMeasureSpec = MSR_128
    Else
        lMeasureSpec = MSR_256
    End If

End Sub

Private Sub UpDownGraph_Change()
        If xyControl1.Visible = True Then
            xyControl1.Visible = False
            picGraph.Visible = True
        Else
            xyControl1.Visible = True
            picGraph.Visible = False
            SetGraphData
        End If
End Sub

Private Sub xlsClear_Click()
GridInit
End Sub
Private Sub SetWin(ByVal rr As Integer, gg As Integer, bb As Integer)

    If rr < 0 Or rr > 255 Then
        rr = 255
    End If
    If gg < 0 Or gg > 255 Then
        gg = 255
    End If
    If bb < 0 Or bb > 255 Then
        bb = 255
    End If

    FormGamma.TextView.BackColor = RGB(rr, gg, bb)

End Sub


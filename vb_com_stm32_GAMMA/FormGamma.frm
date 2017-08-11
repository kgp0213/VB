VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F1EB48E5-7E03-41F5-A4D0-CA86119EF992}#73.0#0"; "CaControl.ocx"
Object = "{F0971ADD-CEF2-46B3-8D7F-C075DE0316B1}#18.0#0"; "MinoltaColorSpaceControl.ocx"
Begin VB.Form FormGamma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gamma2017.8.9"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13965
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   13965
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "标准值"
      Height          =   1335
      Left            =   10800
      TabIndex        =   42
      Top             =   2280
      Width           =   2895
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   52
         Text            =   "100"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   51
         Text            =   "0.3"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   50
         Text            =   "0.3"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   49
         Text            =   "0.3"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   44
         Text            =   "100"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   43
         Text            =   "0.3"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Labeleditxyz 
         Caption         =   "Edit"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Lv:"
         Height          =   255
         Left            =   720
         TabIndex        =   47
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "y:"
         Height          =   255
         Left            =   720
         TabIndex        =   46
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "x:"
         Height          =   255
         Left            =   720
         TabIndex        =   45
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton Command_Saveas 
      Caption         =   "Save as..."
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
      Left            =   9240
      TabIndex        =   41
      ToolTipText     =   "自定义保存目录"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text_barcode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7200
      TabIndex        =   38
      Text            =   "Barcode"
      Top             =   3000
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   1080
      ScaleHeight     =   1035
      ScaleWidth      =   7275
      TabIndex        =   34
      Top             =   5040
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
         Left            =   480
         TabIndex        =   35
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
      TabIndex        =   33
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox TextView 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   31
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
      TabIndex        =   27
      Top             =   2640
      Width           =   2295
      Begin VB.CommandButton comSett 
         Caption         =   "COM"
         Height          =   375
         Left            =   120
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
   Begin VB.Frame Frame2 
      Caption         =   "设置"
      Height          =   2055
      Left            =   6960
      TabIndex        =   23
      Top             =   120
      Width           =   6855
      Begin VB.TextBox Text_standard 
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
         Left            =   2520
         TabIndex        =   39
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox CheckgrayValueText 
         Caption         =   "锁定"
         Height          =   255
         Left            =   5640
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label_readme 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   900
      End
      Begin VB.Label laberGrayinfo 
         Caption         =   "标准Barcode："
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
         TabIndex        =   24
         Top             =   240
         Width           =   4095
      End
   End
   Begin MinoltaCaControl.CaControl CaControl1 
      Height          =   555
      Left            =   6960
      TabIndex        =   21
      Top             =   2160
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
      Left            =   13200
      TabIndex        =   18
      Text            =   "5"
      Top             =   5160
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
      Left            =   11640
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
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
         TabIndex        =   17
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
      Left            =   9000
      TabIndex        =   1
      Top             =   3840
      Width           =   2175
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
      ToolTipText     =   "默认数据以时间命名保存于当前目录"
      Top             =   5040
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmdDiag 
      Left            =   13320
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton xlsClear 
      Caption         =   "清除"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9240
      TabIndex        =   26
      Top             =   6240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdDataList 
      Height          =   2715
      Left            =   60
      TabIndex        =   4
      Top             =   4080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   10401
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
      TabIndex        =   22
      Top             =   -120
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   7541
   End
   Begin VB.Label Label2 
      Caption         =   "当前Barcode："
      Height          =   255
      Left            =   6960
      TabIndex        =   40
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "YJ2017"
      Height          =   375
      Left            =   13200
      TabIndex        =   37
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
      TabIndex        =   36
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
      Left            =   13200
      TabIndex        =   20
      Top             =   5520
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
      Left            =   13200
      TabIndex        =   19
      Top             =   4920
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

'Dim typMeasurementData(1040) As TypeMeasurementData
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
Dim SaveasFlag As Boolean
Dim Mydata(300, 4) As Single
Dim sRedGamma(256) As Single
Dim sGreenGamma(256) As Single
Dim sBlueGamma(256) As Single
Dim sWhiteGamma(256) As Single
Dim lMeasureSpec As Long
Dim lSelectedRow0 As Long
Dim lSelectedRow1 As Long
Dim connectCA310ok As Boolean
Dim n As Integer
Dim WaitTime As Double
Const xmin = 0.27
Const xmax = 0.33
Const ymin = 0.29
Const ymax = 0.35
Const lvmin = 280
Const lvmax = 500
Dim editflag As Integer

Sub SetGraphData()
    Dim i As Integer, j As Integer
    
    'If lMeasureSpec = MSR_256 Then
   '     j = lMeasureSpec - 1
   ' Else
        j = lMeasureSpec
    'End If
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
    grdDataList.Col = 1: grdDataList.Text = "Barcode"
    grdDataList.Col = 2: grdDataList.Text = "reserve1"
    grdDataList.Col = 3: grdDataList.Text = "reserve2"
    grdDataList.Col = 4: grdDataList.Text = "x"
    grdDataList.Col = 5: grdDataList.Text = "y"
    
    ' 021225
    ' grdDataList.Col = 6: grdDataList.Text = "Lv"
    grdDataList.Col = 6: grdDataList.Text = gstrLvOrEv
    
    'grdDataList.Col = 7: grdDataList.Text = "ud"
    'grdDataList.Col = 8: grdDataList.Text = "vd"
    'grdDataList.Col = 9: grdDataList.Text = "T"
    'grdDataList.Col = 10: grdDataList.Text = "duv"
    grdDataList.Col = 8: grdDataList.Text = "Date"
    grdDataList.Col = 9: grdDataList.Text = "Time"
    
    grdDataList.ColWidth(0) = 420 '380   'No
    
    grdDataList.ColWidth(1) = 1000   'Barcode
    grdDataList.ColWidth(2) = 400   'reserve1
    grdDataList.ColWidth(3) = 400   'reserve2
    
    grdDataList.ColWidth(4) = 600   'x
    grdDataList.ColWidth(5) = 600   'y
    grdDataList.ColWidth(6) = 950   'Lv
    
    grdDataList.ColWidth(7) = 400   '--
    grdDataList.ColWidth(8) = 1050  'Date
    grdDataList.ColWidth(9) = 1000   'Time
    
    grdDataList.Col = 0
    For i = 1 To 10400
        grdDataList.Row = i
        grdDataList.Text = Format(i)
    Next i
   
    ListNo = 1

    grdDataList.TopRow = 1

End Sub
Sub SaveData()

    Dim dd(5040, 10) As String   '
    Dim i As Integer, j As Integer
    Dim fm As String, fnum As Integer, fname As String

    For i = 1 To ListNo - 1
        grdDataList.Row = i
        For j = 1 To 10
            grdDataList.Col = j
            dd(i, j) = grdDataList.Text
        Next j
    Next i
    
    If (SaveasFlag) Then
    
        On Error Resume Next
        cmdDiag.CancelError = True
         cmdDiag.FileName = ""
         cmdDiag.Filter = "Data Files (*.csv)|*.csv"
        cmdDiag.FilterIndex = 2
        cmdDiag.Action = 2
     
         If Err.Number = cdlCancel Then
             Exit Sub
          Else
        'fm = App.Path + "\" + Format(Now(), "yyyy-MM-dd") + Format(Now(), "-HHmmss") + "gamma.csv"
             fm = cmdDiag.FileName
             fm = Mid$(fm, 1, InStr(1, fm, ".")) + "csv"
          End If
       Else
       fm = App.Path + "\" + Format(Now(), "yyyy-MM-dd") + Format(Now(), "-HHmmss") + "gamma.csv"
            ' fm = cmdDiag.FileName
             'fm = Mid$(fm, 1, InStr(1, fm, ".")) + "csv"
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
    Write #fnum, "No.", "Barcode", "--", "--", "x", "y", gstrLvOrEv, "-", "Date", "Time"
    
    For i = 1 To ListNo - 1
        Print #fnum, Format(i, "000"); ",";
        For j = 1 To 9
            Print #fnum, dd(i, j); ",";
        Next j
        Print #fnum, dd(i, 10)
    Next i
    
    
    Close fnum
    SaveasFlag = False
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


Private Sub CheckgrayValueText_Click()
'锁定标准barcode输入框
Text_standard.Locked = Not Text_standard.Locked
    If (Text_standard.Locked = False) Then
    Text_standard.SetFocus
    End If

End Sub

Private Sub Command_Saveas_Click()
SaveasFlag = True
SaveData
End Sub


Private Sub Labeleditxyz_dblClick()
'Labeleditxyz.Caption = "Edit"
editflag = editflag + 1
If (editflag = 2) Then
Text1(0).Locked = False
Text1(1).Locked = False
Text2(0).Locked = False
Text2(1).Locked = False
Text3(0).Locked = False
Text3(1).Locked = False
Labeleditxyz.Caption = "Save"
Else
    If (editflag = 3) Then
    Text1(0).Locked = True
    Text1(1).Locked = True
    Text2(0).Locked = True
    Text2(1).Locked = True
    Text3(0).Locked = True
    Text3(1).Locked = True
    Labeleditxyz.Visible = False
    Frame3.Caption = "标准已改:"
    End If
End If
End Sub

Private Sub Text_standard_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Len(Text_standard.Text) > 6) Then   'MsgBox "标准Barcode已录入"

n = Len(Text_standard.Text) '计算标准barcode长度，供后续直接调用
   
 CheckgrayValueText.Visible = True
 laberGrayinfo.Caption = "BC式样锁定："
 MsgBox "标准Barcode已录入"
 CheckgrayValueText.Visible = False    '隐藏锁定按键
 Text_standard.Locked = True           '锁定当前输入框
 Text_barcode.Enabled = True
 Text_barcode.SetFocus
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
'CommandmanualMeasure.Enabled = False
TextView.Text = "确认数据保存后再关闭程序"
End Sub

Public Sub CommandConnect_Click()
'Me.Hide
'FormCa310Connect.Show vbModal

'FormCa310Connect.Label1.Caption = "正在连接CA310, 请等待联机完成..."
'Picture1.Visible = True   ' 提示： 正在连接CA310, 请等待联机完成..."
cmdBusyFlag = True
Label3.Caption = "CA310联机中，请静候连接完成..."
DoEvents
'-------------------------
StartMain
CommandConnect.BackColor = &H8000000F
'CommandmanualMeasure.Enabled = False
Command0Cal.Enabled = True
CommandCloseConnect.Enabled = True
'Unload FormCa310Connect
Picture1.Visible = False   ' 结束提示： 正在连接CA310, 请等待联机完成..."
'CommandMeasure.Enabled = True
cmdBusyFlag = False
CommandSave.Enabled = True
connectCA310ok = True
CheckgrayValueText.Enabled = True
CheckgrayValueText.SetFocus
'Text_barcode.Enabled = True
'FormGamma.Text_barcode.SetFocus
End Sub

Private Sub CommandMeasure_Click()
    Dim bResult As Boolean
    'CMD1Text.FontSize = 12
    
    'gammaTestMode.Enabled = False
    ' 030407
    If objCa.DisplayMode <> COLOR_MODE Then
        objCa.DisplayMode = DSP_LXY
    End If
    
    cmdBusyFlag = True
    
   ' cmdWin.Enabled = False
    bStop = False
    CommandMeasure.Enabled = False
    CommandStop.Enabled = True
    
    DoEvents
    
    MeasureGamma
    
    CommandStop.Enabled = False
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
Private Sub FormGamma_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox KeyCode
End Sub

Private Sub Form_Activate()

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
Text1(0).Text = Format(xmin, "0.0000")
Text1(1).Text = Format(xmax, "0.0000")
Text2(0).Text = Format(ymin, "0.0000")
Text2(1).Text = Format(ymax, "0.0000")
Text3(0).Text = lvmin
Text3(1).Text = lvmax
Text1(0).Locked = True
Text1(1).Locked = True
Text2(0).Locked = True
Text2(1).Locked = True
Text3(0).Locked = True
Text3(1).Locked = True
editflag = 0
dubuggFlag = 0
SaveasFlag = False
m = 0
'Coln = 0
grayNumflag = 0
cmdBusyFlag = True
'Call grayValueText_Change
'Unload frmSet
'MeasureGamma
CheckgrayValueText.Enabled = False
FrameRefData.Enabled = False
'Label4.Caption = ""
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

'Label2.Caption = "上述数据以0开始、以255结尾，数据间以逗号分隔"
'Me.AcceptButton = Me.CommandCconnect
'Me.CommandConnect.SetFocus
CommandConnect.BackColor = vbRed
Command0Cal.Enabled = False
CommandMeasure.Enabled = False
CheckgrayValueText.Value = Checked
'CommandmanualMeasure.Enabled = False
CommandCloseConnect.Enabled = False
'Picture1.Visible = False
' ColorRgb2Bgr.Enabled = False
 rgbChange = False
Label_readme = "首先(连接CA310)，然后：" & vbCrLf & "1, 解锁并输入当前批次任一Barcode作为标准后再次锁定" & vbCrLf & _
                "2, 确认LCD显示后，放置好CA310探头" & vbCrLf & _
                "3, 扫描产品Barcode之后，会自动进行测量"
Label5.Caption = App.Major & "." & App.Minor & "." & App.Revision
GridInit
CommandConnect.Enabled = True
'CommandConnect.SetFocus
Text_barcode.Text = ""
CommandSave.Enabled = False
Text_barcode.Enabled = False
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
   SaveData  '直接退出自动保存数据
   MsgBox "数据已经以时间命名保存于当前程序目录", vbOKOnly  '给出保存提示
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
'



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

Sub SetGammaData(ByVal LisNo As Integer, RGB As String)
    
    grdDataList.Row = LisNo
    
    grdDataList.Col = 1
    grdDataList.Text = Text_barcode.Text
    grdDataList.Col = 2
    grdDataList.Text = RGB
   ' grdDataList.Col = 3
   ' grdDataList.Text = Format(typCurrentMeasurementData.Sz, "0.00")
    grdDataList.Col = 4
    grdDataList.Text = Format(typCurrentMeasurementData.ssx, FORMAT_SXY)
    grdDataList.Col = 5
    grdDataList.Text = Format(typCurrentMeasurementData.ssy, FORMAT_SXY)
    grdDataList.Col = 6
    grdDataList.Text = Format(typCurrentMeasurementData.sLv, FORMAT_LV)
   ' grdDataList.Col = 7
   ' grdDataList.Text = Format(typCurrentMeasurementData.sud, FORMAT_SXY)
   ' grdDataList.Col = 8
   ' grdDataList.Text = Format(typCurrentMeasurementData.svd, FORMAT_SXY)
   ' grdDataList.Col = 9
   ' If typCurrentMeasurementData.LT = -1 Then
   '     grdDataList.Text = "-"
   ' Else
   '     grdDataList.Text = Format(typCurrentMeasurementData.LT, "00000")
   ' End If
   ' grdDataList.Col = 10
   ' If typCurrentMeasurementData.LT = -1 Then
   '     grdDataList.Text = "-"
   ' Else
   '     grdDataList.Text = Format(typCurrentMeasurementData.sduv, "+.000;-.000")
   ' End If
    grdDataList.Col = 8
    grdDataList.Text = Format(typCurrentMeasurementData.dateColorData, "yyyy/mm/dd")
    grdDataList.Col = 9
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
   ' If lMeasureSpec = MSR_16 Then
    '    MaxDataNo = MSR_16 + 1      '17 data
   

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

Private Sub MeasureGamma()

    Dim lVLocation As Long
    Dim lHLocation As Long
    Dim lVLocationMax As Long
    Dim lHLocationMax As Long
    Dim strMsg As String
    
    
    On Error Resume Next
    
    WaitTime = Val(TextIntervalSec.Text) / 10#
   
    InitializeData
   ' GridInit
    
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
    
   
    '=====================
    ' Measure Start
    '=====================
'MeasureWhite:
  
        If ComState <> "" Then
        LabelCom.Caption = ComState
        Else: LabelCom = ""
        End If
        
        xyControl1.Visible = True
        picGraph.Visible = False
        DoEvents
        '-----------显示画面开始---------------------------
        Call frmSet.comsendGrayNum(255, CLR_WHITE, rgbChange)    '参数1：阶数（0～2255）；参数2：全色画面（CLR_RED：全红，CLR_GREEN:全绿，CLR_BLUE:全蓝，CLR_WHITE:全灰）；参数3：rgb是否交换
        SetWin 255, 255, 255
        xpartmeasure "W"
        Call frmSet.comsendGrayNum(255, CLR_RED, rgbChange)
        SetWin 255, 0, 0
        xpartmeasure "R"
        Call frmSet.comsendGrayNum(255, CLR_GREEN, rgbChange)
        SetWin 0, 255, 0
        xpartmeasure ("G")
        Call frmSet.comsendGrayNum(255, CLR_BLUE, rgbChange)
        SetWin 0, 0, 255
        xpartmeasure "B"
        Call frmSet.comsendGrayNum(0, CLR_WHITE, rgbChange)
        SetWin 0, 0, 0
        xpartmeasure "D"
        
        MsgBox "量测结束，更换产品", vbOKOnly
        SetWin 255, 255, 255
      '-------------测试结束---------------------------------
      CommandMeasure.Enabled = False
      Text_barcode.Text = ""  '清空Barcode栏
      Text_barcode.SetFocus  ' 设定焦点于Barcode栏，等待输入
   
   
End Sub
Private Sub xpartmeasure(ByVal rgbname As String)

    With typCurrentMeasurementData
       
            If WaitTime = 0 Then
            Else
                SetWait
            End If              '设定延时
            
            objCa.Measure       '量测数据
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
            'sWhiteGamma(lStep) = .Sy
            
            LabelDataVal(0).Caption = Format(.ssx, FORMAT_SXY)   '当前 x，y Lv显示
            LabelDataVal(1).Caption = Format(.ssy, FORMAT_SXY)
            LabelDataVal(2).Caption = Format(.sLv, FORMAT_LV)
            DoEvents
            
            '----------添加x y Lv 数值范围判断-------
            
            
            
            '----------------------------------------
            Call SetGammaData(ListNo, rgbname)      '  xyLv数据保存于csv窗口
            If ListNo > 8 Then                                   '设定数据窗口内容超过目标行数后自动向上滚动
                grdDataList.TopRow = grdDataList.TopRow + 1
            End If
            'Coln = Coln + 1
            'If (Coln = 4) Then    'Coln 代表画面数量，一个循环里面所有画面数量的统计（放在一行里面）
            ListNo = ListNo + 1
            '‘Coln = 4
            'End If
            
            If bStop = True Then        '设定停止标志位，以响应“Stop”按钮
                Exit Sub
            End If
       
        xyControl1.Visible = True
        'picGraph.Visible = True
        DoEvents
        
    End With
    'SetGraphData
End Sub
'Private Sub Text_barcode_Change()
'If (Len(Text_barcode.Text) = Len(Text_standard.Text)) And Len(Text_standard.Text) > 5 And (connectCA310ok = True) Then
'CheckgrayValueText.Enabled = False
'CommandMeasure.Enabled = True
'CommandMeasure.SetFocus
'Exit Sub
'End If
'End Sub
Private Sub Text_barcode_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        If (Len(Text_barcode.Text) >= Len(Text_standard.Text)) Then   '
           If (Len(Text_barcode.Text) = Len(Text_standard.Text)) Then
           Else
              Text_barcode.Text = Mid(Text_barcode.Text, Len(Text_barcode.Text) - n + 1, n)
              'n为标准的barcode长度
           End If
        Else
            MsgBox "Barcode异常，请确认！"
            Text_barcode.SetFocus
            Exit Sub
        End If
'CheckgrayValueText.Enabled = False  '
CommandMeasure.Enabled = True
'CommandMeasure.SetFocus
Call CommandMeasure_Click
End If
End Sub

Public Sub InitializeData()
    Dim i As Integer
    
    For i = 0 To 255
        sRedGamma(i) = 0#
        sGreenGamma(i) = 0#
        sBlueGamma(i) = 0#
        sWhiteGamma(i) = 0#
    Next i
  
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


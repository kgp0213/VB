VERSION 5.00
Object = "{1B773E42-2509-11CF-942F-008029004347}#3.3#0"; "sysmon.ocx"
Begin VB.Form Form1 
   Caption         =   "系统监视器"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   6900
   StartUpPosition =   3  '窗口缺省
   Begin SystemMonitorCtl.SystemMonitor SystemMonitor1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _Version        =   196611
      _ExtentX        =   11880
      _ExtentY        =   11245
      DisplayType     =   1
      ReportValueType =   0
      MaximumScale    =   200
      MinimumScale    =   50
      ShowLegend      =   -1  'True
      ShowToolbar     =   -1  'True
      ShowScaleLabels =   -1  'True
      ShowHorizontalGrid=   0   'False
      ShowVerticalGrid=   0   'False
      ShowValueBar    =   -1  'True
      ManualUpdate    =   0   'False
      Highlight       =   0   'False
      ReadOnly        =   0   'False
      MonitorDuplicateInstances=   -1  'True
      UpdateInterval  =   1
      BackColorCtl    =   -2147483633
      ForeColor       =   -1
      BackColor       =   -1
      GridColor       =   8421504
      TimeBarColor    =   255
      Appearance      =   -1
      BorderStyle     =   0
      GraphTitle      =   ""
      YAxisLabel      =   ""
      LogFileName     =   ""
      AmbientFont     =   -1  'True
      LegendColumnWidths=   ".103529411764706	.103529411764706	.131764705882353	.103529411764706	.103529411764706	.103529411764706	.131764705882353"
      LegendSortDirection=   0
      LegendSortColumn=   0
      CounterCount    =   0
      MaximumSamples  =   100
      SampleCount     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SysInfo1_ConfigChangeCancelled()

End Sub

Private Sub Form_Load()
    With SystemMonitor1
        .UpdateInterval = 2
        .ReportValueType = 2
        .ShowScaleLabels = False
        .MaximumScale = 200
        .ShowValueBar = True
        .ShowHorizontalGrid = True
        .ShowVerticalGrid = False
        .DisplayType = 1
    End With
End Sub

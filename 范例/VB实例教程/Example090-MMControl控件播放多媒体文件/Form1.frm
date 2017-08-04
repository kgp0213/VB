VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "MMControl"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5130
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command_Play 
      Caption         =   "Play"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command_Open 
      Caption         =   "Open"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin MCI.MMControl MMControl1 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      _Version        =   393216
      UpdateInterval  =   100
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Filename As String

Private Sub Command_Open_Click()
    Me.CommonDialog1.ShowOpen
    m_Filename = Me.CommonDialog1.FileName
    If m_Filename <> "" Then
        Me.MMControl1.FileName = m_Filename
        Me.Caption = m_Filename
        Me.MMControl1.Command = "Open"
    End If
End Sub

Private Sub Command_Play_Click()
    Me.MMControl1.Command = "Play"
End Sub

Private Sub Form_Load()
    m_Filename = ""
    Me.MMControl1.AutoEnable = True
    Me.MMControl1.hWndDisplay = Me.hWnd
End Sub

Private Sub MMControl1_StatusUpdate()
    Me.HScroll1.Max = Me.MMControl1.Length
    Me.HScroll1.Min = 0
    Me.HScroll1.Value = Me.MMControl1.Position
End Sub

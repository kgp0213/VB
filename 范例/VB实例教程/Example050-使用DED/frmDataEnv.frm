VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmDataEnv 
   Caption         =   "员工统计"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4920
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   4440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      DataField       =   "员工人数："
      DataMember      =   "CmdNWind_分组"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   660
   End
   Begin VB.TextBox txtCity 
      DataField       =   "City"
      DataMember      =   "CmdNWind_分组"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   2475
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmDataEnv.frx":0000
      Height          =   2640
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   4657
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      DataMember      =   "CmdNWind"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "员工人数：:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Adodc1.Caption = "员工统计"
    Set Adodc1.Recordset = DataEnvironment1.rsCmdNWind_分组
End Sub


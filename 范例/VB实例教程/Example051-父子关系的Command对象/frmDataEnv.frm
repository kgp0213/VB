VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmDataEnv 
   Caption         =   "父子关系的Command对象"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6195
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   1650
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   485
      Width           =   3300
   End
   Begin VB.TextBox txtEmployeeID 
      DataField       =   "EmployeeID"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   660
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   5040
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmDataEnv.frx":0000
      Height          =   3120
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   5503
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      DataMember      =   "CmdOrders"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FirstName:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LastName:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EmployeeID:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Adodc1.Caption = "移动记录指针"
    Set Adodc1.Recordset = DataEnvironment1.rsCmdNWind
End Sub


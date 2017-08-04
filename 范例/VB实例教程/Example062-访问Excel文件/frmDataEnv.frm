VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmDataEnv 
   Caption         =   "访问Excel文件"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   6990
   StartUpPosition =   3  '窗口缺省
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示记录"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim adoCnn As New ADODB.Connection
    Dim adoRst As New ADODB.Recordset
    adoCnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=c:\auto.xls;Extended Properties='Excel 8.0;HDR=Yes'"
    adoRst.Open "Select * From [employees$]", adoCnn, adOpenKeyset, adLockOptimistic
    Set Me.MSHFlexGrid1.DataSource = adoRst
End Sub


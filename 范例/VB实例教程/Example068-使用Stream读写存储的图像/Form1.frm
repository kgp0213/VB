VERSION 5.00
Begin VB.Form frmDataEnv 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   4725
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      DataField       =   "Photo"
      DataSource      =   "Adodc1"
      Height          =   2535
      Left            =   360
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim mstream As ADODB.Stream

Private Sub Form_Load()
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\db1.mdb;"
    MsgBox cn.ConnectionString
    Set rs = New ADODB.Recordset
    rs.Open "Select * from bmp±í", cn, adOpenKeyset, adLockOptimistic
    
    Set mstream = New ADODB.Stream
    mstream.Type = adTypeBinary
    mstream.Open
    mstream.LoadFromFile App.Path + "\test.bmp"
    rs.AddNew
    rs.Fields("bmp").Value = mstream.Read
    rs.Update
    rs.Close
    cn.Close
    
    'Set mstream = New ADODB.Stream
    'mstream.Type = adTypeBinary
    'mstream.Open
    'mstream.Position = 0
    'mstream.Write rs.Fields("bmp").Value
    'mstream.SaveToFile "d:\copy of test.bmp", adSaveCreateOverWrite
    'rs.Close
    'cn.Close
End Sub

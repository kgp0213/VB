VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim MyEngine As Object
    Dim MyWS As Object
    Dim myDB As Object
    Dim myRS As Object
    Set MyEngine = CreateObject("DAO.DBEngine.36")
    Set MyWS = MyEngine.Workspaces(0)
    Set myDB = MyWS.OpenDatabase("D:\Microsoft Visual Studio\VB98\BIBLIO.MDB")
    Set myRS = myDB.OpenRecordset("Title Author")
    myRS.MoveLast
    myRS.MoveFirst
    Debug.Print myRS.RecordCount
    myRS.Close
    myDB.Close
    Set myRS = Nothing
    Set myDB = Nothing
End Sub

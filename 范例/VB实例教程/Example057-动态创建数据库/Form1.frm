VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6435
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim myDB As DAO.Database
    Set myDB = DAO.Workspaces(0).CreateDatabase("d:\mydb.mdb", dbLangGeneral)
    Dim str_SQL As String
    str_SQL = "Create Table NewTable1(Field1 Text(10),Field2 Short)"
    myDB.Execute str_SQL
        
    str_SQL = "Create Table NewTable2(Field1 Text(10),Field2 Short)"
    myDB.Execute str_SQL
    myDB.Close
End Sub

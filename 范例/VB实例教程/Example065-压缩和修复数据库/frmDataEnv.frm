VERSION 5.00
Begin VB.Form frmDataEnv 
   Caption         =   "Ñ¹ËõºÍÐÞ¸´Êý¾Ý¿â"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err_Handle
    Dim dbE As New DAO.DBEngine
    dbE.CompactDatabase "d:\biblio.mdb", "c:\auto.mdb"
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub
End Sub

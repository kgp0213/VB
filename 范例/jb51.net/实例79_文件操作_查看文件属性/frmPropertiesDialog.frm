VERSION 5.00
Begin VB.Form frmPropertiesDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Displaying the Properties Dialog Box"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmPropertiesDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "C:\Autoexec.bat"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Properties..."
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   285
   End
End
Attribute VB_Name = "frmPropertiesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
  Dim r As Long
  Dim FileName As String
   
  ' Get the file name and path from txtFileName
  FileName = (txtFileName)
   
  ' Show the properties dialog, passing the filename
  ' and the owner of the dialog
  r = ShowProperties(FileName, Me.hwnd)
     
  ' Display an error if the properties dialog could
  ' not be displayed
  If r <= 32 Then MsgBox "Error"
  
End Sub



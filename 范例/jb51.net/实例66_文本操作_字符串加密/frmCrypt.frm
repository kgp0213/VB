VERSION 5.00
Begin VB.Form frmCrypt 
   Caption         =   "Encrypt / Decrypt"
   ClientHeight    =   3525
   ClientLeft      =   3240
   ClientTop       =   2700
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   4155
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtText 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Text to Encrypt / Decrypt"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDecrypt_Click()
  txtText.Text = DeCrypt(txtText.Text, txtPassword.Text)
End Sub

Private Sub cmdEncrypt_Click()
  txtText.Text = Crypt(txtText.Text, txtPassword.Text)
End Sub

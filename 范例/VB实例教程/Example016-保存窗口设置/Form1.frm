VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "±£¥Ê¥∞ø⁄…Ë÷√"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5355
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   480
      Picture         =   "Form1.frx":16102
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   480
      Picture         =   "Form1.frx":2C204
      Stretch         =   -1  'True
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempstr As String
Private Sub Form_Load()
    Me.Width = GetSetting(App.Title, Me.Name, "Width", 7200)
    Me.Height = GetSetting(App.Title, Me.Name, "Height", 6300)
    Me.Top = GetSetting(App.Title, Me.Name, "Top", 100)
    Me.Left = GetSetting(App.Title, Me.Name, "Left", 100)
    tempstr = GetSetting(App.Title, Me.Name, "Picture", "\1.bmp")
    Me.Picture = LoadPicture(App.Path + tempstr)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting(App.Title, Me.Name, "Width", Me.Width)
    Call SaveSetting(App.Title, Me.Name, "Height", Me.Height)
    Call SaveSetting(App.Title, Me.Name, "Top", Me.Top)
    Call SaveSetting(App.Title, Me.Name, "Left", Me.Left)
    Call SaveSetting(App.Title, Me.Name, "Picture", tempstr)
End Sub

Private Sub Image1_Click()
    Me.Picture = Image1.Picture
    tempstr = "\" + "1.bmp"
End Sub

Private Sub Image2_Click()
    Me.Picture = Image2.Picture
    tempstr = "\" + "2.bmp"
End Sub

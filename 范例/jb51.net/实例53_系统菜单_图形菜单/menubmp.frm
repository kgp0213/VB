VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menubmp"
   ClientHeight    =   3405
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   4980
   Begin VB.TextBox Text1 
      Height          =   2475
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "menubmp.frx":0000
      Top             =   180
      Width           =   4815
   End
   Begin VB.Image imCopy 
      Height          =   195
      Left            =   2400
      Picture         =   "menubmp.frx":02B3
      Top             =   2760
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imPrintSetup 
      Height          =   225
      Left            =   2040
      Picture         =   "menubmp.frx":07C5
      Top             =   2760
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imPrint 
      Height          =   210
      Left            =   1560
      Picture         =   "menubmp.frx":0CBB
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imSave 
      Height          =   210
      Left            =   1080
      Picture         =   "menubmp.frx":11DD
      Top             =   2760
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imOpen 
      Height          =   195
      Left            =   600
      Picture         =   "menubmp.frx":16FF
      Top             =   2760
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Print &Setup"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuExtra 
         Caption         =   "&Extra Sub Menu"
         Begin VB.Menu mnuCopy 
            Caption         =   "&Copy"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" _
   (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Const MF_BYPOSITION = &H400&


Private Sub Form_Load()
    Dim mHandle As Long, lRet As Long, sHandle As Long, sHandle2 As Long
    mHandle = GetMenu(hwnd)
    sHandle = GetSubMenu(mHandle, 0)
    lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, imOpen.Picture, imOpen.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 1, MF_BYPOSITION, imSave.Picture, imSave.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 3, MF_BYPOSITION, imPrint.Picture, imPrint.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 4, MF_BYPOSITION, imPrintSetup.Picture, imPrintSetup.Picture)
    sHandle = GetSubMenu(mHandle, 1)
    sHandle2 = GetSubMenu(sHandle, 0)
    lRet = SetMenuItemBitmaps(sHandle2, 0, MF_BYPOSITION, imCopy.Picture, imCopy.Picture)
End Sub

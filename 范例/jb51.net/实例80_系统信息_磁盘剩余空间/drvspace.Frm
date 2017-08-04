VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Space"
   ClientHeight    =   1545
   ClientLeft      =   3045
   ClientTop       =   2820
   ClientWidth     =   1965
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   1965
   Begin VB.CommandButton Command1 
      Caption         =   "Get Space"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Drive Letter"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public Function DiskSpace(DrivePath As String) As Double
' Pass the function the drive letter to get the free space of
  Dim Drive As String
  Dim SectorsPerCluster As Long, BytesPerSector As Long
  Dim NumberOfFreeClusters As Long, TotalClusters As Long, Sts As Long
  Dim DS

  Drive = Left(Trim(DrivePath), 1) & ":\"     ' Ensure path is at the root.
  Sts = GetDiskFreeSpace(Drive, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalClusters)
  If Sts <> 0 Then
    DiskSpace = SectorsPerCluster * BytesPerSector * NumberOfFreeClusters
    DS = Format$(DiskSpace, "###,###")
    Label1 = DS & " bytes"
  Else
    DiskSpace = -1        ' Should Call GetLastError here but -1 will do for example
  End If
End Function

Private Sub Command1_Click()

Dim x

If Text1 = "" Then
    MsgBox "Try typing a drive letter. It works better!"
Else
    x = DiskSpace(Text1.Text)
    Text1.SetFocus
End If

End Sub


Private Sub Form_Load()
    Show
    Text1.SetFocus
End Sub



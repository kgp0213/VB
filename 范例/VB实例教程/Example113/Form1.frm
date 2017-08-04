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
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Private Sub Form_Load()
        Dim fso As New Scripting.FileSystemObject
        Dim drive As Scripting.drive
        
        For Each drive In fso.Drives
            On Error Resume Next
            Debug.Print ("分配给驱动器的字母:" + drive.DriveLetter)
            Select Case drive.DriveType
                Case DRIVE_REMOVABLE:
                    Debug.Print ("驱动器的类型:" + "REMOVABLE")
                Case DRIVE_FIXED:
                    Debug.Print ("驱动器的类型:" + "FIXED")
                Case DRIVE_REMOTE:
                    Debug.Print ("驱动器的类型:" + "REMOTE")
                Case DRIVE_CDROM:
                    Debug.Print ("驱动器的类型:" + "CDROM")
                Case DRIVE_RAMDISK:
                    Debug.Print ("驱动器的类型:" + "RAMDISK")
            End Select
            Debug.Print ("驱动器的卷名:" + drive.VolumeName)
            Debug.Print ("---------------------------------------")
        Next
End Sub

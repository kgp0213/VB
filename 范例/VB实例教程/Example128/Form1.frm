VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Internet"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   2640
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command_Detect 
      Caption         =   "Detect"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
                (ByRef lpdwFlags As Long, _
                ByVal dwReserved As Long) _
                As Boolean

Private Enum Flags
   'Local system uses a LAN to connect to the Internet.
   INTERNET_CONNECTION_LAN = &H2
   'Local system uses a modem to connect to the Internet.
   INTERNET_CONNECTION_MODEM = &H1
   'Local system uses a proxy server to connect to the Internet.
   INTERNET_CONNECTION_PROXY = &H4
   'Local system has RAS installed.
   INTERNET_RAS_INSTALLED = &H10
End Enum

Private Sub Command_Detect_Click()
    Dim lngFlags As Long
    
    If InternetGetConnectedState(lngFlags, 0) Then
    'connected.
        If lngFlags And Flags.INTERNET_CONNECTION_LAN Then
            'LAN connection.
             MsgBox ("LAN connection.")
        ElseIf lngFlags And Flags.INTERNET_CONNECTION_MODEM Then
            'Modem connection.
             MsgBox ("Modem connection.")
        ElseIf lngFlags And Flags.INTERNET_CONNECTION_PROXY Then
            'Proxy connection.
             MsgBox ("Proxy connection.")
        End If
    Else
        'not connected.
         MsgBox ("Not connected.")
    End If
End Sub

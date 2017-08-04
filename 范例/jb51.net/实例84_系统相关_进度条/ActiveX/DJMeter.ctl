VERSION 5.00
Begin VB.UserControl DJMeter 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   ScaleHeight     =   570
   ScaleWidth      =   1740
   Begin VB.PictureBox picMeter 
      Align           =   2  'Align Bottom
      ClipControls    =   0   'False
      Height          =   240
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   1680
      TabIndex        =   1
      Top             =   330
      Width           =   1740
      Begin VB.Shape shpMeter 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   60
      Width           =   75
   End
End
Attribute VB_Name = "DJMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Const conMessageHeight = 0.5
Dim mlngPercent As Long
Const conDefaultPercent = 100
'Default Property Values:
Const m_def_BackColor = 0
'Property Variables:
Dim m_BackColor As OLE_COLOR

Public Event Click()
Attribute Click.VB_Description = "click meter event"
Public Event Change()
Attribute Change.VB_Description = "change meter event"



Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/returns meter caption"
    Caption = lblMessage.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    lblMessage.Caption = NewCaption
    PropertyChanged "Caption"
End Property

Private Sub SetPercent()
    shpMeter.Width = picMeter.Width * Me.Percent / 100
    RaiseEvent Change
End Sub

Property Get Percent() As Long
Attribute Percent.VB_Description = "Sets/returns pecentage of meter filled."
    Percent = mlngPercent
End Property

Property Let Percent(ByVal NewPercent As Long)
    If NewPercent <= 100 Then
        mlngPercent = NewPercent
        Call SetPercent
        
        PropertyChanged "Percent"
    Else
        Err.Raise vbObjectError + 1111, _
         "Meter::Percent (Let)", _
         "Percent must be between 0 and 100."
    End If
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Sets/returns font of caption"
Attribute Font.VB_UserMemId = -512
    Set Font = lblMessage.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set lblMessage.Font = NewFont
    PropertyChanged "Font"
End Property
'
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = lblMessage.BackColor
'End Property
'
'Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
'    lblMessage.BackColor = NewBackColor
'    PropertyChanged "BackColor"
'End Property

Private Sub UserControl_Resize()
    ' Set the width of the label control.
    ' Set the height to the chosen ratio of the
    ' control's height.
    lblMessage.Move 0, 0, _
     UserControl.ScaleWidth, _
     UserControl.ScaleHeight * conMessageHeight
    picMeter.Move 0, lblMessage.Height, _
     lblMessage.Width, _
     UserControl.ScaleHeight * (1 - conMessageHeight)
    shpMeter.Move 0, 0, shpMeter.Width, picMeter.Height
End Sub

Private Sub UserControl_InitProperties()
    Me.Percent = conDefaultPercent
    Me.Caption = Extender.Name
    Me.BackColor = Ambient.BackColor
    Set Me.Font = Ambient.Font
    Debug.Print "InitProperties"
    m_BackColor = m_def_BackColor
End Sub
Private Sub UserControl_WriteProperties( _
 PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", _
     lblMessage.Caption, "")
    Call PropBag.WriteProperty("Percent", _
     mlngPercent, conDefaultPercent)
    Call PropBag.WriteProperty("BackColor", _
     lblMessage.BackColor, vbButtonText)
    Call PropBag.WriteProperty("Font", _
     Font, Ambient.Font)
    Debug.Print "WriteProperties"
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("FillColor", shpMeter.FillColor, &HFF&)
End Sub
Private Sub UserControl_ReadProperties( _
 PropBag As PropertyBag)
    lblMessage.Caption = PropBag.ReadProperty( _
     "Caption", lblMessage.Caption)
    Set Font = PropBag.ReadProperty( _
     "Font", Ambient.Font)
    shpMeter.FillColor = PropBag.ReadProperty( _
    "FillColor", shpMeter.FillColor)
    lblMessage.BackColor = PropBag.ReadProperty( _
     "BackColor", lblMessage.BackColor)
    mlngPercent = PropBag.ReadProperty( _
     "Percent", conDefaultPercent)
    ' Don't forget to set the width of the meter.
    Call SetPercent
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    shpMeter.FillColor = PropBag.ReadProperty("FillColor", &HFF&)
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets/Returns backcolor of meter."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpMeter,shpMeter,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = shpMeter.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    shpMeter.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

Private Sub lblMessage_Click()
    RaiseEvent Click
End Sub

Private Sub picMeter_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMessage,lblMessage,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  lblMessage.Refresh
End Sub


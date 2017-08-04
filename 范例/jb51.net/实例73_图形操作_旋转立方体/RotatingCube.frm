VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Rotating Cube DEMO"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   DrawWidth       =   3
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   -1035
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   772
      TabIndex        =   0
      Top             =   1440
      Width           =   11580
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Move the mouse towards the edges of the form to adjust rotation and speed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3825
      Top             =   2835
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private X(8) As Integer
Private y(8) As Integer
'Integer arrays that hold the actual 2D coordinates of the
'8 corners of the cube.These are the values used to plot
'the cube on the form after the X,Y,Z coordinates of each cube
'corner have been converted to 2 dimensinal X and Y coordinates.
Private Const Pi = 3.14159265358979
'Constant used to convert degrees to radians
Private CenterX As Integer
Private CenterY As Integer
'The center of the 3 dimensional plane,where it's
'X=0 , Y=0 , Z=0
Private Const SIZE = 250
'The length of the cube achmes,therefore also adjusts the overall
'size.
Private Radius As Integer
'The radius of the rotation.Each one of the 8 corners of the cube
'rotates around the vertical Y axis with the same angular speed and radius
'of rotation.
Private Angle As Integer
'The value of this variable loops from 0 to 360 and it is passed
'as an argument to the COS and SIN functions (sine and cosine)
'that return the changing Z and X coordinates of each corner
'as the cube rotates around the Y axis
Private CurX As Integer
Private CurY As Integer
'Variables that hold the current mouse position on the form.
Private CubeCorners(1 To 8, 1 To 3) As Integer
'The array that holds the X,Y and Z coordinates of the 8 corners
'of the cube.Here's how the 8 corners are numbered:
'
'        7___________8
'        /|         /|             |
'       /_|________/ |             |  /
'       |3|       4| |             | /
'       | |        | |             |/
'       | |________|_|       ------|---------------> x(+)
'       | /5       | /6           /|
'       |/_________|/            / |
'      1           2        Z(+)/  |
'                                 \|/ Y(+)
'
'The center of the 3D plane is right on the center of the cube.
'So ,if SIZE the length of one achmes,it's:
'
'CenterCube(1,1) = SIZE/2   ' X coordinate of 1st corner
'CenterCube(1,2) = SIZE/2   ' Y coordinate
'CenterCube(1,3) = SIZE/2   ' Z coordinate
'
'Actually,we only need to give a value for the Y coordinates
'of each corner since that will never change during the rotation
'as all corners rotate around the Y axis ,with only their Z and X
'coordinates changing periodically.
Private Sub Form_Load()
Show
With Picture1
.Width = Label1.Width
.Height = Label1.Height
End With
Picture1.Move ScaleWidth / 2 - Picture1.ScaleWidth / 2, Picture1.Height
CenterX = ScaleWidth / 2
CenterY = ScaleHeight / 2
'Set the center of the 3D plane to reflect the center of the
'form.
Angle = 0
Radius = Sqr(2 * (SIZE / 2) ^ 2)
'Give a value to the radius of the rotation.This is
'the Pythagorean theorem that returns the length of the
'hypotenuse of a right triangle as the square root
'of the sum of the other two sides raised to the 2nd power.
CubeCorners(1, 2) = SIZE / 2
CubeCorners(2, 2) = SIZE / 2
CubeCorners(3, 2) = -SIZE / 2
CubeCorners(4, 2) = -SIZE / 2
CubeCorners(5, 2) = SIZE / 2
CubeCorners(6, 2) = SIZE / 2
CubeCorners(7, 2) = -SIZE / 2
CubeCorners(8, 2) = -SIZE / 2
'Assign a value to the Y coordinates of each cube.This
'will never change through out the rotation since the cube
'rotates around the Y axis.Play around with these if you like
'but the 3D prism will no longer resemble a cube :)
End Sub

Private Sub DrawCube()
Cls
For i = 1 To 8
X(i) = CenterX + CubeCorners(i, 1) + CubeCorners(i, 3) / 8
y(i) = CenterY + CubeCorners(i, 2) + Sgn(CubeCorners(i, 2)) * CubeCorners(i, 3) / 8
'These two lines contain the algorith that converts the
'coordinates of a point on the 3D plane (x,y,z) ,into 2
'dimensional X and Y coordinates that can be used to plot
'a point on the form.Play around with the 8's and see
'what happens...
Next
Line (X(3), y(3))-(X(4), y(4))
Line (X(4), y(4))-(X(8), y(8))
Line (X(3), y(3))-(X(7), y(7))
Line (X(7), y(7))-(X(8), y(8))
Line (X(1), y(1))-(X(3), y(3))
Line (X(1), y(1))-(X(2), y(2))
Line (X(5), y(5))-(X(6), y(6))
Line (X(5), y(5))-(X(1), y(1))
Line (X(5), y(5))-(X(7), y(7))
Line (X(6), y(6))-(X(8), y(8))
Line (X(2), y(2))-(X(4), y(4))
Line (X(2), y(2))-(X(6), y(6))
Line (X(1), y(1))-(X(4), y(4))
Line (X(2), y(2))-(X(3), y(3))

Line (X(4), y(4))-(X(8), y(8))
Line (X(3), y(3))-(X(7), y(7))
'The plotting of the cube onto the form.
'We have to draw each achmes seperately and then
' "connect"  the bottom square with the top square.
DoEvents
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
CurX = X
CurY = y
'Store the current position of the mouse cursor into
'the variable CurX,CurY.
End Sub

Private Sub Timer1_Timer()
Select Case CurX
Case Is > ScaleWidth / 2
Angle = Angle + Abs(CurX - ScaleWidth / 2) / 20
If Angle = 360 Then Angle = 0
Case Else
Angle = Angle - Abs(CurX - ScaleWidth / 2) / 20
If Angle = 0 Then Angle = 360
End Select
'Change the direction and the angular speed of the rotation
'according to the position of the mouse cursor.If it's near
'the left edge of the form then the rotation will be
'anti-clockwise ,it's near the right edge it will be
'clockwise. The closer to the center of the form the
'cursor is,the slower the cube rotates.
'The angular speed of the rotation is controlled by the
'pace at which 'Angle' (the value that we pass to the
'(COS and SIN functions) increases or decreases (increases
'for anti-clockwise rotation and decreases for clockwise
'rotation).
For i = 1 To 3 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle) * Pi / 180)
Next
For i = 2 To 4 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle + 2 * 45) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle + 2 * 45) * Pi / 180)
Next
For i = 5 To 7 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle + 6 * 45) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle + 6 * 45) * Pi / 180)
Next
For i = 6 To 8 Step 2
CubeCorners(i, 3) = Radius * Cos((Angle + 4 * 45) * Pi / 180)
CubeCorners(i, 1) = Radius * Sin((Angle + 4 * 45) * Pi / 180)
Next
'Give the new values to the X and Z coordinates of each one
'of the 8 cube corners by using the COS and SIN mathematical
'functions.Notice that corners 1 and 3 always have the same
'X and Z coordinates, as well as 2 and 4, 5 and 7,6 & 8.
'Take a look at the little scetch on the top of the form
'to see how this is explained (imagine the cube rotating
'around the Y axis)
DrawCube
End Sub

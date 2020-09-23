VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Intro to Texture Mapping - By Hou Xiong"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Step Through"
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start 3D Rotation"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Draw Textured Trapezoid"
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw Trapeziod"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Triangle"
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.PictureBox scene 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   360
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   360
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Intro to Texture Mapping
' By: Hou Xiong

' This is a very basic intro into custom texture mapping.
' It shows the small differences between drawing a
' simple triangle to a trapezoid to a 3d texture.
' It also features a simple 3d texture animation.

'Trig constants
Private Const PI = 3.141592653
Private Const DEGREES = 180 / PI
Private Const RADIANS = PI / 180

'Simplifies gdi functions through my custom classes
Dim BackBuffer As SurfaceGDI
Dim Texture1 As SurfaceGDI
'arrays for direct access bitmaps
Dim bbfbits() As Long
Dim tex1Bits() As Long

Dim doRotation As Boolean
Dim timeToEnd As Boolean
Dim angle As Single

Private Sub Check1_Click()
    'uncomment the code in each draw functions
    MsgBox "Uncomment code first."
End Sub

Private Sub Command1_Click()
    BackBuffer.Clear
    'vertices start from top left going clock-wise
    DrawTriangle 30, 150, 300, 50, 300, 300
    BackBuffer.FlipBuffer scene.hDC
End Sub

Private Sub Command2_Click()
    BackBuffer.Clear
    'vertices start from top left going clock-wise
    DrawTrap 30, 150, 300, 50, 300, 300, 30, 200
    BackBuffer.FlipBuffer scene.hDC
End Sub

Private Sub Command3_Click()
    BackBuffer.Clear
    'vertices start from top left going clock-wise
    DrawTrapTex 30, 150, 300, 50, 300, 300, 30, 200
    BackBuffer.FlipBuffer scene.hDC
End Sub

Private Sub DrawTriangle(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
    Dim y_start As Single, y_end As Single
    Dim y_top_change As Single, y_bottom_change As Single
    Dim y As Long, x As Long
    
    y_start = y0    'initialize vertical line start and end (in pixels of course)
    y_end = y0      'since it's a triangle, the start and end would be the same of the first vertex
    If x1 = x0 Then Exit Sub 'nothing to draw, prevents divide by zero
    y_top_change = (y1 - y0) / (x1 - x0)    'these are basically the slopes (rise/run)
    y_bottom_change = (y2 - y0) / (x1 - x0) 'in the top and bottom edges, these tell how much to move up or down each time we move right a scan line
    
    'scan across vertical lines going right
    For x = x0 To x1
        'go down the line and set the color bits
        For y = y_start To y_end
            bbfbits(x, y) = vbBlue  'will turn out red because of reverse BGR format
            'step through, comment out for speed gain
            'If Check1.Value = vbChecked Then
            '    DoEvents
            '    BackBuffer.FlipBuffer scene.hDC
            'End If
        Next
        
        'update the line start and end
        y_start = y_start + y_top_change
        y_end = y_end + y_bottom_change
    Next
End Sub

'This algorithm automatically removes/ignores backfaces.
'Compare to DrawTriangle().
'A new point is added and a few lines updated from DrawTriangle to allow drawing trapezoid.
Private Sub DrawTrap(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long)
    Dim y_start As Single, y_end As Single
    Dim y_top_change As Single, y_bottom_change As Single
    Dim x As Long, y As Long
    
    y_start = y0
    y_end = y3                          'updated from DrawTriangle
    If x1 = x0 Then Exit Sub
    y_top_change = (y1 - y0) / (x1 - x0)
    y_bottom_change = (y2 - y3) / (x1 - x0)   'updated from DrawTriangle
    
    For x = x0 To x1
        For y = y_start To y_end
            If (x >= 0) And (x < BackBuffer.width) And (y >= 0) And (y < BackBuffer.height) Then
                bbfbits(x, y) = vbRed   'will turn out blue because of reverse BGR format
                'step through, comment out for speed gain
                'If Check1.Value = vbChecked Then
                '    DoEvents
                '    BackBuffer.FlipBuffer scene.hDC
                'End If
            End If
        Next
        
        y_start = y_start + y_top_change
        y_end = y_end + y_bottom_change
    Next
End Sub

'This algorithm automatically removes/ignores backfaces.
'Compare to DrawTrap()
'Only a few added lines to allow for texture mapping.
Private Sub DrawTrapTex(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long)
    Dim y_start As Single, y_end As Single
    Dim y_top_change As Single, y_bottom_change As Single
    Dim x As Long, y As Long
    Dim u As Single, v As Single    'new from DrawTrap
    Dim du As Single, dv As Single  'new from DrawTrap
    
    y_start = y0
    y_end = y3
    If x1 = x0 Then Exit Sub
    y_top_change = (y1 - y0) / (x1 - x0)
    y_bottom_change = (y2 - y3) / (x1 - x0)
    
    du = (Texture1.width - 1) / (x1 - x0)  'new from DrawTrap
    
    For x = x0 To x1
        dv = (Texture1.height - 1) / (y_end - y_start)   'new from DrawTrap
        v = 0                                           'new from DrawTrap
        For y = y_start To y_end
            If (x >= 0) And (x < BackBuffer.width) And (y >= 0) And (y < BackBuffer.height) Then
                On Error Resume Next
                bbfbits(x, y) = tex1Bits(u, v)    'updated from DrawTrap
                'step through, comment out for speed gain
                'If Check1.Value = vbChecked Then
                '    DoEvents
                '    BackBuffer.FlipBuffer scene.hDC
                'End If
            End If
            v = v + dv  'new from DrawTrap
        Next
        
        y_start = y_start + y_top_change
        y_end = y_end + y_bottom_change
        u = u + du  'new from DrawTrap
    Next
End Sub

Private Sub Command4_Click()
    If doRotation Then
        doRotation = False
    Else
        Command4.Caption = "Stop 3D Rotation"
        doRotation = True
        
        Dim x1 As Single, y As Single, z1 As Single
        Dim x2 As Single, z2 As Single, zTrans As Single
        Dim angle As Single
        Dim halfWidth As Single, halfHeight As Single
        Dim x2d1 As Long, y2d1 As Long, x2d2 As Long, y2d2 As Long
        Dim xCenter As Long, yCenter As Long
        Dim convScl As Long
        
        'surface dimensions
        halfWidth = Texture1.width
        halfHeight = Texture1.height
        y = halfHeight
        
        zTrans = 1500
        xCenter = BackBuffer.width / 2
        yCenter = BackBuffer.height / 2
        convScl = BackBuffer.width
        
    
        Do While doRotation
            DoEvents
            BackBuffer.Clear
            
            'calculate the rotation
            x1 = Cos(angle * RADIANS) * halfWidth
            z1 = Sin(angle * RADIANS) * halfWidth
            'since our rotation is done in the center
            'just invert the first vertex for the adjacent vertex
            x2 = -x1
            z2 = -z1
            
            'z translation
            z1 = z1 + zTrans
            z2 = z2 + zTrans
            
            'x translation
            x1 = x1 + 0
            x2 = x2 + 0
            
            'convert to 2D coordinates
            'we'll just keep our rotation centered along the Y-axis so we'll
            'use only two verteces then flip them later for the top two vertices
            x2d2 = convScl * x1 / z1
            y2d2 = convScl * y / z1
            x2d1 = convScl * x2 / z2
            y2d1 = convScl * y / z2
            
            angle = angle + 1
            
            DrawTrapTex x2d1 + xCenter, -y2d1 + yCenter, x2d2 + xCenter, -y2d2 + yCenter, x2d2 + xCenter, y2d2 + yCenter, x2d1 + xCenter, y2d1 + yCenter
            'DrawTrap x2d1 + xCenter, -y2d1 + yCenter, x2d2 + xCenter, -y2d2 + yCenter, x2d2 + xCenter, y2d2 + yCenter, x2d1 + xCenter, y2d1 + yCenter
            'draw back face
            DrawTrap x2d2 + xCenter, -y2d2 + yCenter, x2d1 + xCenter, -y2d1 + yCenter, x2d1 + xCenter, y2d1 + yCenter, x2d2 + xCenter, y2d2 + yCenter
            
            BackBuffer.FlipBuffer scene.hDC
        Loop
        Command4.Caption = "Start 3D Rotation"
        If timeToEnd Then
            'clean up gdi surfaces
            BackBuffer.DeleteSurface
            Texture1.DeleteSurface
            EraseLongPointer bbfbits()
            EraseLongPointer tex1Bits()
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    'Initialize gdi surfaces
    Set BackBuffer = CreateSurface(scene.ScaleWidth, scene.ScaleHeight)
    Set Texture1 = CreateSurfaceFromFile(App.Path & "\Texture.bmp")
    BackBuffer.MakeLongPointer bbfbits()
    Texture1.MakeLongPointer tex1Bits()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not doRotation Then
        'clean up gdi surfaces
        BackBuffer.DeleteSurface
        Texture1.DeleteSurface
        EraseLongPointer bbfbits()
        EraseLongPointer tex1Bits()
    Else
        doRotation = False
        timeToEnd = True
    End If
End Sub

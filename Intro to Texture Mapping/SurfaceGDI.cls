VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SurfaceGDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' SurfaceGDI
' By: Hou Xiong
'
' Simplifies gdi functions.
' If you decide to include these classes in
' your projects, please give me some credit.

Option Explicit

Public hDC As Long
Public hBMP As Long
Public hDib As Long
Public width As Long
Public height As Long

Private lBits() As Long
Private bBits() As RGBQUAD
Private sa As SAFEARRAY2D
Private bitmapSize As Long
Public widthBytes As Long

Public Sub InitSurface()
        With sa
            .cbElements = 4
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = height
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = width
            .pvData = hDib
        End With
        
        CopyMemory ByVal VarPtrArray(lBits()), VarPtr(sa), 4
        CopyMemory ByVal VarPtrArray(bBits()), VarPtr(sa), 4
        
        bitmapSize = width * height * 4
        widthBytes = width * 4
End Sub

Public Function GetPixel(ByVal x As Long, ByVal y As Long) As Long
    If (x >= 0) And (x < width) And (y >= 0) And (y < height) Then
        GetPixel = lBits(x, y)
    Else
        GetPixel = -1
    End If
End Function

Public Sub SetPixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
    If (x >= 0) And (x < width) And (y >= 0) And (y < height) Then
        lBits(x, y) = Color
    End If
End Sub

Public Sub FlipBuffer(ByVal hDC As Long)
    BitBlt hDC, 0, 0, width, height, Me.hDC, 0, 0, vbSrcCopy
End Sub

Public Sub Clear()
    ZeroMemory lBits(0, 0), bitmapSize
End Sub

Public Function MakeLongPointer(Pixels() As Long) As Boolean
    If hDib = 0 Then Exit Function
    CopyMemory ByVal VarPtrArray(Pixels()), VarPtr(sa), 4
    MakeLongPointer = True
End Function

Public Function MakeRGBPointer(ByVal lPixels As Long) As Boolean
    If hDib = 0 Then Exit Function
    CopyMemory ByVal lPixels, VarPtr(sa), 4
    MakeRGBPointer = True
End Function

Public Sub DeleteSurface()
    If hDC = 0 Then Exit Sub
    
    CopyMemory ByVal VarPtrArray(lBits()), 0&, 4
    CopyMemory ByVal VarPtrArray(bBits()), 0&, 4
    
    DeleteObject hBMP
    DeleteDC hDC
    
    hDC = 0
    hBMP = 0
    hDib = 0
    width = 0
    height = 0
    bitmapSize = 0
    widthBytes = 0
End Sub

Private Sub Class_Terminate()
    DeleteSurface
End Sub

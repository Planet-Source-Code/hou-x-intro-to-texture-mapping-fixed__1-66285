Attribute VB_Name = "basGDI"

' SurfaceGDI
' By: Hou Xiong
'
' Simplifies gdi functions.
' If you decide to include these classes in
' your projects, please give me some credit.

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Function VarPtrArray Lib "MSVBVM60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Const LR_LOADFROMFILE = 16

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Public Function CreateSurface(ByVal width As Long, ByVal height As Long) As SurfaceGDI
    Set CreateSurface = New SurfaceGDI
    
    With CreateSurface
        Dim hDib As Long
        
        .width = width
        .height = height
        
        Dim bi As BITMAPINFO
        With bi.bmiHeader
            .biSize = LenB(bi)
            .biWidth = width
            .biHeight = -height
            .biPlanes = 1
            .biBitCount = 32
            '.biSizeImage = ((((width * 32) + 31) \ 32) * 4) * height
        End With
        
        .hDC = CreateCompatibleDC(0)
        If .hDC = 0 Then GoTo CreateSurfaceError
        .hBMP = CreateDIBSection(.hDC, bi, 0, hDib, 0, 0)
        If .hBMP = 0 Then GoTo CreateSurfaceError
        If SelectObject(.hDC, .hBMP) = 0 Then GoTo CreateSurfaceError
        If hDib = 0 Then GoTo CreateSurfaceError
        .hDib = hDib
        
        .InitSurface
    End With
    
    Exit Function
    
CreateSurfaceError:
End Function

Public Function CreateSurfaceFromFile(ByVal FileName As String) As SurfaceGDI
    Set CreateSurfaceFromFile = New SurfaceGDI
    
    With CreateSurfaceFromFile
        Dim tDC As Long, tBMP As Long, hDib As Long
        
        tDC = CreateCompatibleDC(0)
        If tDC = 0 Then GoTo CreateFromFileError
        tBMP = LoadImage(0, FileName, 0, 0, 0, LR_LOADFROMFILE)
        If tBMP = 0 Then GoTo CreateFromFileError
        SelectObject tDC, tBMP
        
        Dim bmp As BITMAP
        GetObject tBMP, LenB(bmp), bmp
        
        .width = bmp.bmWidth
        .height = bmp.bmHeight
        
        Dim bi As BITMAPINFO
        With bi.bmiHeader
            .biSize = LenB(bi)
            .biWidth = CreateSurfaceFromFile.width
            .biHeight = -CreateSurfaceFromFile.height
            .biPlanes = 1
            .biBitCount = 32
            '.biSizeImage = ((((CreateSurfaceFromFile.width * 32) + 31) \ 32) * 4) * CreateSurfaceFromFile.height
        End With
        
        .hDC = CreateCompatibleDC(0)
        If .hDC = 0 Then GoTo CreateFromFileError
        .hBMP = CreateDIBSection(.hDC, bi, 0, hDib, 0, 0)
        If .hBMP = 0 Then GoTo CreateFromFileError
        SelectObject .hDC, .hBMP
        If hDib = 0 Then GoTo CreateFromFileError
        
        BitBlt .hDC, 0, 0, .width, .height, tDC, 0, 0, vbSrcCopy
        DeleteObject tBMP
        DeleteDC tDC
        
        .hDib = hDib
        
        .InitSurface
    End With
    
    Exit Function
    
CreateFromFileError:
End Function

Public Sub EraseLongPointer(Pixels() As Long)
    CopyMemory ByVal VarPtrArray(Pixels()), 0&, 4
End Sub

Public Sub EraseRGBPointer(Pixels() As RGBQUAD)
    CopyMemory ByVal VarPtrArray(Pixels()), 0&, 4
End Sub

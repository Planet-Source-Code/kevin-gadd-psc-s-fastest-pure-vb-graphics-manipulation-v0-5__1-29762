Attribute VB_Name = "mdlGDI"
'
'   Image Processing demo for Planet-Source-Code
'   All code within is copyright (c) 2001 Kevin Gadd unless otherwise noted
'   Uses 'CapturePicture' by Benjamin Marty
'

'   Let's declare all our variables, shall we?
Option Explicit

'   Look out, API declarations on the loose!
'   Code starts at 'CreateMemoryDC'
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef pPictDesc As PICTDESC, ByRef riid As iid, ByVal fOwn As Boolean, ByRef ppvObj As StdPicture) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal sWidth As Long, ByVal sHeight As Long, ByVal PointerToBits As Long, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'   It's a rect. What else is there to say?
Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'   Used to create the IPictureDisp object.
Public Type iid
    X As Long
    S1 As Integer
    S2 As Integer
    C(0 To 7) As Byte
End Type

'   Describes the picture to be created.
Public Type PICTDESC
    lngSize As Long
    lngType As Long
    lngBitmap As Long
    lngPalette As Long
End Type

'   Settings for retrieving/changing bitmap information
Public Type BITMAPINFOHEADER
        lngSize As Long
        lngWidth As Long
        lngHeight As Long
        intPlanes As Integer
        intBitCount As Integer
        lngCompression As Long
        lngSizeImage As Long
        lngXPelsPerMeter As Long
        lngYPelsPerMeter As Long
        lngClrUsed As Long
        lngClrImportant As Long
End Type

'   32-bit BGRA Pixel (Used by GDI)
Public Type Pixel32
        Blue As Byte
        Green As Byte
        Red As Byte
        Alpha As Byte
End Type

'   The BitmapInfo struct passed to GetDIBits/SetDIBits/StretchDIBits
Public Type BITMAPINFO
        bihHeader As BITMAPINFOHEADER
        p32Palette As Pixel32
End Type

Public Const IMAGE_BITMAP As Long = 0
Public Const IMAGE_ICON As Long = 1
Public Const IMAGE_CURSOR As Long = 2
Public Const IMAGE_ENHMETAFILE As Long = 3
Public Const BLACK_BRUSH As Long = 4
Public Const BI_RGB As Long = 0
Public Const DIB_RGB_COLORS As Long = 0
Public Const OBJ_PAL As Long = 5
Public Const S_OK As Long = 0

'   Wrappers for DC creation/deletion
Public Function CreateMemoryDC() As Long
'On Error Resume Next
Dim m_lngWindow As Long, m_lngDC As Long
    m_lngWindow = GetDesktopWindow
    m_lngDC = GetDC(m_lngWindow)
    CreateMemoryDC = CreateCompatibleDC(m_lngDC)
    ReleaseDC m_lngWindow, m_lngDC
End Function

Public Sub DeleteMemoryDC(ByRef hDC As Long)
'On Error Resume Next
    DeleteDC hDC
    hDC = 0
End Sub

'   CopyPixelsToDC
'   Copies an array of pixels to a DC, none of that clipping stuff. Faster.
Sub CopyPixelsToDC(ByVal hDC As Long, ByRef Pixels() As Pixel32)
Dim m_bmiPixels As BITMAPINFO
    With m_bmiPixels.bihHeader
        ' Fill the data structure
        .lngSize = Len(m_bmiPixels.bihHeader)
        ' 32 bits per pixel (BGRA)
        .intBitCount = 32
        ' Single plane
        .intPlanes = 1
        ' Width
        .lngWidth = UBound(Pixels, 1) + 1
        ' Height, negative to indicate the image is NOT bottom-up
        .lngHeight = -(UBound(Pixels, 2) + 1)
        ' Total size of the image
        .lngSizeImage = .lngWidth * -.lngHeight
        ' No compression
        .lngCompression = BI_RGB
        ' Copy the pixels to the DC
        Call StretchDIBits(hDC, 0, 0, .lngWidth, -.lngHeight, 0, 0, .lngWidth, -.lngHeight, VarPtr(Pixels(0, 0)), m_bmiPixels, DIB_RGB_COLORS, vbSrcCopy)
    End With
End Sub

'   DrawPixelsToDC
'   Draws an array of pixels to a DC, with clipping.
Sub DrawPixelsToDC(ByVal hDC As Long, ByRef Pixels() As Pixel32, ByVal X As Long, ByVal Y As Long, Optional ByVal Width As Long = 0, Optional ByVal Height As Long = 0, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0)
Dim m_bmiPixels As BITMAPINFO
    With m_bmiPixels.bihHeader
        ' Fill the data structure
        .lngSize = Len(m_bmiPixels.bihHeader)
        ' 32 bits per pixel (BGRA)
        .intBitCount = 32
        ' Single plane
        .intPlanes = 1
        ' Width
        .lngWidth = UBound(Pixels, 1) + 1
        ' Height, negative to indicate the image is NOT bottom-up
        .lngHeight = (UBound(Pixels, 2) + 1)
        ' Total size of the image
        .lngSizeImage = .lngWidth * .lngHeight
        ' No compression
        .lngCompression = BI_RGB
        ' If extra parameters were not included, fill them
        If Width <= 0 Then Width = .lngWidth
        If Height <= 0 Then Height = .lngHeight
        ' Copy the pixels to the DC
        Call StretchDIBits(hDC, X, Y, Width, Height, X2, Y2 + Height + 1, Width, -Height, VarPtr(Pixels(0, 0)), m_bmiPixels, DIB_RGB_COLORS, vbSrcCopy)
    End With
End Sub

'   GetPictureArrayInv
'   Retrieves a Long array of an image's pixels, with the correct orientation.
'   This DOES fix the fact that GDI bitmaps are bottom-up.
'   Loosely based on some GetDIBits code Ben Marty showed me a long time ago.
Function GetPictureArrayInv(ByRef Picture As IPictureDisp) As Pixel32()
Dim m_lngYOffset As Long
Dim m_bmiDest As BITMAPINFO
Dim m_p32Pixels() As Pixel32, m_lngDC As Long
     
    ' Quickly create a DC in memory
    m_lngDC = CreateMemoryDC
    
    With m_bmiDest.bihHeader
        .lngSize = Len(m_bmiDest.bihHeader)
        .intPlanes = 1
    End With
    
    ' Get header information (interested mainly in size)
    If GetDIBits(m_lngDC, Picture.Handle, 0, 0, ByVal 0&, m_bmiDest, DIB_RGB_COLORS) = 0 Then
        ' Error occurred, clean up
        DeleteMemoryDC m_lngDC
        Exit Function
    End If
    
    With m_bmiDest.bihHeader
        .intBitCount = 32
        .lngCompression = BI_RGB
    End With
    
    m_lngYOffset = (m_bmiDest.bihHeader.lngHeight - 1)
    
    ' Allocate space according to retrieved size
    ReDim m_p32Pixels(0 To m_bmiDest.bihHeader.lngWidth - 1, 0 To m_bmiDest.bihHeader.lngHeight - 1)
    
    m_bmiDest.bihHeader.lngHeight = -m_bmiDest.bihHeader.lngHeight
    
    ' Now get the bits
    If GetDIBits(m_lngDC, Picture.Handle, 0, Abs(m_bmiDest.bihHeader.lngHeight), m_p32Pixels(0, 0), m_bmiDest, DIB_RGB_COLORS) = 0 Then
        ' Error occurred, clean up
        DeleteMemoryDC m_lngDC
        Exit Function
    End If
    
    GetPictureArrayInv = m_p32Pixels()
    
    DeleteMemoryDC m_lngDC
    
End Function
'
''   CapturePicture
''   Creates an IPictureDisp from an area of an hDC
''   Written by Ben Marty, cleaned up by me
'Public Function CapturePicture(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As IPictureDisp
'    Dim m_iidPicture As iid
'    Dim m_pcdPicture As PICTDESC
'    Dim m_lngMemDC As Long
'    Dim m_lngBitmap As Long
'    Dim m_lngOldBitmap As Long
'    Dim m_picPicture As IPictureDisp
'    Dim m_rctBitmap As Rect
'
'    ' IID_IPictureDisp
'    m_iidPicture.X = &H7BF80981
'    m_iidPicture.S1 = &HBF32
'    m_iidPicture.S2 = &H101A
'    m_iidPicture.C(0) = &H8B
'    m_iidPicture.C(1) = &HBB
'    m_iidPicture.C(2) = &H0
'    m_iidPicture.C(3) = &HAA
'    m_iidPicture.C(4) = &H0
'    m_iidPicture.C(5) = &H30
'    m_iidPicture.C(6) = &HC
'    m_iidPicture.C(7) = &HAB
'
'    m_pcdPicture.lngSize = Len(m_pcdPicture)
'    m_lngMemDC = CreateCompatibleDC(hDC)
'    If m_lngMemDC = 0 Then
'        Exit Function
'    End If
'    m_lngBitmap = CreateCompatibleBitmap(hDC, Width, Height)
'    If m_lngBitmap = 0 Then
'        DeleteDC m_lngMemDC
'        Exit Function
'    End If
'    m_lngOldBitmap = SelectObject(m_lngMemDC, m_lngBitmap)
'    If Left >= 0 And Top >= 0 Then
'        If BitBlt(m_lngMemDC, 0, 0, Width, Height, hDC, Left, Top, vbSrcCopy) = 0 Then
'            SelectObject m_lngMemDC, m_lngOldBitmap
'            DeleteDC m_lngMemDC
'            DeleteObject m_lngBitmap
'            Exit Function
'        End If
'    Else
'        With m_rctBitmap
'           .Left = 0
'           .Top = 0
'           .Right = Width
'           .Bottom = Height
'        End With
'        FillRect m_lngMemDC, m_rctBitmap, GetStockObject(BLACK_BRUSH)
'    End If
'    SelectObject m_lngMemDC, m_lngOldBitmap
'    m_pcdPicture.lngBitmap = m_lngBitmap
'    m_pcdPicture.lngPalette = GetCurrentObject(hDC, OBJ_PAL)
'    m_pcdPicture.lngType = IMAGE_BITMAP
'    If OleCreatePictureIndirect(m_pcdPicture, m_iidPicture, True, m_picPicture) <> S_OK Then
'        DeleteDC m_lngMemDC
'        DeleteObject m_lngBitmap
'        Exit Function
'    End If
'    Set CapturePicture = m_picPicture
'    Set m_picPicture = Nothing
'    DeleteDC m_lngMemDC
'End Function
'

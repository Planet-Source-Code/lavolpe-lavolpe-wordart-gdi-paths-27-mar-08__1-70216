VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPathTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FillPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32.dll" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, ByRef pcRGBQuad As Any) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Const BLACK_BRUSH As Long = 4
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Type BITMAPINFOHEADER
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 1) As Long   ' 2 color bitmap
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private m_DC As Long
Private m_Bitmap As Long
Private m_OldBmp As Long

Private m_OffsetSrc As POINTAPI ' offset from left edge of m_Bitmap & start of drag image inside bitmap
Private m_MouseOffset As POINTAPI
Private m_DestSize As POINTAPI  ' destination width/height of m_Bitmap
Private m_SrcSize As POINTAPI   ' original size of the source used to create drag image
Private m_Position As POINTAPI  ' the current X,Y where drag image is rendered




Public Sub InitializeSizeMove(ByVal hdc As Long, ByVal MouseX As Single, ByVal MouseY As Single, SourcePath As clsWApath)

    Dim tDC As Long, hBrush As Long, tRect As RECT
    Dim tDIB As BITMAPINFO, leftTopPt As POINTAPI
    
    ' create a black & white, XOR-able bitmap
    tDC = GetDC(0&)
    
    SourcePath.GetBoundingBox_Long False, leftTopPt.x, leftTopPt.y, m_SrcSize.x, m_SrcSize.y
    m_DestSize.x = Sqr(m_SrcSize.x * m_SrcSize.x + m_SrcSize.y * m_SrcSize.y)
    m_DestSize.y = m_DestSize.x
    
    With tDIB.bmiHeader
        .biBitCount = 1
        .biClrImportant = 2
        .biClrUsed = 2
        .biWidth = m_DestSize.x
        .biHeight = .biWidth * 2 ' include mask half
        .biPlanes = 1
        .biSize = 40
    End With
    tDIB.bmiColors(1) = vbWhite
    m_Bitmap = CreateDIBSection(tDC, tDIB, 0&, ByVal 0&, 0&, 0&)
    
    If m_Bitmap Then
        m_DC = CreateCompatibleDC(tDC)
        If m_DC = 0& Then
            ReleaseDC 0&, tDC
            DeleteObject m_Bitmap
        Else
    
            m_OldBmp = SelectObject(m_DC, m_Bitmap)
            ReleaseDC 0&, tDC
    
            ' paint the bottom half white
            hBrush = GetStockObject(0&)
            tRect.Top = m_DestSize.y
            tRect.Bottom = m_DestSize.y * 2
            tRect.Right = m_DestSize.x
            FillRect m_DC, tRect, hBrush
            
            m_OffsetSrc.x = (m_DestSize.x - m_SrcSize.x) \ 2
            m_OffsetSrc.y = (m_DestSize.y - m_SrcSize.y) \ 2
            
            ' using GDI, not GDI+, render the top half, then the bottom half, alternating pen colors
            BeginPath m_DC
            SourcePath.RenderOutline_XOR m_DC, vbBlack, , psXOR_Solid, -leftTopPt.x + m_OffsetSrc.x, -leftTopPt.y + m_OffsetSrc.y
            EndPath m_DC
            FillPath m_DC
            StretchBlt m_DC, 0, tRect.Top, tRect.Right, tRect.Top, m_DC, 0, 0, tRect.Top, tRect.Right, vbSrcInvert
            
            ' draw the bitmap for first time
            
            
            m_Position.x = leftTopPt.x - m_OffsetSrc.x
            m_Position.y = leftTopPt.y - m_OffsetSrc.y
            
            StretchBlt hdc, m_Position.x, m_Position.y, m_DestSize.x, m_DestSize.y, m_DC, 0, 0, m_DestSize.x, m_DestSize.y, vbSrcInvert
            ' create offsets from current mouse X,Y to top/left edge of path
    
            m_MouseOffset.x = MouseX - m_Position.x
            m_MouseOffset.y = MouseY - m_Position.y
            m_SrcSize = m_DestSize
            
        End If
    End If
    

End Sub

Public Sub TerminateSizeMove(ByVal hdc As Long, ByVal Cancel As Boolean, newX As Long, newY As Long, ScaleX As Single, ScaleY As Single)

    If m_DC Then
        ' erase the last drawn path
        StretchBlt hdc, m_Position.x, m_Position.y, m_DestSize.x, m_DestSize.y, m_DC, 0, 0, m_SrcSize.x, m_SrcSize.y, vbSrcInvert
        ' clean up
        If m_Bitmap Then DeleteObject SelectObject(m_DC, m_OldBmp)
        DeleteDC m_DC
    End If
    If Not Cancel Then
'        Dim tRect As RECT
'        tRect.Left = m_Position.x: tRect.Top = m_Position.y
'        tRect.Right = m_Position.x + m_DestSize.x
'        tRect.Bottom = m_Position.y + m_DestSize.y
'        FrameRect hdc, tRect, GetStockObject(0&)
        
        ScaleX = m_DestSize.x / m_SrcSize.x
        ScaleY = m_DestSize.y / m_SrcSize.y
        newX = m_Position.x + m_OffsetSrc.x * ScaleX
        newY = m_Position.y + m_OffsetSrc.y * ScaleY
    End If


End Sub

Public Sub UpdateSizeMove(ByVal hdc As Long, ByVal x As Single, ByVal y As Single, ByVal mode As Long)

        
    ' Mode are positions from center left edge to bottom right corner in clockwise order ....
    ' Mode 0 is for moving, not sizing
    
    ' 2     3     4
    '
    ' 1           5
    '
    ' 8     7     6
    
    If m_DC Then
        ' erase last drawn path
        StretchBlt hdc, m_Position.x, m_Position.y, m_DestSize.x, m_DestSize.y, m_DC, 0, 0, m_SrcSize.x, m_SrcSize.y, vbSrcInvert
        ' update left/top position
        Select Case mode
        Case 0 ' moving, not sizing
            m_Position.x = x - m_MouseOffset.x
            m_Position.y = y - m_MouseOffset.y
        Case 1 ' left edge center, size WE
            m_DestSize.x = m_DestSize.x + (m_Position.x - (x - m_MouseOffset.x))
            m_Position.x = x - m_MouseOffset.x
        Case 2 ' top left corner, size NWSE
            m_DestSize.x = m_DestSize.x + (m_Position.x - (x - m_MouseOffset.x))
            m_DestSize.y = m_DestSize.y + (m_Position.y - (y - m_MouseOffset.y))
            m_Position.x = x - m_MouseOffset.x
            m_Position.y = y - m_MouseOffset.y
        Case 3 ' top center, size NS
            m_DestSize.y = m_SrcSize.y + (m_Position.y - (y - m_MouseOffset.y))
            m_Position.y = y - m_MouseOffset.y
        Case 4 ' top right corner, size NESW
            m_DestSize.x = m_SrcSize.x - (m_Position.x - (x - m_MouseOffset.x))
            m_DestSize.y = m_DestSize.y + (m_Position.y - (y - m_MouseOffset.y))
            m_Position.y = y - m_MouseOffset.y
        Case 5 ' right edge center, size WE
            m_DestSize.x = m_SrcSize.x - (m_Position.x - (x - m_MouseOffset.x))
        Case 6 ' bottom right corner, size NWSE
            m_DestSize.x = m_SrcSize.x - (m_Position.x - (x - m_MouseOffset.x))
            m_DestSize.y = m_SrcSize.y - (m_Position.y - (y - m_MouseOffset.y))
        Case 7 ' bottom edge center, size NS
            m_DestSize.y = m_SrcSize.y - (m_Position.y - (y - m_MouseOffset.y))
        Case 8 ' bottom left corner, size NESW
            m_DestSize.x = m_DestSize.x + (m_Position.x - (x - m_MouseOffset.x))
            m_DestSize.y = m_SrcSize.y - (m_Position.y - (y - m_MouseOffset.y))
            m_Position.x = x - m_MouseOffset.x
        End Select
        
        ' render at new position
        StretchBlt hdc, m_Position.x, m_Position.y, m_DestSize.x, m_DestSize.y, m_DC, 0, 0, m_SrcSize.x, m_SrcSize.y, vbSrcInvert
    
    End If

End Sub

'   some thoughts about rubberband objects
'   1. Path first created, its origin points are kept forever
'   2. When moving, also move source points
'   3. When resizing, maintain scaleX, scaleY values forever
'   4. When rotating, maintain rotation value forever
'   5. When warping, maintain warp points
'   6. Original path is never modified
'
'   UDT would look something like the following:
'       Origin(0 to 3) as PointF
'       Warp(0 to 3) as PointF
'       Rotation As Single
'       Scaler As PointF
'
'   So when rendering is required:
'   1. Rotate & scale using cached values
'   2. Warp the scaled/rotated points
'

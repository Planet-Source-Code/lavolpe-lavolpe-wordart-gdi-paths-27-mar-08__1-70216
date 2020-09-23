VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWApath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CLASS PURPOSE & GOALS
' -------------------------------
' 1. Wrap many common GDI+ path-related methods/routines
' 2. Handle complicated calculations & path-related matrix operations
' 3. Provide patches/workarounds for GDI+ path-related buggy code (warping, XOR drawing)
' 4. Provide flexible rendering routines & clone routines
' 5. Re-usable and expandable

' IMPORTANT:  ARRAY POINTER USAGE
' --------------------------------
' Since this is a class and passing UDTs to it isn't allowed unless the UDTs are delcared public in a bas module.
' Some of the functions ask for array pointers. This has some benefits and only one drawback.
' The benfits are that you can use various storage items: Singles(), POINTF(), Bytes()

' First, GDI+ uses POINTF structures quite often. This is simply an X,Y structure where X,Y are Singles & not Longs.
' Since the structure are Singles, you can either declare and use an array of POINTF structures in your project or
' use an array of Singles instead.  Always declare your arrays like the following:
'       -- array of Singles: Dim mySingleArray(0 To 1, 0 To numberPoints-1) << 0 & 1 are X,Y coords respectively
'       -- array of POINTF:  Dim myPOINTFarray(0 To numberPoints-1)
' Now when you need to pass pointers to the class, simply pass it using VB's VarPtr function:
'       -- array of Singles:  VarPtr(mySingleArray(0, 0))
'       -- array of POINTF:   VarPtr(myPOINTFarray(0))
' The drawback, crashes if you pass invalid pointers or do not size your array appropriately
' All functions that expect pointers have a PTR suffix in their function name.
' ALWAYS ENSURE THE ARRAY IS SIZED BEFORE ATTEMPTING TO PASS ARRAY POINTERS

' ABOUT RENDERING, SCALING, OFFSETS
' ---------------------------------
' The path is never scaled, rotated or moved/repositioned once it is created.
' The class tracks the scale desired, angle desired, position desired, etc.
' When rendering occurs, a matrix is created to be used to place the path
' using set Angle, Scale, Offsets and then rendered to the DC via the matrix.
' This actually has several benefites and one or two drawbacks.

' The benefits? Original path is never altered, this promotes better quality
'   renderings. Since path is never physically modified, gradient/path brushes
'   do not need to be destroyed/modifiied/recreated when a path is resized or moved.
'   Rendering speed is improved in many scenarios, but not all
'   Paths can be cloned "as rendered" via the Clone function

' The drawbacks? Only major one I can think of :: hit testing a path on a DC if the project
'   requires the ability to see if mouse is over a path. Since the path handle in this
'   class is not the same dimensions/rotation as the rendered path, hit testing will fail.
'   To resolve this, recommend creating a second class and calling this class' Clone
'   method to get the path in rendered form.  When/If the rendered path changes scale,
'   rotation or position, then replace the path with another call to the Clone method.
'   -- Another drawback is that repainting via a matrix is slower than if the path
'      was created at specific angle/size/position. However, this drawback only occurs
'      when a DC's AutoRedraw=False otherwise, rendering generally only occurs when
'      path properties change which is less common overall. Animating via matrix is faster
'      overall than continually recreating a path to match the animated position/scale/rotation.
' -------------------------------------------------------------------------------------

' ---=====| NON-GDI+ APIS |====---

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
' following used for XOR rendering of path
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long
Private Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hDC As Long, ByRef lppt As Any, ByVal cCount As Long) As Long
Private Declare Function SetROP2 Lib "gdi32.dll" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
' following used to saving path as an image
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function UnionRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long

' ---=====| GDI+ TEXT RELEATED APIS |====---
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTF, ByVal mFormat As Long) As Long

' ---=====| GDI+ PATH RENDERING APIS |====---
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long

' ---=====| GDI+ DC RELATED APIS |====---
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMode As SmoothingMode) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal interpolation As InterpolationMode) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixOffsetMode As PixelOffsetMode) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipSetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mMatrix As Long) As Long
Private Declare Function GdipSaveGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByRef mState As Long) As Long
Private Declare Function GdipRestoreGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mState As Long) As Long
Private Declare Function GdipIsVisiblePathPoint Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mGraphics As Long, ByRef mResult As Long) As Long

' ---=====| GDI+ PATH CREATION/MODIFICATION APIS |====---
Private Declare Function GdipSetPathMarker Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipAddPathArc Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathBezier Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single, ByVal mX3 As Single, ByVal mY3 As Single, ByVal mX4 As Single, ByVal mY4 As Single) As Long
Private Declare Function GdipAddPathBeziers Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPoints As Any, ByVal mCount As Long) As Long
Private Declare Function GdipAddPathCurve2 Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPoints As Any, ByVal mCount As Long, ByVal mTension As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipAddPathLine Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipAddPathLine2 Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPoints As Any, ByVal mCount As Long) As Long
Private Declare Function GdipAddPathPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mAddingPath As Long, ByVal mConnect As Long) As Long
Private Declare Function GdipAddPathPie Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathPolygon Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPoints As Any, ByVal mCount As Long) As Long
Private Declare Function GdipAddPathRectangle Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipClearPathMarkers Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mFillMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipCreatePath2 Lib "GdiPlus.dll" (ByVal mPOINTF As Long, ByVal mBYTE As Long, ByVal mCount As Long, ByVal mFillMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipClonePath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mClonePath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipSetPathFillMode Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mFillMode As Long) As Long
Private Declare Function GdipWindingModeOutline Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mMatrix As Long, ByVal mFlatness As Single) As Long

' ---=====| GDI+ MATRIX RELEATED APIS |====---
Private Declare Function GdipCreateMatrix Lib "GdiPlus.dll" (ByRef mMatrix As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "GdiPlus.dll" (ByVal mMatrix As Long, ByVal mOffsetX As Single, ByVal mOffsetY As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipDeleteMatrix Lib "GdiPlus.dll" (ByVal mMatrix As Long) As Long
Private Declare Function GdipCreateMatrix2 Lib "GdiPlus.dll" (ByVal mM11 As Single, ByVal mM12 As Single, ByVal mM21 As Single, ByVal mM22 As Single, ByVal mDx As Single, ByVal mDy As Single, ByRef mMatrix As Long) As Long
Private Declare Function GdipCloneMatrix Lib "GdiPlus.dll" (ByVal mMatrix As Long, ByRef mCloneMatrix As Long) As Long
Private Declare Function GdipTransformPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mMatrix As Long) As Long

' ---=====| GDI+ PATH PROPERTIES APIS |====---
Private Declare Function GdipGetPointCount Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mCount As Long) As Long
Private Declare Function GdipGetPathData Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPathData As PATHDATA) As Long
Private Declare Function GdipGetPathPoints Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mPointsPtr As Long, ByVal mCount As Long) As Long
Private Declare Function GdipGetPathTypes Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mTypesPtr As Long, ByVal mCount As Long) As Long
Private Declare Function GdipGetPathWorldBounds Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mBounds As RECTF, ByVal mMatrix As Long, ByVal mPen As Long) As Long
Private Declare Function GdipGetPathLastPoint Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mLastPoint As POINTF) As Long

' ---=====| GDI+ IMAGE SAVING APIS |====---
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, BITMAP As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long


' ---=====| Private Enumerations & Constants |====---
Private Const QualityModeInvalid As Long = -1&
Private Const QualityModeDefault As Long = 0&
Private Const QualityModeLow As Long = 1&
Private Const QualityModeHigh As Long = 2&
Private Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone = QualityModeHigh + 1
    SmoothingModeAntiAlias8x4 = QualityModeHigh + 2
    SmoothingModeAntiAlias = SmoothingModeAntiAlias8x4
    SmoothingModeAntiAlias8x8 = QualityModeHigh + 3
End Enum
Private Enum InterpolationMode
   InterpolationModeInvalid = QualityModeInvalid
   InterpolationModeDefault = QualityModeDefault
   InterpolationModeLowQuality = QualityModeLow
   InterpolationModeHighQuality = QualityModeHigh
   InterpolationModeBilinear
   InterpolationModeBicubic
   InterpolationModeNearestNeighbor
   InterpolationModeHighQualityBilinear
   InterpolationModeHighQualityBicubic
End Enum
Private Enum PixelOffsetMode
   PixelOffsetModeInvalid = QualityModeInvalid
   PixelOffsetModeDefault = QualityModeDefault
   PixelOffsetModeHighSpeed = QualityModeLow
   PixelOffsetModeHighQuality = QualityModeHigh
   PixelOffsetModeNone    ' No pixel offset
   PixelOffsetModeHalf     ' Offset by -0.5 -0.5 for fast anti-alias perf
End Enum
Private Enum PathPointType
    PathPointTypeStart = 0
    PathPointTypeLine = 1
    PathPointTypeBezier = 3
    PathPointTypePathTypeMask = &H7
    PathPointTypeDashMode = &H10
    PathPointTypePathMarker = &H20
    PathPointTypeCloseSubpath = &H80
End Enum
Private Enum MatrixOrder
    MatrixOrderPrepend = &H0
    MatrixOrderAppend = &H1
End Enum
Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum

' ---=====| Public Custom Enumerations |====---
Public Enum ePathBrushFillMode
    FillModeAlternate = &H0
    FillModeWinding = &H1
End Enum
Public Enum ePathXORstyle
    psXOR_Solid = 0
    psXOR_Dashed = 1
    psXOR_Dotted = 2
    psXOR_DashDot = 3
    psXOR_DashDotDot = 4
End Enum
Public Enum eFontStyle
    FontStyleRegular = &H0
    FontStyleBold = &H1
    FontStyleItalic = &H2
    FontStyleBoldItalic = &H3
    FontStyleUnderline = &H4
    FontStyleStrikeout = &H8
End Enum
Public Enum eSaveAs
    waSaveAsPNG = 0
    waSaveAsJPG = 1
    waSaveAsBMP = 2
End Enum


' ---=====| API UDTs |====---
Private Type PATHDATA       ' GDI+ UDT for retrieving path points & types
    Count As Long
    PointsPTR As Long
    TypesPtr As Long
End Type
Private Type RECTF          ' GDI+ UDT :: RECT structure using Singles vs Longs
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
Private Type POINTF         ' GDI+ UDT :: POINTAPI structure using singles vs Longs
    X As Single
    Y As Single
End Type
Private Type EncoderParameter   ' GDI+ UDT: Image conversion codec parameters
    GUID(0 To 3)   As Long
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type
Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type
Private Type ImageCodecInfo     ' GDI+ UDT: Image conversion codec references
    ClassID(0 To 3)   As Long
    FormatID(0 To 3)  As Long
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type
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
    bmiColors As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' ---=====| Custom UDTs |====---
Private Type PATHSEGMENTEX  ' one per marked subpath (marking is optional)
    StartIndex As Long      ' first point in subpath/marker
    EndIndex As Long        ' last point in subpath/marker
    Count As Long           ' number points in segment: may differ from EndIndex-StartIndex+1
    Segments() As POINTF    ' small array (1 to a few entries), sub segment indexes/lengths
End Type
Private Type PATHMARKERS    ' summary of subpath information
    Count As Long           ' number of path segments/subpaths
    Length As Double        ' total length of entire path (measured on demand)
    Markers() As PATHSEGMENTEX ' see PATHSEGMENTEX above
End Type
Private Type LINEARCURVEREF  ' one per marked subpath when NavigatePath is called (marking is optional)
    Distance() As Single     ' time-to-distance lookup of size 1/m_BezierPrecision. See NavigatePath
    Index As Long            ' indicates which subpath segment lookup is referencing
End Type
Private Type PATHATTRIBUTES
    Origin(0 To 3) As POINTF            ' Path first created, its origin
    DisplayLocale As POINTF             ' Display coordinates for top/left edge
    ptOffsets(0 To 3) As POINTF         ' Warp & relocation offsets (see below)
    Transformation(0 To 5) As Single    ' Cached transformation matrix
    TransOverride() As Single           ' user-defined transformation matrix
    Scaler As POINTF                    ' The ScaleX & ScaleY of the path
    Rotation As Single                  ' The angle of rotation
    handle As Long                      ' The path handle
    WarpStyle As Long                   ' 0=none, 1=bilinear, 2=perspective, 32/33 re-warp path
End Type
' ^^ About the offsets. Source path is never modified, its offsets/scales are used to render
' PtOffsets(0-3).X from Origin(0-3).X to Warped X. Relational distances only, not actual X,Y coordinates
' PtOffsets(0-3).Y from Origin(0-3).Y to Warped Y. Relational distances only, not actual X,Y coordinates
' When path is rendered, it is rendered in this order
' 1. Warp if needed -- requires cloned path for rendering
' 2. Rotate if needed - scaling & rotating are done together
' 3. Scale if needed

' ---=====| Private Custom Enumerations |====---
Private Enum ePathState
    ps_Measured = 1                 ' Path has been measured else it hasn't
    ps_PointsCached = 2             ' See CachePathPoints property
    ps_LinearXRefCreated = 4        ' See CachePathPoints property & NavigatePath routine
    ps_ClipBox = 8                  ' Cache clipbox when first called, getting clipbox requires cloning path
    ps_TransformMatrix = 16         ' Transformation matrix cached
    ps_TransformOffsets = 32        ' Transformation is overridden
    ps_PointsCachable = 64          ' user called CachePathPoints property
End Enum

Private m_Source As PATHATTRIBUTES  ' the raw path handle
Private m_clipBox As RECTF          ' cached minimal rectangle around path. Used often
Private m_State As ePathState       ' flags indicating current state of settings and/or path
Private m_BezierPrecision As Single ' time interval for measuring bezier's default = .02 (see property)

Private m_Brushes() As clsGDIbrush  ' fill, shadow, outline brushes - user defined
Private m_BrushKeys() As String     ' keys for those brushes
Private m_BrushCount As Long        ' number brushes used when rendering

Private m_SubPaths As PATHMARKERS   ' path offsets/lengths/indexes per marked subpath

Private m_PathPoints() As POINTF    ' these 2 used when CachePathPoints=True or temporarily within class
Private m_PathPtType() As Byte      ' ^^ destroyed after use unless CachePathPoints=True
Private m_LinearXRef() As LINEARCURVEREF  ' temporary use/cache when NavigatePath is called & CachePathPoints=True

Private m_HitTestPath As Long       ' on-demand use only in IsMouseOverPath method. Destroyed when path changes size/shape/posiiton

' ---=====| Initialization |====---
Private Sub Class_Initialize()
    m_BezierPrecision = 0.02!   ' can be changed in CurveMeasurementPrecision property
End Sub

' ---=====| CLEAN UP|====---
Private Sub Class_Terminate()
    ' Note: Hitting STOP or executing END statements prevents clean up
    ' of resources, leaks, and potential crashes with GDI+
    Erase m_Brushes()
    If m_Source.handle Then GdipDeletePath m_Source.handle
    If m_HitTestPath Then GdipDeletePath m_HitTestPath
End Sub

' ---=====| Bezier Curve Measurement Precision |====---
Public Property Let CurveMeasurementPrecision(ByVal Precision As Single)
    ' prevent unreasonable precision values. Used in NavigatePath & zSpt_MeasurePath
    If Precision >= 0.001! And Precision <= 0.5 Then
        If Not Precision = m_BezierPrecision Then
            m_BezierPrecision = Precision
            m_State = ((m_State And Not ps_Measured) And Not ps_LinearXRefCreated)
            Erase m_LinearXRef
        End If
    End If
End Property
Public Property Get CurveMeasurementPrecision() As Single
    CurveMeasurementPrecision = m_BezierPrecision
End Property

' ---=====| PATH POINT CACHE SETTINGS |====---
Public Property Let CachePathPoints(ByVal bCache As Boolean)
    
    ' If path points are to be queried within a loop, caching speeds up loop
    ' The data cached can be very large, depending on complexit of the path
    If bCache Then
        m_State = (m_State Or ps_PointsCachable)
        zSpt_GetPathData m_PathPoints(), m_PathPtType(), True
    Else
        Erase m_PathPoints
        Erase m_PathPtType
        Erase m_LinearXRef()
        m_State = (((m_State And Not ps_PointsCachable) And Not ps_LinearXRefCreated) And Not ps_PointsCached)
    End If
    
End Property
Public Property Get CachePathPoints() As Boolean
    CachePathPoints = ((m_State And ps_PointsCachable) = ps_PointsCachable)
End Property

' ---=====| LENGTH PROPERTIES (READ-ONLY) |====---
Public Property Get Length(Optional ByVal MarkerIndex As Long = 0&) As Double
    ' Pass MarkerIndex if wanting length of a marked subpath
    Length = zSpt_MeasurePath(MarkerIndex)
End Property

' ---=====| PATH HANDLE (READ-ONLY) |====---
Public Property Get handle() As Long
    handle = m_Source.handle ' DO NOT destroy the handle. The class owns it
End Property

' ---=====| PATH MARKER COUNT (READ-ONLY) |====---
Public Property Get MarkerCount() As Long
    If zSpt_GetSubpathInfo(True) Then MarkerCount = m_SubPaths.Count
End Property


' ---=====| BRUSHES/PENS ASSIGNMENTS |====---
Public Property Let Brush(ByVal Key As String, newBrush As clsGDIbrush)
    Set Brush(Key) = newBrush
End Property
Public Property Set Brush(ByVal Key As String, newBrush As clsGDIbrush)

    ' Note: Multiple brushes can be used to create special effecs such as
    ' shadows, embossing, engraving, glows, etc.  Each brush requires a
    ' unique key to identify it. Brushes are always rendered in lowest
    ' zOrder to highest. Change zOrder as neededed with the BrushZOrder routine
    
    ' Passing Nothing as the newBrush deletes the Keyed brush.
    ' (i.e., Set .Brush("myKey")=Nothing )
    
    Dim I As Long
    If Not Key = vbNullString Then
        For I = 1 To m_BrushCount   ' expect few brushes so no Binary Search lookup is employed
            If StrComp(Key, m_BrushKeys(I), vbTextCompare) = 0& Then Exit For
        Next
        If I > m_BrushCount Then
            If newBrush Is Nothing Then Exit Property
            Me.Brush(Key).OffsetX = 0!  ' forces new brush to be created
        End If
        If newBrush Is Nothing Then     ' removing a brush completely
            If m_BrushCount = 1 Then
                Erase m_Brushes         ' last brush, clear array
                Erase m_BrushKeys
            Else
                Set m_Brushes(I) = Nothing      ' clear ObjPtr brush class
                m_BrushKeys(I) = vbNullString   ' clear StrPtr of key
                If I < m_BrushCount Then        ' remove brush/key & adjust array as needed
                    CopyMemory ByVal VarPtr(m_Brushes(I)), ByVal VarPtr(m_Brushes(I + 1)), (m_BrushCount - I) * 4&
                    CopyMemory ByVal VarPtr(m_BrushKeys(I)), ByVal VarPtr(m_BrushKeys(I + 1)), (m_BrushCount - I) * 4&
                    CopyMemory ByVal VarPtr(m_Brushes(m_BrushCount)), 0&, 4&
                    CopyMemory ByVal VarPtr(m_BrushKeys(m_BrushCount)), 0&, 4&
                End If
                ReDim Preserve m_Brushes(1 To m_BrushCount - 1)
                ReDim Preserve m_BrushKeys(1 To m_BrushCount - 1)
            End If
            m_BrushCount = m_BrushCount - 1     ' decrement nr of brushes
        Else
            Set m_Brushes(I) = newBrush ' changing a brush; no action required
            m_Brushes(I).UpdateGradientBounds = True
            ' ^^ set its flag so it can be updated next rendering. We don't know where the brush
            ' came from & it may be tied to path bounds or shape.
        End If
        
    End If
    
End Property
Public Property Get Brush(ByVal Key As String) As clsGDIbrush

    Dim I As Long
    If Not Key = vbNullString Then
        For I = 1 To m_BrushCount ' expect few brushes so no Binary Search lookup is employed
            If StrComp(Key, m_BrushKeys(I), vbTextCompare) = 0& Then Exit For
        Next
        If I > m_BrushCount Then            ' if brush does not exist, create it
            m_BrushCount = m_BrushCount + 1 ' and assign the key automatically
            ReDim Preserve m_Brushes(1 To m_BrushCount)
            ReDim Preserve m_BrushKeys(1 To m_BrushCount)
            Set m_Brushes(m_BrushCount) = New clsGDIbrush
            m_BrushKeys(m_BrushCount) = Key
        End If
        Set Brush = m_Brushes(I)
    End If
    
End Property

' ---=====| Brush Order |====---
Public Function BrushZOrder(ByVal Key As String, ByVal FirstBrush As Boolean) As Boolean

    ' FirstBrush :: If true, brush is used first when rendering path, else used last
    Dim I As Long, brushPtr As Long, keyPtr As Long
    If Not Key = vbNullString Then
        For I = 1 To m_BrushCount ' expect few brushes so no Binary Search lookup is employed
            If StrComp(Key, m_BrushKeys(I), vbTextCompare) = 0& Then Exit For
        Next
        If I <= m_BrushCount Then
            If FirstBrush = True Then ' moving to bottom of the array
                If I > 1 Then
                    brushPtr = ObjPtr(m_Brushes(I))
                    keyPtr = StrPtr(m_BrushKeys(I))
                    CopyMemory ByVal VarPtr(m_Brushes(2)), ByVal VarPtr(m_Brushes(1)), (I - 1) * 4&
                    CopyMemory ByVal VarPtr(m_BrushKeys(2)), ByVal VarPtr(m_BrushKeys(1)), (I - 1) * 4&
                    CopyMemory ByVal VarPtr(m_Brushes(1)), brushPtr, 4&
                    CopyMemory ByVal VarPtr(m_BrushKeys(1)), keyPtr, 4&
                End If
            ElseIf I < m_BrushCount Then ' moving to end of the array (LastBrush)
                brushPtr = ObjPtr(m_Brushes(I))
                keyPtr = StrPtr(m_BrushKeys(I))
                CopyMemory ByVal VarPtr(m_Brushes(I)), ByVal VarPtr(m_Brushes(I + 1)), (m_BrushCount - I) * 4&
                CopyMemory ByVal VarPtr(m_BrushKeys(I)), ByVal VarPtr(m_BrushKeys(I + 1)), (m_BrushCount - I) * 4&
                CopyMemory ByVal VarPtr(m_Brushes(m_BrushCount)), brushPtr, 4&
                CopyMemory ByVal VarPtr(m_BrushKeys(m_BrushCount)), keyPtr, 4&
            End If
            BrushZOrder = True
        End If
    End If

End Function

' ---=====| Brush Count |====---
Public Property Get BrushCount() As Long
    BrushCount = m_BrushCount
End Property

' ---=====| Brush Keys |====---
Public Function BrushKeys(brushNames() As String) As Long
    ' return Brush keys in ZOrder from lowest to highest
    ' The return value is the number of brushes returned
    ' and the LBound of the array is always zero
    If m_BrushCount Then
        ReDim brushNames(0 To m_BrushCount - 1)
        For m_BrushCount = 0 To m_BrushCount - 1
            brushNames(m_BrushCount) = m_BrushKeys(m_BrushCount + 1)
        Next
    End If
End Function

' ---=====| Path Boundaries -- Visual as seen on DC |====---
Public Property Get DisplayTop() As Single
    DisplayTop = m_Source.DisplayLocale.Y
End Property
Public Property Get DisplayLeft() As Single
    DisplayLeft = m_Source.DisplayLocale.X
End Property
Public Property Get DisplayWidth() As Single
    ' negative value indicates width is mirrored and left > right
    Dim tSize As Single
    zSpt_GetBoundingBox True, 0!, 0!, tSize, 0!
    DisplayWidth = tSize * m_Source.Scaler.X
End Property
Public Property Get DisplayHeight() As Single
    ' negative value indicates height is mirrored and top > bottom
    Dim tSize As Single
    zSpt_GetBoundingBox True, 0!, 0!, 0!, tSize
    DisplayHeight = tSize * m_Source.Scaler.Y
End Property

' ---=====| Path Boundaries -- Actual before any offsets/scales/etc are applied |====---
' The Clipped parameter when True, uses the tightest boundary around a path
'   but when False uses the actual boundary returned from GdipGetPathWorldBounds
' Note: Calling GetBoundingBox will return all values in one call

Public Property Get RawLeft(Optional ByVal Clipped As Boolean) As Single
    zSpt_GetBoundingBox Clipped, RawLeft, 0!, 0!, 0!
End Property
Public Property Get RawTop(Optional ByVal Clipped As Boolean) As Single
    zSpt_GetBoundingBox Clipped, 0!, RawTop, 0!, 0!
End Property
Public Property Get RawWidth(Optional ByVal Clipped As Boolean) As Single
    zSpt_GetBoundingBox Clipped, 0!, 0!, RawWidth, 0!
End Property
Public Property Get RawHeight(Optional ByVal Clipped As Boolean) As Single
    zSpt_GetBoundingBox Clipped, 0!, 0!, 0!, RawHeight
End Property

' ---=====| Rotation Option |====---
Public Property Get Angle() As Single
    Angle = m_Source.Rotation
End Property
Public Property Let Angle(ByVal theAngle As Single)
    theAngle = (Int(theAngle) Mod 360) + theAngle - Int(theAngle)
    If Not theAngle = m_Source.Rotation Then
        m_Source.Rotation = theAngle
        m_State = (m_State And Not ps_TransformMatrix)
    End If
End Property

' ---=====| Scale Options |====---
Public Property Get ScaleX() As Single
    ScaleX = m_Source.Scaler.X
End Property
Public Property Let ScaleX(ByVal theScale As Single)
    If Not m_Source.Scaler.X = theScale Then
        m_Source.Scaler.X = theScale
        m_State = (m_State And Not ps_TransformMatrix)
    End If
End Property
Public Property Get ScaleY() As Single
    ScaleY = m_Source.Scaler.Y
End Property
Public Property Let ScaleY(ByVal theScale As Single)
    If Not m_Source.Scaler.Y = theScale Then
        m_Source.Scaler.Y = theScale
        m_State = (m_State And Not ps_TransformMatrix)
    End If
End Property

Public Function ClonePath(Optional ByVal MarkerIndex As Long, _
                        Optional ByVal WithTransformations As Boolean = False) As Long

    ' ---=====| Returns a path clone |====---
    ' Destroying the path with GdipDeletePath is your responsibility
    
    '   MarkerIndex :: If non-zero, a subpath will be cloned vs the entire path
    '   WithTransformations :: applies any scale/rotations/offset transformations to clone

    Dim pClone As Long
    If Not m_Source.handle = 0& Then
        zSpt_GetSubpathInfo True
        If (MarkerIndex = 0&) Or (MarkerIndex = 1& And m_SubPaths.Count = MarkerIndex) Then
            ' entire path cloned or no marked paths
            If WithTransformations Then ' return as rendered
                pClone = zSpt_TransformPath(True, False)
            Else
                GdipClonePath m_Source.handle, pClone
            End If
        Else
            pClone = zSpt_CreateSubPathPath(MarkerIndex, False, WithTransformations)
        End If
    End If
    ClonePath = pClone
        
End Function

Public Sub ClearMarkers()
    ' ---=====| Clears any existing path markers |====---
    If m_Source.handle Then
        GdipClearPathMarkers m_Source.handle
        If m_SubPaths.Count > 1 Then m_SubPaths.Count = 0
        Erase m_SubPaths.Markers
    End If
End Sub

Public Sub Clear()

    ' ---=====| Deletes an existing path completely and its brushes |====---

    If m_Source.handle Then
        GdipDeletePath m_Source.handle
        m_Source.handle = 0&
    End If
    If m_HitTestPath Then
        GdipDeletePath m_HitTestPath
        m_HitTestPath = 0&
    End If
    Erase m_PathPoints
    Erase m_PathPtType
    Erase m_LinearXRef
    With m_SubPaths
        Erase .Markers
        .Count = 0&
    End With
    Erase m_Source.ptOffsets
    Erase m_Source.TransOverride
    
    Erase m_Brushes
    Erase m_BrushKeys
    m_BrushCount = 0&
    
    m_State = 0&
    m_Source.Rotation = 0!
    m_Source.Scaler.X = 0!
    m_Source.Scaler.Y = 0!

End Sub

Public Sub Reset(Optional ByVal FillType As ePathBrushFillMode)
    
    ' ---=====| Zeroizes the path and keeps brush data |====---
    
    If Not FillType = ePathBrushFillMode.FillModeAlternate Then
        FillType = ePathBrushFillMode.FillModeWinding
    End If
    If m_Source.handle Then
        GdipResetPath m_Source.handle
        GdipSetPathFillMode m_Source.handle, FillType
    Else
        GdipCreatePath FillType, m_Source.handle
    End If
    If m_HitTestPath Then
        GdipDeletePath m_HitTestPath
        m_HitTestPath = 0&
    End If
    Erase m_Source.Origin
    Erase m_Source.ptOffsets
    Erase m_SubPaths.Markers
    m_SubPaths.Count = 0&
    Erase m_LinearXRef

    m_Source.Scaler.X = 1!: m_Source.Scaler.Y = 1!
    m_Source.Rotation = 0!
    m_Source.WarpStyle = 0&
    
End Sub

Public Function Append_Arc(ByVal X As Single, ByVal Y As Single, _
                    ByVal arcWidth As Single, ByVal arcHeight As Single, _
                    ByVal StartAngle As Single, ByVal ArcAngle As Single, _
                    Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds an arc to a path. If no path exists, it is created |====---
    ' Notes
    '   - For an upwards arc, make arcAngle positive: 180
    '   - For a downards arc, make arcAngle negative: -180
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathArc(m_Source.handle, X, Y, arcWidth, arcHeight, StartAngle, ArcAngle) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_Arc = True
    End If
    zSpt_DoPostUpdateActions
    
End Function

Public Function Append_Bezier(ByVal X1 As Single, ByVal Y1 As Single, _
                    ByVal CtrlPtX1 As Single, ByVal CtrlPtY1 As Single, _
                    ByVal CtrlPtX2 As Single, ByVal CtrlPtY2 As Single, _
                    ByVal X2 As Single, ByVal Y2 As Single, _
                    Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a bezier curve to a path. If no path exists, it is created |====---
    ' Note: See Append_Wave for a more user friendly version
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathBezier(m_Source.handle, X1, Y1, CtrlPtX1, CtrlPtY1, CtrlPtX2, CtrlPtY2, X2, Y2) = 0& Then
        Append_Bezier = True
        If AppendMarker Then GdipSetPathMarker m_Source.handle
    End If
    zSpt_DoPostUpdateActions
End Function

Public Function Append_BeziersPTR(ByVal bzPointsPointer As Long, ByVal bzPointCount As Long, _
                                Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds multiple bezier curves to a path. If no path exists, it is created |====---
    ' Note: this function expects a pointer to an array of points. The array can be Singles, PointF
    ' or anything else that equates to a pair of Single values for each point coordinate.
    ' Pass bzPointsPointer as:  VarPtr(myArray(n))  where n is the 1st coordinate's X coordinate
    ' If using a Singles array, declare your array as: myArray(0 to 1, 0 to nrPts-1) & 1st dimension are X,Y respectively
    
    ' Caution: The pointer passed must be valid else a crash will occur
    
    If bzPointsPointer Then
        If m_Source.handle = 0& Then Me.Reset
        If GdipAddPathBeziers(m_Source.handle, ByVal bzPointsPointer, bzPointCount) = 0& Then
            If AppendMarker Then GdipSetPathMarker m_Source.handle
            Append_BeziersPTR = True
        End If
        zSpt_DoPostUpdateActions
    End If

End Function

Public Function Append_CurvePTR(ByVal cvPointsPointer As Long, ByVal cvPointCount As Long, _
                                Optional Tension As Single = 0&, Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds curve to a path. If no path exists, it is created |====---
    ' The curve will pass thru each point contained in array at cvPointsPointer
    ' Tension is a postive value that controls the curve and how the curve bends.
    '   -- A value of 0 specifies that the spline is a sequence of straight line segments. As the value increases, the curve becomes fuller
    
    ' Note: this function expects a pointer to an array of points. The array can be Singles, PointF or anything else that
    ' equates to a pair of Single values for each point coordinate.
    ' Pass cvPointsPointer as:  VarPtr(myArray(n))  where n is the 1st coordinate's X coordinate
    ' If using a Singles array, declare your array as: myArray(0 to 1, 0 to nrPts-1) & 1st dimension are X,Y respectively
    
    ' Caution: The pointer passed must be valid else a crash will occur
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathCurve2(m_Source.handle, ByVal cvPointsPointer, cvPointCount, Tension) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_CurvePTR = True
    End If
    zSpt_DoPostUpdateActions
End Function

Public Function Append_Ellipse(ByVal X As Single, ByVal Y As Single, _
                            ByVal Width As Single, ByVal Height As Long, _
                            Optional ByVal AppendMarker As Boolean = False) As Boolean
                    
    ' ---=====| Adds a circle to a path. If no path exists, it is created |====---
    
    ' Note. Think of a clock
    '   StartAngle=180 then Clock hand = 9
    '   StartAngle=0 then Clock hand = 3
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathEllipse(m_Source.handle, X, Y, Width, Height) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_Ellipse = True
    End If
    zSpt_DoPostUpdateActions

End Function

Public Function Append_Line(ByVal X1 As Single, ByVal Y1 As Single, _
                            ByVal X2 As Single, ByVal Y2 As Single, _
                            Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a line to a path. If no path exists, it is created |====---
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathLine(m_Source.handle, X1, Y1, X2, Y2) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_Line = True
    End If
    zSpt_DoPostUpdateActions

End Function

Public Function Append_LinesPTR(ByVal lnPointsPointer As Long, ByVal lnPointCount As Long, _
                                Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a multiple segment line to a path. If no path exists, it is created |====---
    ' the 1st set of singles are the starting the point, the last is the ending point, each between are the multi-segment points
    
    ' Note: this function expects a pointer to an array of points. The array can be Singles, PointF or anything else that
    ' equates to a pair of Single values for each point coordinate.
    ' Pass lnPointsPointer as:  VarPtr(myArray(n))  where n is the 1st coordinate's X coordinate
    ' If using a Singles array, declare your array as: myArray(0 to 1, 0 to nrPts-1) & 1st dimension are X,Y respectively
    
    ' Caution: The pointer passed must be valid else a crash will occur
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathLine2(m_Source.handle, ByVal lnPointsPointer, lnPointCount) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_LinesPTR = True
    End If
    zSpt_DoPostUpdateActions

End Function

Public Function Append_Path(ByVal gdiPlus_Path As Long, ByVal ConnectToAnyExistingPath As Boolean, _
                            Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a passed GDI+ path to a path. If this class path does not exist, it is created |====---
    
    Dim nrPts As Long
    GdipGetPointCount gdiPlus_Path, nrPts
    If nrPts > 1 Then
        If m_Source.handle = 0& Then Me.Reset
        If GdipAddPathPath(m_Source.handle, gdiPlus_Path, Abs(ConnectToAnyExistingPath)) = 0& Then
            If AppendMarker Then GdipSetPathMarker m_Source.handle
            Append_Path = True
        End If
        zSpt_DoPostUpdateActions
    End If
    
End Function

Public Function Append_Pie(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, _
                            ByVal StartAngle As Single, ByVal ArcAngle As Single, _
                            Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a pie/wedge to a path. If the path does not exist, it is created |====---
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathPie(m_Source.handle, X, Y, Width, Height, StartAngle, ArcAngle) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_Pie = True
    End If
    zSpt_DoPostUpdateActions
    
End Function

Public Function Append_PolygonPTR(ByVal lnPointsPointer As Long, ByVal lnPointCount As Long, _
                                Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a multiple segment polygon to a path. If no path exists, it is created |====---
    ' This function & Append_LinesPTR are very similar, however, this function automatically closes the path
    
    ' Note: this function expects a pointer to an array of points. The array can be Singles, PointF or anything else that
    ' equates to a pair of Single values for each point coordinate.
    ' Pass lnPointsPointer as:  VarPtr(myArray(n))  where n is the 1st coordinate's X coordinate
    ' If using a Singles array, declare your array as: myArray(0 to 1, 0 to nrPts-1) & 1st dimension are X,Y respectively
    
    ' Caution: The pointer passed must be valid else a crash will occur
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathPolygon(m_Source.handle, ByVal lnPointsPointer, lnPointCount) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_PolygonPTR = True
    End If
    zSpt_DoPostUpdateActions

End Function

Public Function Append_Rectangle(ByVal X As Single, ByVal Y As Single, _
                                ByVal Width As Single, ByVal Height As Single, _
                                Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Adds a rectangle to a path. If the path does not exist, it is created |====---
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathRectangle(m_Source.handle, X, Y, Width, Height) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_Rectangle = True
    End If
    zSpt_DoPostUpdateActions
    
End Function

Public Function Append_Wave(ByVal X1 As Single, ByVal Y1 As Single, _
                            ByVal X2 As Single, ByVal Y2 As Single, ByVal CrestHeight As Long, _
                            Optional ByVal WaveCount As Long = 1, Optional ByVal AppendMarker As Boolean = False) As Boolean
    
    ' ---=====| Creates waves. If the path does not exist, it is created |====---
    ' not a standard GDI+ call, but rather a series of calculated Beziers
    
    Dim ptWave() As POINTF
    Dim I As Long
    Dim Cx As Single, Cy As Single
    Dim T As Single, p As Single, midP As Single
    
    If WaveCount < 1 Then Exit Function
    On Error GoTo ExitRoutine
    
    Cx = X2 - X1
    Cy = Y2 - Y1
    p = 1! / WaveCount          ' scale per wave
    midP = p * 0.5!             ' mid point offset for bezier control point
    
    WaveCount = WaveCount * 3   ' 3 points per wave
    ReDim ptWave(0 To WaveCount) ' plus a starting point
    
    ' set first point in curve
    ptWave(0).X = X1: ptWave(0).Y = Y1
    
    For I = 1 To WaveCount Step 3
        T = T + p
        ' :: Control Points #1 & #2
        If X1 = X2 Then ' vertical line, waves are horizontal/parallel to line
            ptWave(I).X = X1 + CrestHeight * 3
            ptWave(I).Y = Y1 + (T - midP) * Cy      ' control point is mid point
            ptWave(I + 1).X = X1 - CrestHeight * 3
            ptWave(I + 1).Y = ptWave(I).Y           ' same for 2nd control point
        Else
            ptWave(I).X = X1 + (T - midP) * Cx      ' control point is mid point
            ptWave(I).Y = Y1 + CrestHeight * 3 + 0.5! * Cy
            ptWave(I + 1).X = ptWave(I).X           ' same for 2nd control point
            ptWave(I + 1).Y = Y1 - CrestHeight * 3 + 0.5! * Cy
        End If
        ' :: end point of this wave & start point of next wave, if any
        ptWave(I + 2).X = X1 + T * Cx
        ptWave(I + 2).Y = Y1 + T * Cy
    Next
    ' set final point in curve. Should be set, but rounding could have changed it by a fraction
    ptWave(WaveCount).X = X2: ptWave(WaveCount).Y = Y2
    
    If m_Source.handle = 0& Then Me.Reset
    If GdipAddPathBeziers(m_Source.handle, ptWave(0), WaveCount + 1) = 0& Then
        If AppendMarker Then GdipSetPathMarker m_Source.handle
        Append_Wave = True
    End If
    zSpt_DoPostUpdateActions

ExitRoutine:
End Function

Public Function Append_String(ByVal X As Single, ByVal Y As Single, _
                            ByVal theString As String, ByVal FontName As String, _
                            ByVal FontSize As Single, _
                            Optional fontStyle As eFontStyle = FontStyleRegular) As Boolean

    ' ---=====| Adds text to a path. If the path does not exist, it is created |====---
    ' Note about markers.  When GDI+ adds a string to a path, each character is
    ' written with a marker at the end of the character's subpaths. Should you want
    ' to add a string without markers then you can accomplish this using the following:
    '   1. Create a new path class
    '   2. Call this function in that other class
    '   3. Call that other class' ClearMarkers routine
    '   4. Call this class' Append_Path function, passing the other class' .Handle property
    '       -- include the AppendMarker parameter to that call if desired
    '   5. Destroy the other class
    
    Dim fontFam As Long, fRect As RECTF
    
    If Len(theString) Then
        If m_Source.handle = 0& Then Me.Reset
        Call GdipCreateFontFamilyFromName(StrPtr(FontName), 0, fontFam)
        If fontFam Then
            fRect.Left = X: fRect.Top = Y
            If GdipAddPathString(m_Source.handle, StrPtr(theString), -1, fontFam, (fontStyle And &HF), FontSize, fRect, 0&) = 0& Then
                Append_String = True
            End If
            GdipDeleteFontFamily fontFam
        End If
        zSpt_DoPostUpdateActions
    End If

End Function

Public Function ClosePath(Optional ByVal CloseAllPaths As Boolean = False) As Boolean
    
    ' ---=====| Closes an open path - joins last point with first point |====---
    
    Dim bReturn As Boolean
    If m_Source.handle Then
        If CloseAllPaths Then
            bReturn = (GdipClosePathFigures(m_Source.handle) = 0&)
        Else
            bReturn = (GdipClosePathFigure(m_Source.handle) = 0&)
        End If
        If bReturn Then zSpt_DoPostUpdateActions
    End If
    ClosePath = bReturn
    
End Function

Public Function ApplyTransformations() As Boolean

    ' ---=====| Permanently rewrites path to set scales/rotation/offsets |====---
    ' As mentioned at top of this class in the remarks, transformations to the path
    ' are never applied directly in order to provide best quality renderings when
    ' the path can change scales/rotations frequently. If the path should be
    ' rewritten to include the transformations, call this function. Calling this
    ' function resets all properties to default: scales, rotation, offsets, warp
    
    ' To return a path handle that will include transformation but not permanently
    ' rewrite this path, call the ClonePath function
    
    Dim mx As Long, curDisplayPts As POINTF
    curDisplayPts = m_Source.DisplayLocale
    m_Source.DisplayLocale.X = 0!: m_Source.DisplayLocale.Y = 0!
    m_State = (m_State And Not ps_TransformMatrix)
    mx = zSpt_TransformPath(False, True)
    m_Source.DisplayLocale = curDisplayPts
    If Not mx = 0& Then
        GdipTransformPath m_Source.handle, mx
        GdipDeleteMatrix mx
        zSpt_DoPostUpdateActions
        m_Source.Scaler.X = 1!
        m_Source.Scaler.Y = 1!
        m_Source.Rotation = 0!
        ApplyTransformations = True
    End If
End Function

Public Sub OverrideWorldTransform(ByVal SourceCtrX As Single, ByVal SourceCtrY As Single, _
                                 ByVal DestCtrX As Single, ByVal DestCtrY As Single, _
                                 ByVal ScaledLeft As Single, ByVal ScaledTop As Single)

    ' ---=====| Overrides point reference for scaling, offsetting, rotation |====---
    
    ' Supply this function with the necessary information for scaling and rotating at
    ' different location/axis than the path's own world coordinates. For example,
    ' if you have multiple overlapped paths that are not the same size and/or
    ' not centered over each other exactly, each will rotate on their own axis
    ' or scale to their own left/top coordinates. This function can override the normal
    ' behavior. Also see: OverrideWorldTransformReset() & zSpt_TransformPath
    
    ' Parameters:
    '   :: SourceCtrX, SourceCtrY are the source rotation coordinates in relation to this path
    '       Note: Default is to rotate from center of Path, regardless of screen location
    '   :: DestCtrX, DestCtrY are target rotation coordinates where rotation is to occur
    '       Note: Default is to rotate to center of displayed path on screen
    '   :: ScaledLeft, ScaledTop are left/top coordinates where scaled path should be rendered to
    '       Note: Default is the left/top coordinates of displayed path on screen
    
    m_State = ((m_State And Not ps_TransformMatrix) Or ps_TransformOffsets)
    With m_Source
        ReDim .TransOverride(0 To 5)
        .TransOverride(0) = SourceCtrX
        .TransOverride(1) = SourceCtrY
        .TransOverride(2) = DestCtrX
        .TransOverride(3) = DestCtrY
        .TransOverride(4) = ScaledLeft
        .TransOverride(5) = ScaledTop
    End With
    
End Sub

Public Sub OverrideWorldTransformReset()

    ' ---=====| Resets point reference for scaling, offsetting, rotation |====---
    ' Resets any applied override values via the OverrideWorldTransform method
    
    m_State = ((m_State And Not ps_TransformOffsets) And Not ps_TransformMatrix)
    Erase m_Source.TransOverride()
    
End Sub

Public Function GetPathData_PTR(Optional ByVal PointsPTR As Long, Optional ByVal PointTypesBytePtr As Long) As Long

    ' ---=====| Returns the path's point coordinates and types |====---
    
    ' If PointsPtr & PointTypesBytePtr are zero, then the function returns the number of
    ' points in the path which is to be used to dimension your array(s)
    
    ' If PointsPtr is non-zero, it must be a valid pointer to an array large enough to hold the
    '   Point structure: 8 bytes per point
    ' Pass this pointer as VarPtr(myPoints(n)) where n is usually the LBound of the sized array
    '   -- If your array is 2-dimensional then: VarPtr(myPoints(n,n))
    
    ' If PointTypesBytePtr is non-zero, the point types will be returned. It must be  a valid
    '   pointer to a byte array large enough to hold the data: 1 byte per point
    ' Pass this pointer as VarPtr(myBytes(n)) where n is usually the LBound of the sized array
    
    ' The return function is always the amount of points returned
    ' Caution: The pointer(s) passed must be valid else a crash will occur
    
    Dim pd As PATHDATA
    If m_Source.handle Then
        GdipGetPointCount m_Source.handle, pd.Count
        If pd.Count > 1 Then
            If Not (PointsPTR = 0& And PointTypesBytePtr = 0&) Then
                If PointTypesBytePtr = 0& Then 'return just the path points
                    GdipGetPathPoints m_Source.handle, PointsPTR, pd.Count
                ElseIf PointsPTR = 0& Then ' return just the path types
                    GdipGetPathTypes m_Source.handle, PointTypesBytePtr, pd.Count
                Else    ' return points & types
                    pd.PointsPTR = PointsPTR
                    pd.TypesPtr = PointTypesBytePtr
                    GdipGetPathData m_Source.handle, pd
                End If
            End If
        End If
    End If
    GetPathData_PTR = pd.Count
    
End Function

Public Function Move(ByVal X As Single, ByVal Y As Single, _
                    Optional asTransformed As Boolean = False, _
                    Optional MoveOnCenter As Boolean = False) As Boolean

    ' ---=====| Moves the path to passed X,Y values |====---
    ' Return value is false if no path points exist
    
    ' Note: Rendering with thick pens, special pen caps, etc may result in the
    ' rendered path slightly left/top of X,Y. Take this into consideration and
    ' append offset values to X,Y before calling this function. In this case
    ' recommend offseting X,Y by the pen thickness & its offset if any
    
    If m_Source.handle Then
        If MoveOnCenter Then ' X,Y are center points; center path over X,Y
            zSpt_GetBoundingBox True, 0!, 0!, 0!, 0!
            m_Source.DisplayLocale.X = X - (m_clipBox.Width * m_Source.Scaler.X) / 2
            m_Source.DisplayLocale.Y = Y - (m_clipBox.Height * m_Source.Scaler.Y) / 2
        
        ElseIf asTransformed = False Then
            m_Source.DisplayLocale.X = X
            m_Source.DisplayLocale.Y = Y
        
        Else    '  asTransformed
'            Dim transCtr As POINTF, targetCtr As POINTF, Radian As Single
'            Dim Cx As Single, Cy As Single
'            zSpt_GetBoundingBox True, 0!, 0!, 0!, 0!
'            zSpt_TransformPath False, False
'            With m_Source
''                targetCtr.X = X
''                targetCtr.Y = Y
'                Cx = ((m_clipBox.Width * .Scaler.X) / 2)
'                Cy = ((m_clipBox.Height * .Scaler.Y) / 2)
'                targetCtr.X = ((m_clipBox.Left * .Transformation(0)) + (m_clipBox.Top * .Transformation(2)) + .Transformation(4))
'                targetCtr.Y = ((m_clipBox.Left * .Transformation(1)) + (m_clipBox.Top * .Transformation(3)) + .Transformation(5))
'                Radian = .Rotation * (Atn(1) * 4) / 180
'                targetCtr.X = targetCtr.X + Cos(Radian) * Cx
'                targetCtr.Y = targetCtr.Y + Sin(Radian) * Cx
''
''                frmWordArt.picCanvas.Line (targetCtr.X - 4, targetCtr.Y - 4)-(targetCtr.X + 4, targetCtr.Y + 4), vbRed, B
''                frmWordArt.picCanvas.Circle (transCtr.X, transCtr.Y), 4, vbRed
''
''                targetCtr.X = transCtr.X - Cos(0) * Cx - Cx
''                targetCtr.Y = transCtr.Y - Sin(0) * Cy - Cy
'
'
'
'                'transCtr.X = ((.DisplayLocale.X * .Transformation(0)) + (.DisplayLocale.Y * .Transformation(2)) + .Transformation(4))
'                'transCtr.Y = ((.DisplayLocale.X * .Transformation(1)) + (.DisplayLocale.Y * .Transformation(3)) + .Transformation(5))
'            End With
'            m_Source.DisplayLocale.X = X - m_Source.DisplayLocale.X
'            m_Source.DisplayLocale.Y = Y - m_Source.DisplayLocale.Y
        End If
        
        m_State = (m_State And Not ps_TransformMatrix)
        Move = True
    
    End If


End Function

Public Function Offset(ByVal X As Single, ByVal Y As Single) As Boolean

    ' ---=====| Offsets the path by passed X,Y values |====---
    ' Return value is false if no path points exist

    ' Note: Rendering with thick pens, special pen caps, etc may result in the
    ' rendered path slightly left/top of X,Y. Take this into consideration and
    ' append offset values to X,Y before calling this function
    
    If m_Source.handle Then
        m_Source.DisplayLocale.X = m_Source.DisplayLocale.X + X
        m_Source.DisplayLocale.Y = m_Source.DisplayLocale.Y + Y
        m_State = (m_State And Not ps_TransformMatrix)
        Offset = True
    End If

End Function

Public Function GetBoundingBox(ByVal UseClipBox As Boolean, _
                                ByRef Left As Single, ByRef Top As Single, _
                                ByRef Width As Single, ByRef Height As Single, _
                                Optional ByVal MarkerIndex As Long = 0&) As Boolean
                                
    ' ---=====| Returns the dimensions of a bounding box containing the path as Single values |====---
    GetBoundingBox = zSpt_GetBoundingBox(UseClipBox, Left, Top, Width, Height, MarkerIndex)
        
End Function

Public Function GetBoundingBox_Long(ByVal UseClipBox As Boolean, _
                                ByRef Left As Long, ByRef Top As Long, _
                                ByRef Width As Long, ByRef Height As Long, _
                                Optional ByVal MarkerIndex As Long = 0&) As Boolean
                                    
    ' ---=====| Returns the dimensions of a bounding box containing the path as Long values |====---
    ' Returns Long values vs Single values
    Dim X As Single, Y As Single, Cx As Single, Cy As Single
    GetBoundingBox_Long = GetBoundingBox(UseClipBox, X, Y, Cx, Cy)
    Left = CLng(X): Top = CLng(Y)
    Width = CLng(Cx): Height = CLng(Cy)
                                
End Function

Public Function GetBoundingPoints_PTR(ByVal UseClipBox As Boolean, ByVal UseTransformations As Boolean, ByVal PointsPTR As Long) As Long
    
    ' ---=====| Returns the bounds as points and optionally transforms the points |====---
    
    ' Note: this function expects a pointer to an array of points. The array can be Singles, PointF or anything else that
    ' equates to a pair of Single values for each point coordinate.
    ' Pass lnPointsPointer as:  VarPtr(myArray(n))  where n is the 1st coordinate's X coordinate
    ' If using a Singles array, declare your array as: myArray(0 to 1, 0 to nrPts-1) & 1st dimension are X,Y respectively
    
    ' Caution: The pointer passed must be valid else a crash will occur
    
    Dim thePts(0 To 3) As POINTF, p As Long, tmpPT As POINTF
    If Not PointsPTR = 0& Then
        If zSpt_TransformPath(False, False) = p Then Exit Function
        zSpt_GetBoundingBox UseClipBox, thePts(0).X, thePts(0).Y, tmpPT.X, tmpPT.Y
        ' transpose rect structure to points structure
        thePts(1).X = thePts(0).X + tmpPT.X - 1: thePts(1).Y = thePts(0).Y
        thePts(2).X = thePts(1).X: thePts(2).Y = thePts(0).Y + tmpPT.Y - 1
        thePts(3).X = thePts(0).X: thePts(3).Y = thePts(2).Y
        ' now transform the points
        If UseTransformations Then
            With m_Source
                For p = 0 To 3
                    tmpPT = thePts(p)
                    thePts(p).X = ((tmpPT.X * .Transformation(0)) + (tmpPT.Y * .Transformation(2)) + .Transformation(4))
                    thePts(p).Y = ((tmpPT.X * .Transformation(1)) + (tmpPT.Y * .Transformation(3)) + .Transformation(5))
                Next
            End With
        End If
        CopyMemory ByVal PointsPTR, thePts(0), 32&
        GetBoundingPoints_PTR = 4&
    End If
        
        
End Function

Public Function IsMouseOverPath(ByVal hDC As Long, ByVal X As Single, ByVal Y As Single) As Boolean

    ' ---=====| Determine Hit Test on a path |====---

    Dim lHit As Long, hGraphics As Long
    If Not hDC = 0& Then
        If m_HitTestPath = 0& Then m_HitTestPath = Me.ClonePath(, True)
        If Not m_HitTestPath = 0& Then
            GdipCreateFromHDC hDC, hGraphics
            GdipIsVisiblePathPoint m_HitTestPath, X, Y, hGraphics, lHit
            GdipDeleteGraphics hGraphics
        End If
    End If
    IsMouseOverPath = (lHit <> 0&)
    
End Function

Public Function GetFirstDisplayPoint(X As Single, Y As Single) As Boolean
    
    ' ---=====| Returns the X,Y coordinates of the first point of a transformed path |====---
    
    If m_Source.handle Then
        Dim pt As POINTF
        If zSpt_TransformPath(False, False) <> 0& Then
            NavigatePath 0, X, Y, , True
            GetFirstDisplayPoint = True
        End If
    End If
End Function

Public Function GetLastDisplayPoint(X As Single, Y As Single) As Boolean
    
    ' ---=====| Returns the X,Y coordinates of the first point of a transformed path |====---
    
    If m_Source.handle Then
        Dim pt As POINTF
        If zSpt_TransformPath(False, False) <> 0& Then
            NavigatePath 1, X, Y, , True
            GetLastDisplayPoint = True
        End If
    End If
End Function

Public Function NavigatePath(ByVal Distance As Double, _
                            ByRef X As Single, ByRef Y As Single, _
                            Optional ByVal MarkerIndex As Long = 0&, _
                            Optional UseTransformations As Boolean) As Boolean
    
    ' ---=====| Determine points on path for relative & specific X,Y coordinates |====---
    ' NOTE: If this is to be called in loop, set CachePathPoints=True first and when done, CachePathPoints=False
    
    ' Pass Distance as length of travel from start of path.  The beginning of the path is
    ' always zero. Distance is measured in pixels
    ' X & Y when returned will be the X,Y coordinates where distance falls on the path
    ' -- If the distance exceeds the path length, function returns false
    '    If distance is < 0 or no path yet created, function returns false
    ' MarkerIndex if non-zero will navigate only the subpath identified by MarkerIndex
    ' UseTransformations will return X,Y as path is drawn on a DC else X,Y are in relation
    '   to the 1st path point in the path/subpath being navigated

    ' this is an estimated distance on bezier.
    ' Bezier formulas calculate time vs distance. This results in points being
    ' distributed unevenly along the curve; more points on sharp bends and less
    ' on the smoother splines/curves. So, 50% in time may be less/greater than
    ' 50% of the actual curve distance
    
    ' Here I chose to use a time to distance ratio and estimating bezier X,Y distance
    ' in relation to the amount of time traveled.  The calculations appear accurate
    ' within a pixel fraction
    
    Dim pSeg(0 To 8) As POINTF
    Dim tSquared As Single, tCubed As Single
    Dim z As Long, s As Long, T As Double
    Dim b As Single, p As Long
    Dim bReturn As Boolean
    Dim timeInterval As Long
    
    T = zSpt_MeasurePath(MarkerIndex)
    If T = 0! Then Exit Function
    
    ' validate distance requested is within size of overall path
    If Distance <= 0! Then
        If MarkerIndex = 0& Then MarkerIndex = 1& ' use start of path else start of Marker
        pSeg(8) = m_PathPoints(m_SubPaths.Markers(MarkerIndex).StartIndex)
        z = -1 ' flag indicating not to process
        bReturn = (Distance = 0!)
    End If
        
    If Distance >= T Then ' exceeded length, use last point & return False to indicate end of path
        If MarkerIndex = 0& Then MarkerIndex = m_SubPaths.Count ' use end of entire path else end of Maker
        pSeg(8) = m_PathPoints(m_SubPaths.Markers(MarkerIndex).EndIndex)
        z = -1 ' flag indicating not to process
        bReturn = (Distance = T)
    End If
    
    If z = 0& Then
        ' determine which segment the Distance value falls on
        ' When each segment was calculated its accumulative distance from zero is stored in .X
        If MarkerIndex = 0& Then            ' using entire path vs Marked subpaths
            If m_SubPaths.Count = 1 Then    ' only one marked path, therefore, entire path
                MarkerIndex = 1
            Else                            ' marked paths exist, locate marked path to use
                For p = 1 To m_SubPaths.Count
                    ' move thru the marked paths, subtracting distance traveled along the way
                    b = m_SubPaths.Markers(p).Segments(UBound(m_SubPaths.Markers(p).Segments)).X
                    If Distance <= b Then Exit For ' distance falls on this marker
                    Distance = Distance - b
                Next
                ' handle possibility of rounding errors
                If p > m_SubPaths.Count Then MarkerIndex = m_SubPaths.Count Else MarkerIndex = p
            End If
        End If
        ' locate the segment within the subpath where the distance falls on
        With m_SubPaths.Markers(MarkerIndex)
            For z = 1 To UBound(.Segments)
                If Distance <= .Segments(z).X Then Exit For
            Next
            ' handle rounding errors; possible for z to exceed UBound(.Segments) by one
            If z > UBound(.Segments) Then z = UBound(.Segments)
        End With
        
        ' get the path data if not already cached
        zSpt_GetPathData m_PathPoints(), m_PathPtType(), False
        
        ' determine start, end & control points of the segment to be used
        ' and precalc some formula constants
        With m_SubPaths.Markers(MarkerIndex)
            If (m_PathPtType(.Segments(z - 1).Y) And PathPointTypeCloseSubpath) = PathPointTypeCloseSubpath _
                Or (m_PathPtType(.Segments(z).Y) And PathPointTypePathTypeMask) = PathPointTypeLine Then
                ' convert lines to flat beziers so constant formula can be used below
                If (m_PathPtType(.Segments(z - 1).Y) And PathPointTypeCloseSubpath) = PathPointTypeCloseSubpath Then
                    pSeg(0) = m_PathPoints(.Segments(z - 1).Y)
                    pSeg(1) = m_PathPoints(.Segments(z - 1).Y)
                Else
                    pSeg(0) = m_PathPoints(.Segments(z).Y - 1)
                    pSeg(1) = m_PathPoints(.Segments(z).Y - 1)
                End If
                pSeg(2) = m_PathPoints(.Segments(z).Y)
                pSeg(3) = m_PathPoints(.Segments(z).Y)
            Else    ' bezier curve, copy 4 points
                CopyMemory pSeg(0), m_PathPoints(.Segments(z).Y - 3), 32&
            End If
        End With
        For p = 5 To 7                      ' cache precalc'd bezier multipliers
            pSeg(p).X = pSeg(p - 5).X * 3   ' pts(0,1,2) * 3
            pSeg(p).Y = pSeg(p - 5).Y * 3
        Next
        pSeg(4) = pSeg(0)                   ' cache moving starting point
            
        ' do we need to initialize lookup table(s)?
        If (m_State And ps_LinearXRefCreated) = 0& Then
            m_State = (m_State Or ps_LinearXRefCreated)
            ReDim m_LinearXRef(1 To m_SubPaths.Count)
        End If
        timeInterval = 1 / m_BezierPrecision ' number of measurements required for timeInterval

        ' if not done for this segment, we need to build a time to distance cross-reference
        If z <> m_LinearXRef(MarkerIndex).Index Then
            ReDim m_LinearXRef(MarkerIndex).Distance(0 To timeInterval)
            s = 1
            m_LinearXRef(MarkerIndex).Index = z ' flag so this segement not processed again if not needed
            For p = 1 To timeInterval    ' grab timeInterval measurements at timeInterval intervals
                If p = timeInterval Then T = 1! Else T = m_BezierPrecision * p
                tSquared = T * T
                tCubed = tSquared * T
                ' calc point on curve in relation to t : time
                pSeg(8).X = pSeg(0).X + (-pSeg(5).X + T * (pSeg(5).X - pSeg(0).X * T)) * T
                pSeg(8).X = pSeg(8).X + (pSeg(6).X + T * (-6 * pSeg(1).X + pSeg(6).X * T)) * T
                pSeg(8).X = pSeg(8).X + (pSeg(7).X - pSeg(7).X * T) * tSquared
                pSeg(8).X = pSeg(8).X + pSeg(3).X * tCubed
            
                pSeg(8).Y = pSeg(0).Y + (-pSeg(5).Y + T * (pSeg(5).Y - pSeg(0).Y * T)) * T
                pSeg(8).Y = pSeg(8).Y + (pSeg(6).Y + T * (-6 * pSeg(1).Y + pSeg(6).Y * T)) * T
                pSeg(8).Y = pSeg(8).Y + (pSeg(7).Y - pSeg(7).Y * T) * tSquared
                pSeg(8).Y = pSeg(8).Y + pSeg(3).Y * tCubed
                
                ' calc true distance from previous point
                T = pSeg(4).X - pSeg(8).X: b = pSeg(4).Y - pSeg(8).Y
                ' store accumulative distance in relation to time t
                m_LinearXRef(MarkerIndex).Distance(s) = m_LinearXRef(MarkerIndex).Distance(s - 1) + Sqr(T * T + b * b)
                ' set new previous point
                pSeg(4) = pSeg(8)
                s = s + 1
            Next
        End If
    
        ' subtract requested distance from accumulated distance of previous segment
        Distance = Int(Distance - m_SubPaths.Markers(MarkerIndex).Segments(m_LinearXRef(MarkerIndex).Index - 1).X)
        For z = 1 To timeInterval   ' check the samples to find which segment the Distance falls on
            If Distance <= m_LinearXRef(MarkerIndex).Distance(z) Then
                ' calc ratio of Distance change in relation to previous point
                With m_LinearXRef(MarkerIndex)
                    b = (.Distance(z) - .Distance(z - 1)) ' could be zero
                    If Not b = 0! Then b = (Distance - .Distance(z - 1)) / b
                End With
                ' now calculate time change using the above ratio b
                T = ((z - 1) * m_BezierPrecision) + (m_BezierPrecision * b)
                ' get the point on the bezier for time t
                tSquared = T * T
                tCubed = tSquared * T
                ' bezier time-to-point formula
                pSeg(8).X = pSeg(0).X + (-pSeg(5).X + T * (pSeg(5).X - pSeg(0).X * T)) * T
                pSeg(8).X = pSeg(8).X + (pSeg(6).X + T * (-6 * pSeg(1).X + pSeg(6).X * T)) * T
                pSeg(8).X = pSeg(8).X + (pSeg(7).X - pSeg(7).X * T) * tSquared
                pSeg(8).X = pSeg(8).X + pSeg(3).X * tCubed
            
                pSeg(8).Y = pSeg(0).Y + (-pSeg(5).Y + T * (pSeg(5).Y - pSeg(0).Y * T)) * T
                pSeg(8).Y = pSeg(8).Y + (pSeg(6).Y + T * (-6 * pSeg(1).Y + pSeg(6).Y * T)) * T
                pSeg(8).Y = pSeg(8).Y + (pSeg(7).Y - pSeg(7).Y * T) * tSquared
                pSeg(8).Y = pSeg(8).Y + pSeg(3).Y * tCubed
                Exit For
            End If
        Next
        bReturn = True
    End If
    
    ' return the point
    If UseTransformations Then
        If (m_State And ps_TransformMatrix) = 0& Then zSpt_TransformPath False, False
        With m_Source
            X = ((pSeg(8).X * .Transformation(0)) + (pSeg(8).Y * .Transformation(2)) + .Transformation(4))
            Y = ((pSeg(8).X * .Transformation(1)) + (pSeg(8).Y * .Transformation(3)) + .Transformation(5))
        End With
    Else
        X = pSeg(8).X: Y = pSeg(8).Y
    End If
    
    If (m_State And ps_PointsCachable) = 0& Then
        m_State = (m_State And Not ps_LinearXRefCreated)
        Erase m_PathPoints
        Erase m_PathPtType
        Erase m_LinearXRef()
    End If
    
    NavigatePath = bReturn

End Function

Public Sub Render(ByVal hDC As Long, _
                Optional ByVal OffsetX As Single = 0!, Optional ByVal OffsetY As Single = 0!, _
                Optional ByVal gdipGraphicsObject As Long = 0&)
                
    ' ---=====| Renders the path to a DC |====---
    ' Parameters:
    '   hDC :: the destination DC
    '   OffsetX/Y :: Optional rendering position offsets. Not applied permanently
    '   gdipGraphicsObject :: if passed, that graphics object is used for rendering
    '       -- this may allow you more control over how the path will be rendered
    '       -- Default in this routine is to set the Graphics object to the highest quality
    
    If hDC = 0& Then Exit Sub                   ' no hDC?
    If m_Source.handle = 0& Then Exit Sub       ' no path to render
    If m_BrushCount = 0& Then Exit Sub          ' no brushes/pens to render with
                
    Dim hGraphics As Long
    Dim prevGState As Long, offsetState As Long
    Dim wmx As Long, brMx As Long, br As Long
    
    ' return a matrix used for rotation, scaling, translating/offsetting
    wmx = zSpt_TransformPath(False, True, OffsetX, OffsetY)
    If wmx = 0& Then Exit Sub
    
    If Not gdipGraphicsObject = 0& Then ' user provided GDI+ graphics object
        GdipSaveGraphics gdipGraphicsObject, prevGState ' test it, while also saving its current state so we can restore it
        If prevGState = 0& Then
            gdipGraphicsObject = 0&        ' failure :: also doubles as a flag for clean up at end of routine
        Else
            hGraphics = gdipGraphicsObject ' use the passed object
        End If
    End If
    If gdipGraphicsObject = 0& Then         ' create graphics object
        GdipCreateFromHDC hDC, hGraphics
        If hGraphics = 0& Then Exit Sub     ' test it & then set output quality
        GdipSetInterpolationMode hGraphics, InterpolationModeHighQualityBicubic
        GdipSetPixelOffsetMode hGraphics, PixelOffsetModeHighQuality
        GdipSetSmoothingMode hGraphics, SmoothingModeHighQuality
    End If
    
    
    GdipSetWorldTransform hGraphics, wmx   ' apply world transformation

    ' reassign brush boundaries as needed & then render
    For br = 1 To m_BrushCount
        With m_Brushes(br)
            ' update the path bounds if needed
            If .UpdateGradientBounds Then .AssignGradientBrush m_Source.handle
            If Not (.FillHandle = 0& And .OutlineHandle = 0&) Then
                
                If Not (.OffsetX = 0! And .OffsetY = 0!) Then
                    ' the brush/pen is to be offset, do that now
                    GdipSaveGraphics hGraphics, offsetState
                    If wmx = 0& Then
                        GdipCreateMatrix brMx
                    Else
                        GdipCloneMatrix wmx, brMx
                    End If  ' offset and apply to the hGraphics object
                    GdipTranslateMatrix brMx, .OffsetX, .OffsetY, MatrixOrderPrepend
                    GdipSetWorldTransform hGraphics, brMx
                End If
                
                ' now fill and/or outline the path
                If Not .FillHandle = 0& Then GdipFillPath hGraphics, .FillHandle, m_Source.handle
                If Not .OutlineHandle = 0& Then GdipDrawPath hGraphics, .OutlineHandle, m_Source.handle
                
                If Not brMx = 0& Then ' reset the if shifted above
                    GdipRestoreGraphics hGraphics, offsetState
                    GdipDeleteMatrix brMx
                    brMx = 0&
                End If
            
            End If
        End With
    Next
    
    ' clean up
    If Not wmx = 0& Then GdipDeleteMatrix wmx   ' destroy any matrix created
    If gdipGraphicsObject = 0& Then             ' destroy locally created object
        GdipDeleteGraphics hGraphics
    Else                                        ' restore passed graphics handle
        GdipRestoreGraphics hGraphics, prevGState
    End If

End Sub

Public Sub RenderOutline_XOR(ByVal hDC As Long, ByVal PenColor As Long, _
            Optional ByVal PenThickness As Long = 1, Optional ByVal Style As ePathXORstyle = psXOR_Solid, _
            Optional ByVal OffsetX As Single = 0!, Optional ByVal OffsetY As Single = 0!)
            
    ' ---=====| Renders the path to a DC using XOR pen/style|====---
    ' See the Render method for parameter comments
            
    ' XOR will draw something and then if the called again with the same settings
    ' will erase what was previously drawn. Unfortunately, we do this via GDI
    ' because GDI+ does not support XOR drawing -- big drawback
    
    ' Tip: If this is being called to create a sizing image while being sized, then
    ' call CachePathPoints=True as sizing starts and CachePathPoints=False when done.
    
    If m_Source.handle = 0& Then Exit Sub
    If hDC = 0& Then Exit Sub
    
    Dim hPen As Long, hOldPen As Long, lROP2 As Long
    Dim Index As Long, Pts(0 To 1, 0 To 3) As Long, p As Long
    
    ' Since we can't use GDI+ to XOR render, we can't use GDI+ to transform the points/world.
    ' Really, we could -- we'd have to clone the path, transform it, get the points then
    ' render those points. But, we will just process the transformation matrix manually
    If zSpt_TransformPath(False, False, OffsetX, OffsetY) = 0 Then Exit Sub
    If zSpt_GetPathData(m_PathPoints(), m_PathPtType(), False) = False Then Exit Sub
    
    If PenThickness < 1 Then PenThickness = 1
    hPen = CreatePen(Style, PenThickness, PenColor)
    If hPen Then
        
        hOldPen = SelectObject(hDC, hPen)
        lROP2 = SetROP2(hDC, vbNotXorPen)
        
        For Index = 0 To UBound(m_PathPtType)
        
            Select Case (m_PathPtType(Index) And PathPointTypePathTypeMask)
            
            Case PathPointTypeStart
                ' beginning of a new segment
                With m_Source
                    Pts(0, 3) = ((m_PathPoints(Index).X * .Transformation(0)) + (m_PathPoints(Index).Y * .Transformation(2)) + .Transformation(4))
                    Pts(1, 3) = ((m_PathPoints(Index).X * .Transformation(1)) + (m_PathPoints(Index).Y * .Transformation(3)) + .Transformation(5))
                End With
                MoveToEx hDC, Pts(0, 3), Pts(1, 3), ByVal 0&
                
            Case PathPointTypeLine
                ' simple line, always 2 points. The 1st pt is pt of last segment
                With m_Source
                    Pts(0, 1) = ((m_PathPoints(Index).X * .Transformation(0)) + (m_PathPoints(Index).Y * .Transformation(2)) + .Transformation(4))
                    Pts(1, 1) = ((m_PathPoints(Index).X * .Transformation(1)) + (m_PathPoints(Index).Y * .Transformation(3)) + .Transformation(5))
                End With
                LineTo hDC, Pts(0, 1), Pts(1, 1)
                
            Case PathPointTypeBezier
                ' always 4 points per bezier. The 1st pt is pt of last segment
                With m_Source
                    For p = 0 To 2
                        Pts(0, p) = ((m_PathPoints(Index + p).X * .Transformation(0)) + (m_PathPoints(Index + p).Y * .Transformation(2)) + .Transformation(4))
                        Pts(1, p) = ((m_PathPoints(Index + p).X * .Transformation(1)) + (m_PathPoints(Index + p).Y * .Transformation(3)) + .Transformation(5))
                    Next
                End With
                PolyBezierTo hDC, Pts(0, 0), 3
                Index = Index + 2
                
            End Select
            ' close open path if required
            If (m_PathPtType(Index) And PathPointTypeCloseSubpath) Then LineTo hDC, Pts(0, 3), Pts(1, 3)
        
        Next
        DeleteObject SelectObject(hDC, hOldPen)
        SetROP2 hDC, lROP2
    
    End If
    
    If (m_State And ps_PointsCachable) = 0& Then
        Erase m_PathPoints
        Erase m_PathPtType
        m_State = (m_State And Not ps_PointsCached)
    End If

End Sub

Public Function WarpPath_PTR(ByVal sourcePtsPTR As Long, ByVal destPtsPTR As Long, _
                                Optional ByVal WarpStyle As WarpModes = warpBilinear, _
                                Optional ByVal PtOrder As ePathOrder = TopLeft_zPattern) As Boolean

    ' ---=====| Warps/Translates points from source to destination |====---
    ' Results are placed in the destination array
    
    ' ApplyToPath: if True, the points are warped then applied to the path.
    ' PtOrder:
    '       TopLeft_Clockwise - sourcePtsPTR = Top/Left X, sourcePtsPTR+4 = Top/Left Y
    '                           sourcePtsPTR+8 = Top/Right X, sourcePtsPTR+12 = Top/Right Y
    '                           sourcePtsPTR+16 = Bottom/Right X, sourcePtsPTR+20 = Bottom/Right Y
    '                           sourcePtsPTR+24 = Bottom/Left X, sourcePtsPTR+28 = Bottom/Left Y
    '       TopLeft_zPattern  - sourcePtsPTR = Top/Left X, sourcePtsPTR+4 = Top/Left Y
    '                           sourcePtsPTR+8 = Top/Right X, sourcePtsPTR+12 = Top/Right Y
    '                           sourcePtsPTR+16 = Bottom/Left X, sourcePtsPTR+20 = Bottom/Left Y
    '                           sourcePtsPTR+24 = Bottom/Right X, sourcePtsPTR+28 = Bottom/Right Y
    '       Note that the last two points are swapped between the two patterns
    
    ' The SourcePtsPTR & destPtsPTR must be valid array pointers containing Single values, not Long.
    ' Each array must contain at least 32 bytes of data, equivalent to 4 point structures each.
    ' If the array are Single variable types, size as (0 to 1, 0 to 3) where 1st dim are X,Y and 2nd dim are the order
    ' The pointers can be passed as:  VarPtr(mySingles(0,0)), VarPtr(myPOINTF(0)) for example.
    
    ' If invalid array pointers are passed, a crash can occur
    
    If sourcePtsPTR = 0 Or destPtsPTR = 0 Then Exit Function
    
    Dim tstSrc(0 To 3) As POINTF, tstDst(0 To 3) As POINTF
    
    If PtOrder = TopLeft_zPattern Then
        CopyMemory tstSrc(0), ByVal sourcePtsPTR, 32&
        CopyMemory tstDst(0), ByVal destPtsPTR, 32&
    Else
        CopyMemory tstDst(0), ByVal destPtsPTR, 16&
        CopyMemory tstDst(3), ByVal destPtsPTR + 24, 8&
        CopyMemory tstDst(2), ByVal destPtsPTR + 16, 8&

        CopyMemory tstSrc(0), ByVal sourcePtsPTR, 16&
        CopyMemory tstSrc(3), ByVal sourcePtsPTR + 24, 8&
        CopyMemory tstSrc(2), ByVal sourcePtsPTR + 16, 8&
    End If
    
    If Not zSpt_WarpMatrix(tstSrc(), tstDst(), WarpStyle) = 0 Then
        CopyMemory ByVal destPtsPTR, tstDst(0), 32&
        WarpPath_PTR = True
    End If

End Function

Public Function ExportPath(exportArray() As Byte) As Boolean

    ' ---=====| Exports a path into a custom-used format. |====---
    ' If return value is False, exportArray() should be considered null
    
    ' The format returned in the array will be as follows
    ' 1st 4 Bytes: number of points in the path
    ' next nrPts*9 bytes: path points (8 bytes each) then path types (1 byte per point)
    ' if warped then the value of 1 and following is also exported; else value of zero
    ' next 32 bytes: warp points
    ' next 32 bytes: unwarped points
    ' next 1 byte: warp style: 0=bilinear,1=perspective
    
    If m_Source.handle Then
        Dim nrPts As Long, aIndex As Long
        GdipGetPointCount m_Source.handle, nrPts
        ' if warped then use: ReDim exportArray(0 To nrPts * 9 + 69)
        ReDim exportArray(0 To nrPts * 9 + 4)
        CopyMemory exportArray(0), nrPts, 4&
        GdipGetPathPoints m_Source.handle, VarPtr(exportArray(4)), nrPts
        aIndex = nrPts * 8 + 4
        GdipGetPathTypes m_Source.handle, VarPtr(exportArray(aIndex)), nrPts
        
        ' if warped then, 65 more bytes
        ' copy the warp points; 32 bytes
        ' copy the unwarp points; 32 bytes
        ' set final byte to 1 if perspetive warp
        ExportPath = True
    End If

End Function

Public Function ImportPath(PATHDATA() As Byte) As Boolean
    ' ---=====| Imports a path from array created from this class's ExportPath routine |====---
    
    On Error GoTo ExitRoutine
    Dim nrPts As Long, aIndex As Long, requiredSize As Long
    Dim pPath As Long
    aIndex = LBound(PATHDATA)
    If UBound(PATHDATA) > aIndex + 4 Then
        CopyMemory nrPts, PATHDATA(aIndex), 4&
        ' now calculate the required size of the array before doing any imports
        requiredSize = requiredSize + nrPts * 9 + 4 + aIndex
        If UBound(PATHDATA) >= requiredSize Then
            If PATHDATA(requiredSize) = 1 Then ' warped data exists
                requiredSize = requiredSize + 65
            End If
        End If
        If UBound(PATHDATA) >= requiredSize Then ' continue importing
            aIndex = aIndex + 4 ' location of path type data
            GdipCreatePath2 VarPtr(PATHDATA(aIndex)), VarPtr(PATHDATA(aIndex + nrPts * 8)), nrPts, ePathBrushFillMode.FillModeWinding, pPath
            If Not pPath = 0& Then
                If m_Source.handle Then GdipDeletePath m_Source.handle
                m_Source.handle = pPath
                zSpt_DoPostUpdateActions
                ImportPath = True
            End If
        End If
    End If
ExitRoutine:
End Function


' =========================================================================================================


Private Function zSpt_WarpMatrix(srcPts() As POINTF, dstPts() As POINTF, _
                                ByVal WarpStyle As WarpModes) As Boolean
'                                ByVal WarpStyle As WarpModes, Optional ByVal Rotation As Single = 0!, _
'                                Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!) As Boolean

    ' ---=====| Warps points from one realm to another realm |====---

    ' Visit the following link to better understand. This matrix and the associated
    ' zSpt_SovlePolyEquation routines are required to manually warp points vs using the
    ' flawed GDI+ function.  Again, look at the link
    ' http://www.codeguru.com/cpp/g-m/gdi/gdi/article.php/c3657/
    
    If m_Source.handle = 0& Then Exit Function
    
    Dim aMatrix() As Single, aWarpFactors() As Single
    Dim Row As Long, Index As Long
    Dim tPath As Long
    Dim X As Single, Y As Single, tVal As Single
    
    ' Get the path point data
    If zSpt_GetPathData(m_PathPoints(), m_PathPtType(), False) = False Then Exit Function
    
    If WarpStyle Then
   
        '// In perspective mode, the warping functions are:
        '//  x' = (a0 + a1 x + a2 y) / (c0 x + c1 y + 1)
        '//  y' = (b0 + b1 x + b2 y) / (c0 x + c1 y + 1)
        '//
        '// The following calculates the factors a#, b# and c#.
        '// We do this by creating a set of eight equations with a#, b# and c# as unknowns.
        '// The equations are derived by:
        '// 1. substituting the srcPoints for (x, y);
        '// 2. substituting the corresponding destPoints for (x', y');
        '// 3. solving the resulting set of equations, with the factors as unknowns.
        '//
        '// The equations are like these:
        '//  a0  x a1    y a2    0       0       0       -xx'c0  -yx'c1  = x'
        '//  0   0       0       b0      x b1    y b2    -xy'c0  -yy'c1  = y'
        '// The known factors of left hand side ar put in the 8x8 matrix mxLeft for
        '// all four point pairs, and the right hand side in the one column matrix mxRight.
        '// After solving, m_mxWarpFactors contains a0, a1, a2, b0, b1, b2, c0, c1.
        '
        ReDim aMatrix(0 To 7, 0 To 7)
        ReDim aWarpFactors(0 To 7, 0 To 0)
        
        For Index = 0 To 3
            aMatrix(Row, 0) = 1!
            aMatrix(Row, 1) = srcPts(Index).X
            aMatrix(Row, 2) = srcPts(Index).Y
            aMatrix(Row, 6) = -srcPts(Index).X * dstPts(Index).X
            aMatrix(Row, 7) = -srcPts(Index).Y * dstPts(Index).X
            aWarpFactors(Row, 0) = dstPts(Index).X
            Row = Row + 1
            aMatrix(Row, 3) = 1!
            aMatrix(Row, 4) = srcPts(Index).X
            aMatrix(Row, 5) = srcPts(Index).Y
            aMatrix(Row, 6) = -srcPts(Index).X * dstPts(Index).Y
            aMatrix(Row, 7) = -srcPts(Index).Y * dstPts(Index).Y
            aWarpFactors(Row, 0) = dstPts(Index).Y
            Row = Row + 1
        Next
        
    Else ' BiLinear & Simple Skew
    
        '// In bilinear mode, the warping functions are:
        '//  x' = a0 + a1 x y + a2 x + a3 y
        '//  y' = b0 + b1 x y + b2 x + b3 y
        '//
        '// Here, we have two sets of four equations. In the first set, the a# factors
        '// are the unknowns, in the second set the b# factors.
        '// The equations are of the form:
        '//  a0      + xy a1     + x a2      + y a3  = x'
        '// The left hand side is identical for both sets. The right hand side differs.
        '// Therefore, we can solve them in one operation.
        '// The left hand side factors are put in the 4x4 matrix mxLeft, the right side
        '// factors are put in the 4x2 matrix mxRight.
        '// After solving, the first column of m_mxWarpFactors contains a0, a1, a2, a3; the
        '// second columne contains b0, b1, b2, b3.

        ReDim aMatrix(0 To 3, 0 To 3)
        ReDim aWarpFactors(0 To 3, 0 To 1)

        For Index = 0 To 3
            aMatrix(Index, 0) = 1
            aMatrix(Index, 1) = srcPts(Index).X * srcPts(Index).Y
            aMatrix(Index, 2) = srcPts(Index).X
            aMatrix(Index, 3) = srcPts(Index).Y
            
            aWarpFactors(Index, 0) = dstPts(Index).X
            aWarpFactors(Index, 1) = dstPts(Index).Y
        Next
        
    End If
    
    If zSpt_SovlePolyEquation(aMatrix, aWarpFactors) Then
        
        If WarpStyle Then                   ' formula for perspective warping
            For Index = 0 To UBound(m_PathPoints)
                X = m_PathPoints(Index).X
                Y = m_PathPoints(Index).Y
                tVal = aWarpFactors(6, 0) * X + aWarpFactors(7, 0) * Y + 1!
                m_PathPoints(Index).X = ((aWarpFactors(0, 0) + aWarpFactors(1, 0) * X + aWarpFactors(2, 0) * Y) / tVal)
                m_PathPoints(Index).Y = ((aWarpFactors(3, 0) + aWarpFactors(4, 0) * X + aWarpFactors(5, 0) * Y) / tVal)
            Next
        Else
            For Index = 0 To UBound(m_PathPoints)  ' formula for bilinear & simple skew warping
                X = m_PathPoints(Index).X
                Y = m_PathPoints(Index).Y
                tVal = X * Y
                m_PathPoints(Index).X = aWarpFactors(0, 0) + aWarpFactors(1, 0) * tVal + aWarpFactors(2, 0) * X + aWarpFactors(3, 0) * Y
                m_PathPoints(Index).Y = aWarpFactors(0, 1) + aWarpFactors(1, 1) * tVal + aWarpFactors(2, 1) * X + aWarpFactors(3, 1) * Y
            Next
    
        End If
            
        ' now replace the transformed path
        If GdipCreatePath2(VarPtr(m_PathPoints(0)), VarPtr(m_PathPtType(0)), UBound(m_PathPoints) + 1, ePathBrushFillMode.FillModeWinding, tPath) = 0& Then
            GdipDeletePath m_Source.handle
            m_Source.handle = tPath
            zSpt_DoPostUpdateActions
            zSpt_WarpMatrix = True
        End If
        
        If (m_State And ps_PointsCachable) = 0& Then
            Erase m_PathPoints
            Erase m_PathPtType
            m_State = (m_State And Not ps_PointsCached)
        End If
        
    End If
    
End Function

Private Function zSpt_SovlePolyEquation(sourceArray() As Single, RHSarray() As Single) As Boolean

    ' ---=====| Helper function for WarpMatrix |====---
    ' Using a matrix to solve simultaneous equations
    ' http://www.java2s.com/Tutorial/VB/0040__Data-Type/Solveequationsusingmatrices.htm
    ' Not commented: see above link. I'd need a couple of pages to try to explain it
    
    Dim I As Long, J As Long, K As Long
    Dim Rank As Long, RankUnk As Long
    Dim a1 As Single
    Dim temp() As Single
    
    Rank = UBound(sourceArray, 2)
    
On Error GoTo eh
    
    RankUnk = UBound(RHSarray, 2) + 1
    ReDim temp(0 To Rank, 0 To Rank + RankUnk)
    
    For I = 0 To Rank
        For J = 0 To Rank
            temp(I, J) = sourceArray(I, J)
        Next
        For K = 0 To RankUnk - 1
            temp(I, J + K) = RHSarray(I, K)
        Next
    Next
    
    For K = 0 To Rank
        If zSpt_PivotMatrix(temp, K, Rank) = False Then
            ' singular matrix, no solution
            Exit Function
        End If
        a1 = temp(K, K)
        If a1 = 0 Then Stop ' debugging
        For J = K To Rank + RankUnk
            temp(K, J) = temp(K, J) / a1
        Next
        For I = K + 1 To Rank
            a1 = temp(I, K)
            For J = K To Rank + RankUnk
                temp(I, J) = temp(I, J) - a1 * temp(K, J)
            Next
        Next
    Next
    
    For K = 0 To RankUnk - 1
        For I = Rank To 0 Step -1
            RHSarray(I, K) = temp(I, Rank + 1 + K)
            For J = I + 1 To Rank
                RHSarray(I, K) = RHSarray(I, K) - temp(I, J) * RHSarray(J, K)
            Next
        Next
    Next
    zSpt_SovlePolyEquation = True
eh:
If Err Then
    Stop                ' debugging purposes only
    Err.Clear
    Resume
End If
End Function

Private Function zSpt_PivotMatrix(srcMatrix() As Single, Row As Long, Rank As Long) As Boolean

    ' ---=====| Helper function for zSpt_SovlePolyEquation |====---
    
    Dim K As Long, I As Long
    Dim aMax As Single, tmpVal As Single
    
    K = Row
    For I = Row To Rank ' row
        tmpVal = Abs(srcMatrix(I, Row))
        If tmpVal > aMax Then
            If tmpVal <> 0! Then
                aMax = tmpVal
                K = I
            End If
         End If
    Next
    If Not srcMatrix(K, Row) = 0! Then
        If K <> Row Then
            For I = 0 To UBound(srcMatrix, 2)
                tmpVal = srcMatrix(K, I)
                srcMatrix(K, I) = srcMatrix(Row, I)
                srcMatrix(Row, I) = tmpVal
            Next
        End If
        zSpt_PivotMatrix = True
    End If
End Function

Private Function zSpt_MeasurePath(MarkerIndex As Long) As Double

    ' ---=====| Guesstimate length of path - reasonable accuracy only |====---
    ' calc point on bezier: http://lists.apple.com/archives/cocoa-dev/2002/Feb/msg01806.html
    
    ' Costly routine due to use of Square Roots, Squares & Cubes. Therefore,
    ' this is only run when the path has changed.
    
    If m_Source.handle = 0& Then Exit Function
    If (m_State And ps_Measured) = ps_Measured Then
        If MarkerIndex <= m_SubPaths.Count Then
            If MarkerIndex = 0& Or (MarkerIndex = 1& And m_SubPaths.Count = 1&) Then
                zSpt_MeasurePath = m_SubPaths.Length
            ElseIf MarkerIndex > 0& Then
                zSpt_MeasurePath = m_SubPaths.Markers(MarkerIndex).Segments(UBound(m_SubPaths.Markers(MarkerIndex).Segments)).X
            End If
        End If
        Exit Function
    End If
    
    Dim a As Single, b As Single
    Dim tSquared As Single, tCubed As Single, T As Single
    Dim pSeg(0 To 1, 0 To 8) As Single
    Dim segCount As Long, markCount As Long
    Dim ttLen As Single
    Dim X As Long, I As Long, p As Long, lastOpenPt As Long
    Dim timeSize As Long
    
    If m_SubPaths.Count = 0& Then
        If zSpt_GetSubpathInfo(False) = False Then Exit Function
    Else
        If zSpt_GetPathData(m_PathPoints(), m_PathPtType(), False) = False Then Exit Function
    End If
    
    timeSize = 1 / m_BezierPrecision
    markCount = 1&
    For I = 1 To UBound(m_PathPtType)
        Select Case (m_PathPtType(I) And PathPointTypePathTypeMask)
        
        Case PathPointTypeBezier
            segCount = segCount + 1
            ReDim Preserve m_SubPaths.Markers(markCount).Segments(0 To segCount)
            m_SubPaths.Markers(markCount).Segments(segCount).Y = I + 2
            CopyMemory pSeg(0, 0), m_PathPoints(I - 1), 32&     ' copy 4 points to array
            CopyMemory pSeg(0, 4), pSeg(0, 0), 8&               ' cache moving starting point
            
            For X = 5 To 7                          ' cache precalced multipliers
                pSeg(0, X) = pSeg(0, X - 5) * 3     ' pts(0,1,2) * 3
                pSeg(1, X) = pSeg(1, X - 5) * 3
            Next
            
            ' subsegment the bezier into timeSize-related points
            For p = 1 To timeSize
                If p = timeSize Then T = 1! Else T = p * m_BezierPrecision ' ensure last point is measured
                tSquared = T * T
                tCubed = tSquared * T
                
                ' calc point on spline in relation to t
                For X = 0 To 1
                    pSeg(X, 8) = pSeg(X, 0) + (-pSeg(X, 5) + T * (pSeg(X, 5) - pSeg(X, 0) * T)) * T
                    pSeg(X, 8) = pSeg(X, 8) + (pSeg(X, 6) + T * (-6 * pSeg(X, 1) + pSeg(X, 6) * T)) * T
                    pSeg(X, 8) = pSeg(X, 8) + (pSeg(X, 7) - pSeg(X, 7) * T) * tSquared
                    pSeg(X, 8) = pSeg(X, 8) + pSeg(X, 3) * tCubed
                Next
                ' calc true distance from previous point & cache
                a = pSeg(0, 4) - pSeg(0, 8): b = pSeg(1, 4) - pSeg(1, 8)
                ttLen = ttLen + Sqr(a * a + b * b)
                ' set new previous point
                CopyMemory pSeg(0, 4), pSeg(0, 8), 8&
            Next
            m_SubPaths.Markers(markCount).Segments(segCount).X = ttLen
            'm_Segments(segCount).X = ttLen   ' cache spline length
            I = I + 2                   ' move to next path segment
            
        Case PathPointTypeLine
            segCount = segCount + 1     ' lines are easier
            ReDim Preserve m_SubPaths.Markers(markCount).Segments(0 To segCount)
            m_SubPaths.Markers(markCount).Segments(segCount).Y = I
            a = m_PathPoints(I - 1).X - m_PathPoints(I).X
            b = m_PathPoints(I - 1).Y - m_PathPoints(I).Y
            ttLen = ttLen + Sqr(a * a + b * b)
            m_SubPaths.Markers(markCount).Segments(segCount).X = ttLen
            'm_Segments(segCount).X = ttLen   ' total length
            
        Case PathPointTypeStart
            lastOpenPt = I   'keep track if an open path is later closeed; it will be the end point
            ' ^^ see below after the End Select statement
        Case Else
            Stop ' should never get here
        End Select
        
        If (m_PathPtType(I) And PathPointTypeCloseSubpath) = PathPointTypeCloseSubpath Then
            ' special case, path does not end on last starting point, but should be drawn to it....
            segCount = segCount + 1     ' lines are easier
            ReDim Preserve m_SubPaths.Markers(markCount).Segments(0 To segCount)
            ' track position of where seg ends, it starts on last point
            m_SubPaths.Markers(markCount).Segments(segCount).Y = lastOpenPt
            a = m_PathPoints(lastOpenPt).X - m_PathPoints(I).X
            b = m_PathPoints(lastOpenPt).Y - m_PathPoints(I).Y
            ttLen = ttLen + Sqr(a * a + b * b)
            m_SubPaths.Markers(markCount).Segments(segCount).X = ttLen   ' total length
        End If
        If (m_PathPtType(I) And PathPointTypePathMarker) = PathPointTypePathMarker Then
            If I < UBound(m_PathPtType) Then ' start new marker segment
                m_SubPaths.Length = m_SubPaths.Length + ttLen
                ttLen = 0!
                segCount = 0&
                markCount = markCount + 1&
            End If
        End If
        
    Next
    m_SubPaths.Length = m_SubPaths.Length + ttLen
    m_State = (m_State Or ps_Measured)
    
    If MarkerIndex = 0 Then
        zSpt_MeasurePath = m_SubPaths.Length
    Else
        zSpt_MeasurePath = m_SubPaths.Markers(MarkerIndex).Segments(UBound(m_SubPaths.Markers(MarkerIndex).Segments)).X
    End If
    
End Function

Private Sub zSpt_DoPostUpdateActions()

    ' ---=====| Modify state/flags when path is modified |====---

    m_State = ((m_State And ps_PointsCachable) And ps_TransformOffsets)
    If (m_State And ps_PointsCachable) Then Me.CachePathPoints = True
    Erase m_SubPaths.Markers
    m_SubPaths.Count = 0&
    Erase m_LinearXRef
    For m_BrushCount = 0 To m_BrushCount - 1
        m_Brushes(m_BrushCount + 1).UpdateGradientBounds = True
    Next
    
    Dim pRect As RECTF
    GdipGetPathWorldBounds m_Source.handle, pRect, 0&, 0&
    With m_Source
        .Origin(0).X = pRect.Left: .Origin(0).Y = pRect.Top
        .Origin(1).X = pRect.Left + pRect.Width - 1: .Origin(1).Y = pRect.Top
        .Origin(2).X = pRect.Left: .Origin(2).Y = pRect.Top + pRect.Height - 1
        .Origin(3).X = .Origin(1).X: .Origin(3).Y = .Origin(2).Y
        Erase .ptOffsets
    End With
    
End Sub

Private Function zSpt_GetPathData(Pts() As POINTF, pTypes() As Byte, ReQueryIfCached As Boolean) As Boolean

    ' ---=====| Returns this path's data to passed arrays |====---

    If m_Source.handle = 0& Then Exit Function
    If (m_State And ps_PointsCached) = ps_PointsCached Then
        If ReQueryIfCached = False Then ' else a forced refresh
            zSpt_GetPathData = True
            Exit Function
        End If
    End If
    
    Dim pd As PATHDATA
    GdipGetPointCount m_Source.handle, pd.Count
    If pd.Count > 1 Then        ' valid paths have at least 2 points (a line)
        ReDim Pts(0 To pd.Count - 1)
        ReDim pTypes(0 To pd.Count - 1)
        pd.PointsPTR = VarPtr(Pts(0))
        pd.TypesPtr = VarPtr(pTypes(0))
        If GdipGetPathData(m_Source.handle, pd) = 0& Then
            m_State = m_State Or ps_PointsCached
            zSpt_GetPathData = True
        End If
    End If
    
End Function

Private Function zSpt_GetSubpathInfo(ByVal bReleasePathData As Boolean) As Boolean

    ' ---=====| Determines subpath data when Markers are used |====---
    ' Markers are optional parameters to the various Append_xxx functions
    ' and offer a way of "naming" portions of a path (subdividing the path)
    
    If m_SubPaths.Count Then            ' already done; return
        zSpt_GetSubpathInfo = True
        Exit Function
    End If
    
    Dim I As Long, lastPt As Long
    
    ' get path data
    If zSpt_GetPathData(m_PathPoints(), m_PathPtType(), False) Then
        
        m_SubPaths.Count = 1 ' start with 1st marker (even if path is unmarked)
        ReDim m_SubPaths.Markers(1 To m_SubPaths.Count)
        
        ' find coded markers within the path & track marker indexes
        For I = 1 To UBound(m_PathPtType)
        
            Select Case (m_PathPtType(I) And PathPointTypePathTypeMask)
            Case PathPointTypeStart
                lastPt = I
            Case PathPointTypeBezier
                I = I + 2
            Case Else ' PathPointTypeLine
            End Select
            
            If (m_PathPtType(I) And PathPointTypePathMarker) = PathPointTypePathMarker Then
                ' marker located: update indexes
                With m_SubPaths.Markers(m_SubPaths.Count)
                    .Count = I - .StartIndex + 1    ' nr of points in this subpath
                    ' determine last point of subpath
                    If (m_PathPtType(I) And PathPointTypeCloseSubpath) = 0& Then
                        .EndIndex = I
                    Else
                        .EndIndex = lastPt
                    End If
                End With
                ' resize marker array & increment to next marker if not on last point
                If I < UBound(m_PathPtType) Then
                    m_SubPaths.Count = m_SubPaths.Count + 1
                    ReDim Preserve m_SubPaths.Markers(1 To m_SubPaths.Count)
                    m_SubPaths.Markers(m_SubPaths.Count).StartIndex = I + 1
                End If
            End If
        Next
        
        ' handle last point if not done
        If (m_PathPtType(I - 1) And PathPointTypePathMarker) = 0& Then
            With m_SubPaths.Markers(m_SubPaths.Count)
                .Count = I - .StartIndex ' nr of points in this subpath
                ' determine last point of subpath
                If (m_PathPtType(I - 1) And PathPointTypeCloseSubpath) = 0& Then
                    .EndIndex = I - 1
                Else
                    .EndIndex = lastPt
                End If
            End With
        End If
        
        m_SubPaths.Length = 0#      ' reset length
        If bReleasePathData Then    ' release path data
            If (m_State And ps_PointsCached) = 0& Then
                Erase m_PathPoints
                Erase m_PathPtType
                m_State = (m_State And Not ps_PointsCached)
            End If
        End If
        zSpt_GetSubpathInfo = True  ' return
    End If
    
End Function

Private Function zSpt_GetBoundingRectToPOINTF(srcPts() As POINTF)
            
    ' ---=====| Simply converts 4 POINTF values to a RECTF structure |====---
            
    Dim pRect As RECTF
    ReDim srcPts(0 To 3)
    If m_Source.handle Then
        GdipGetPathWorldBounds m_Source.handle, pRect, 0&, 0&
        srcPts(0).X = pRect.Left: srcPts(0).Y = pRect.Top
        srcPts(1).X = pRect.Left + pRect.Width - 1: srcPts(1).Y = pRect.Top
        srcPts(2).X = pRect.Left: srcPts(2).Y = pRect.Top + pRect.Height - 1
        srcPts(3).X = srcPts(1).X: srcPts(3).Y = srcPts(2).Y
    End If
    
End Function

Private Function zSpt_CreateSubPathPath(MarkerIndex As Long, getClipBox As Boolean, _
                                        applyTransform As Boolean) As Long

    ' ---=====| Creates a GDI+ path object from a marked subpath |====---
    ' the returned path handle must be destroyed

    Dim subPts() As POINTF, subPtTypes() As Byte, nrPts As Long, Flags As Long
    Dim pClone As Long, mx As Long
    
    If MarkerIndex > m_SubPaths.Count Or m_SubPaths.Count = 0& Then Exit Function
    
    ' we will be partly building the path ourselves
    zSpt_GetPathData m_PathPoints, m_PathPtType, False
    
    With m_SubPaths.Markers(MarkerIndex)
        ' if the last point in the subpath closes a path, then we'll have
        ' to add the final point in the subpath ourselves
        If (m_PathPtType(.StartIndex + .Count - 1) And PathPointTypeCloseSubpath) = 0& Then
            nrPts = .Count
        Else
            nrPts = .Count + 1
            Flags = 2           ' need to add ending point
        End If
        ' if the subpath starts from other than a natural starting point,
        ' then we will have to add the first point to the subpath ourselves
        If Not m_PathPtType(.StartIndex) = PathPointTypeStart Then
            nrPts = nrPts + 1
            Flags = Flags Or 1  ' need to add starting point
        End If
        
        ReDim subPts(0 To nrPts - 1)
        ReDim subPtTypes(0 To nrPts - 1)
        
        If (Flags And 1) = 1 Then ' add first point
            subPts(0) = m_PathPoints(.StartIndex - 1)
            subPtTypes(0) = PathPointTypeStart
            CopyMemory subPts(1), m_PathPoints(.StartIndex), .Count * 8&
            CopyMemory subPtTypes(1), m_PathPtType(.StartIndex), .Count
        
        Else ' starts from a natural starting point
            CopyMemory subPts(0), m_PathPoints(.StartIndex), .Count * 8&
            CopyMemory subPtTypes(0), m_PathPtType(.StartIndex), .Count
        
        End If
        
        If (Flags And 2) = 2 Then ' add the extra closing point & tweak point type
            CopyMemory subPts(nrPts - 1), m_PathPoints(.EndIndex), 8&
            subPtTypes(nrPts - 1) = PathPointTypeLine ' make it a natural starting point
            ' remove the flag on last point indicating it is to close the path
            subPtTypes(nrPts - 2) = (subPtTypes(nrPts - 2) And Not PathPointTypeCloseSubpath)
        End If
        
        If getClipBox Then ' see GetBoundingBox for remarks on what we are doing here
            For nrPts = 1 To UBound(subPts)
                Select Case (subPtTypes(nrPts) And PathPointTypePathTypeMask)
                Case PathPointTypeStart
                    subPtTypes(nrPts) = PathPointTypeLine
                Case PathPointTypeBezier
                    nrPts = nrPts + 2
                End Select
                If (subPtTypes(nrPts) And PathPointTypeCloseSubpath) = PathPointTypeCloseSubpath Then
                    subPtTypes(nrPts) = (subPtTypes(nrPts) And Not PathPointTypeCloseSubpath)
                End If
            Next
            subPtTypes(nrPts - 1) = (subPtTypes(nrPts - 1) Or PathPointTypeCloseSubpath)
        End If
        
    End With
    
    If (m_State And ps_PointsCachable) = 0& Then
        Erase m_PathPoints
        Erase m_PathPtType
        m_State = (m_State And Not ps_PointsCached)
    End If
    
    GdipCreatePath2 VarPtr(subPts(0)), _
                    VarPtr(subPtTypes(0)), _
                    nrPts, ePathBrushFillMode.FillModeWinding, pClone
                    
    If applyTransform Then
        mx = zSpt_TransformPath(False, True)
        GdipTransformPath pClone, mx
        GdipDeleteMatrix mx
    End If
    Dim fRect As RECTF
    zSpt_CreateSubPathPath = pClone

End Function

Private Function zSpt_TransformPath(ReturnPath As Boolean, ReturnMatrix As Boolean, _
                                    Optional ByVal OffsetX As Single = 0!, _
                                    Optional ByVal OffsetY As Single = 0!) As Long
                    
    ' ---=====| Transforms path coords to other scales, angles, location |====---
    
    ' If ReturnPath = True and ReturnMatrix=False then Path handle is returned which must be destroyed
    ' If ReturnPath = False and ReturnMatrix=True then Matrix handle is returned which must be destroyed
    ' If ReturnPath=ReturnMatrix then
    '       If not warped, return is handle to a matrix & ReturnPath is set to false & ReturnMatrix is True
    '       If warped, return is handle to path & ReturnMatrix is set to False & ReturnPath is true
    
                    
    If m_Source.Scaler.X = 0! Or m_Source.Scaler.Y = 0! Then Exit Function
                    
    Dim warpPts() As POINTF, mx As Long, pClone As Long
    Dim Position As POINTF, centerPt As POINTF
    Dim angleRadian As Single
    
    If (m_State And ps_TransformMatrix) = 0& Then
        
        If Not m_HitTestPath = 0& Then
            GdipDeletePath m_HitTestPath
            m_HitTestPath = 0&
        End If
        
        With m_Source
        
            If .Rotation = 0! Then ' just scaling
            
                If (m_State And ps_TransformOffsets) = 0& Then
                    If (m_State And ps_ClipBox) = 0& Then ' get clip box for balanced rotation
                        If zSpt_GetBoundingBox(True, 0!, 0!, 0!, 0!) = False Then Exit Function
                    End If
                    Position.X = .DisplayLocale.X - m_clipBox.Left + m_clipBox.Left * (1 - .Scaler.X) + OffsetX
                    Position.Y = .DisplayLocale.Y - m_clipBox.Top + m_clipBox.Top * (1 - .Scaler.Y) + OffsetY
                Else
                    Position.X = .TransOverride(4)
                    Position.Y = .TransOverride(5)
                End If
                
                .Transformation(0) = .Scaler.X
                .Transformation(1) = .Rotation
                .Transformation(2) = .Rotation
                .Transformation(3) = .Scaler.Y
                .Transformation(4) = Position.X
                .Transformation(5) = Position.Y
        
            Else ' rotation and maybe scaling
        
                If (m_State And ps_TransformOffsets) = 0& Then
                    If (m_State And ps_ClipBox) = 0& Then ' get clip box for balanced rotation
                        If zSpt_GetBoundingBox(True, 0!, 0!, 0!, 0!) = False Then Exit Function
                    End If
                    centerPt.X = (m_clipBox.Width / 2 + m_clipBox.Left)
                    centerPt.Y = (m_clipBox.Height / 2 + m_clipBox.Top)
                    Position.X = .DisplayLocale.X - m_clipBox.Left - (m_clipBox.Width * (1 - .Scaler.X)) / 2
                    Position.Y = .DisplayLocale.Y - m_clipBox.Top - (m_clipBox.Height * (1 - .Scaler.Y)) / 2
                Else
                    centerPt.X = .TransOverride(0)
                    centerPt.Y = .TransOverride(1)
                    Position.X = .TransOverride(2)
                    Position.Y = .TransOverride(3)
                End If
            
                ' convert passed angle to radians
                angleRadian = .Rotation * (Atn(1) * 4) / 180
                ' build the matrix
                .Transformation(0) = Cos(angleRadian) * .Scaler.X   ' get X axis rotation angles
                .Transformation(1) = Sin(angleRadian) * .Scaler.X   ' while handling scaling/mirroring as needed
                .Transformation(2) = -Sin(angleRadian) * .Scaler.Y  ' same for the Y axis axis but force clockwise rotation
                .Transformation(3) = Cos(angleRadian) * .Scaler.Y
                .Transformation(4) = ((-centerPt.X) * .Transformation(0)) + ((-centerPt.Y) * .Transformation(2)) + centerPt.X + Position.X
                .Transformation(5) = ((-centerPt.Y) * .Transformation(3)) + ((-centerPt.X) * .Transformation(1)) + centerPt.Y + Position.Y
                
            End If
            
        End With
        
        m_State = (m_State Or ps_TransformMatrix)
        
    End If
    
    If m_Source.WarpStyle Then ' path is also warped
'        ReturnMatrix = False
'        ReturnPath = True
'        ' warp the path points
'        ReDim warpPts(0 To 3)
'        For mx = 0 To 3
'            warpPts(mx).x = m_Source.Origin(mx).x + m_Source.PtOffsets(mx).x
'            warpPts(mx).y = m_Source.Origin(mx).x + m_Source.PtOffsets(mx).y
'        Next
'        ' do the warp
'        zSpt_GetPathData m_PathPoints, m_PathPtType, False
'
'        ' get scale,rotation matrix values
'        zSpt_GetTransformMatrix Rotation, OffsetX, OffsetY, ScaleX, ScaleY, VarPtr(mxValues(0)), pRect.Left, pRect.Top
'
'        ' manually apply the matrix to the points
'        For mx = 0 To UBound(m_PathPoints)
'            With m_PathPoints(mx)
'                .x = (.x * mxValues(0)) + (.y * mxValues(2)) + mxValues(4) + OffsetX
'                .y = (.x * mxValues(1)) + (.y * mxValues(3)) + mxValues(5) + OffsetY
'            End With
'        Next
'
'        ' create the warped path object & clean up
'        GdipCreatePath2 VarPtr(m_PathPoints(0)), VarPtr(m_PathPtType(0)), mx, FillModeWinding, zSpt_TransformPath
'        If (m_State And ps_PointsCached) = 0& Then
'            Erase m_PathPoints
'            Erase m_PathPtType
'        Else
'            zSpt_GetPathData m_PathPoints, m_PathPtType, True ' requery because we warped the points
'        End If
    End If
    
    If ReturnPath Then
    
        ReturnPath = True
        ReturnMatrix = False
        With m_Source
            GdipCreateMatrix2 .Transformation(0), .Transformation(1), .Transformation(2), _
                          .Transformation(3), .Transformation(4), .Transformation(5), mx
            If pClone = 0& Then GdipClonePath .handle, pClone
        End With
        GdipTransformPath pClone, mx
        GdipDeleteMatrix mx
        zSpt_TransformPath = pClone
        
    ElseIf ReturnMatrix = True Then
    
        With m_Source
            GdipCreateMatrix2 .Transformation(0), .Transformation(1), .Transformation(2), _
                          .Transformation(3), .Transformation(4), .Transformation(5), zSpt_TransformPath
        End With
        
    Else
    
        zSpt_TransformPath = True
        
    End If
    
End Function


Private Function zSpt_GetBoundingBox(ByVal UseClipBox As Boolean, _
                                ByRef Left As Single, ByRef Top As Single, _
                                ByRef Width As Single, ByRef Height As Single, _
                                Optional ByVal MarkerIndex As Long = 0&) As Boolean

    ' ---=====| Returns the dimensions of a bounding box containing the path |====---

    ' UseClipBox returns smallest rectangle to surround the path.
    ' Loose fitting is the actual bounding rectangle is returned by GDI+.
    ' What is the difference?  Loose fitting bounds include bezier control points which
    ' may be several to dozens of pixels outside of the visible path.
    
    ' When to use UseClipBox?  Recommend UseClipBox=True when you need to position a path
    ' object, as it is displayed, at a specific location on the screen; especially when
    ' rotating. The math involved would look like the following:
    '   Center on DC: dcWidth=500: dcHeight=300
    '   Call GetBoundingBox(True, X, Y, Cx, Cy)
    '   Call Move(True, (dcWidth-Cx)/2, (dcHeight-Cy)/2)
    
    ' When not to use ClipBox? Don't use it when transformations/positioning
    ' should apply to the entire bounding box, not just just the clipped box (i.e., scaling)
    
    Dim pRect As RECTF, pClone As Long, I As Long
    
    If MarkerIndex Then zSpt_GetSubpathInfo True
    If m_Source.handle Then
    
        If (MarkerIndex = 0&) Or (MarkerIndex = 1& And m_SubPaths.Count = MarkerIndex) Then
            ' no subpath markers or a single path
            If UseClipBox = False Then
            
                GdipGetPathWorldBounds m_Source.handle, pRect, 0&, 0&
                
            Else ' a little creativity gets us what we want
                
                If (m_State And ps_ClipBox) = 0& Then
                    
                    ' Note: GdipWindingModeOutline gives us the clipped box, but...
                    ' If multiple open subpaths and even multiple closed subpaths exist, then
                    ' it won't always. Using GdipClosePathFigures doesn't help either.
                    
                    ' Therefore we will join all the path figures contiguously and close the
                    ' path ourselves when done. This creates a single closed path vs the
                    ' possibility of multiple closed/open path combinations. And doing so,
                    ' GdipWindingModeOutline appears to work perfectly. This is far easier
                    ' and faster than measuring the bounding box ourselves -- which can be done.
                    
                    ' get path data points & enumerate them
                    zSpt_GetPathData m_PathPoints, m_PathPtType, False
                    For I = 1 To UBound(m_PathPtType)
                        Select Case (m_PathPtType(I) And PathPointTypePathTypeMask)
                        Case PathPointTypeStart
                            ' start of new subpath: change it to continuation via a line
                            m_PathPtType(I) = PathPointTypeLine
                        Case PathPointTypeBezier
                            I = I + 2   ' offset to next path point (beziers have 3 points)
                        ' Case PathPointTypeLine -- nothing special for lines
                        End Select
                        ' check to see if this point closes the current subpath...
                        If (m_PathPtType(I) And PathPointTypeCloseSubpath) = PathPointTypeCloseSubpath Then
                            ' it does, so we remove the flag so our path continues on
                            m_PathPtType(I) = (m_PathPtType(I) And Not PathPointTypeCloseSubpath)
                        End If
                    Next
                    ' done. Tweak the final point so the path will close from the last point
                    ' to the very first point, giving us one contiguous path
                    m_PathPtType(I - 1) = (m_PathPtType(I - 1) Or PathPointTypeCloseSubpath)
                    ' create the path now
                    GdipCreatePath2 VarPtr(m_PathPoints(0)), VarPtr(m_PathPtType(0)), I, ePathBrushFillMode.FillModeWinding, pClone
                    ' get the minimal bounding box & destroy the path
                    If Not pClone = 0& Then
                        GdipWindingModeOutline pClone, 0&, 3
                        GdipGetPathWorldBounds pClone, pRect, 0&, 0&
                        GdipDeletePath pClone
                    End If
                    If (m_State And ps_PointsCachable) = 0& Then  ' release path data
                        Erase m_PathPoints()
                        Erase m_PathPtType()
                        m_State = (m_State And Not ps_PointsCached)
                    Else                                        ' requery path data
                        zSpt_GetPathData m_PathPoints(), m_PathPtType(), True ' force refresh
                    End If
                    m_clipBox = pRect
                    m_State = (m_State Or ps_ClipBox)
                Else
                    pRect = m_clipBox
                End If
            
            End If
        
        Else    ' get bounding rect for specific subPath
            pClone = zSpt_CreateSubPathPath(MarkerIndex, UseClipBox, False)
            If Not pClone = 0& Then
                If UseClipBox Then GdipWindingModeOutline pClone, 0&, 3
                GdipGetPathWorldBounds pClone, pRect, 0&, 0&
                GdipDeletePath pClone
            End If
        End If
        
    End If
    
    Left = pRect.Left
    Top = pRect.Top
    Width = pRect.Width
    Height = pRect.Height
    zSpt_GetBoundingBox = (Not (pRect.Height = 0! And pRect.Width = 0!))
        
End Function

Public Function SaveToImageArray(outArray() As Byte, ByVal ImageType As eSaveAs, _
                                Optional ByVal BackColor As Long = vbWhite, _
                                Optional ByVal JPG_Quality As Long = 70) As Boolean

    If m_Source.handle = 0& Then Exit Function
    
    Dim mx As Long, br As Long
    Dim hDib As Long, pPtr As Long
    Dim dDC As Long, hObj As Long
    Dim penWidth As Long
    
    Dim bmpi As BITMAPINFO
    Dim ptDisplayLoc As POINTF, ptOffsets As POINTF
    Dim tmpPT As POINTF, rotPT As POINTF
    Dim uRect As RECT, tRect As RECT
    Dim sRect As RECT, dRect As RECT
    
    Dim bReplaceOverrides As Boolean
    
    If (m_State And ps_TransformOffsets) = ps_TransformOffsets Then
        bReplaceOverrides = True
        m_State = ((m_State And Not ps_TransformOffsets) And Not ps_TransformMatrix)
    End If
    ' get untransformed bounding box
    Me.GetBoundingBox_Long True, sRect.Left, sRect.Top, sRect.Right, sRect.Bottom
    sRect.Right = sRect.Left + sRect.Right
    sRect.Bottom = sRect.Top + sRect.Bottom
    
    ' reset path position & get transformation matrix
    ptDisplayLoc = m_Source.DisplayLocale
    m_Source.DisplayLocale = tmpPT
    If zSpt_TransformPath(False, False) = 0 Then Exit Function
    
    ' calculate position of each point of the overall bounds
    tRect = sRect
    For mx = 0 To 3
        Select Case mx
        Case 0: rotPT.X = tRect.Left: rotPT.Y = tRect.Top       ' top left
        Case 1: rotPT.X = tRect.Right: rotPT.Y = tRect.Top      ' top right
        Case 2: rotPT.X = tRect.Right: rotPT.Y = tRect.Bottom   ' bottom right
        Case 3: rotPT.X = tRect.Left: rotPT.Y = tRect.Bottom    ' bottom left
        End Select
        With m_Source
            tmpPT.X = ((rotPT.X * .Transformation(0)) + (rotPT.Y * .Transformation(2)) + .Transformation(4))
            tmpPT.Y = ((rotPT.X * .Transformation(1)) + (rotPT.Y * .Transformation(3)) + .Transformation(5))
        End With
        ' add this point to a union rectangle for calculating overall size
        dRect.Left = tmpPT.X
        dRect.Top = tmpPT.Y
        dRect.Right = dRect.Left + 1
        dRect.Bottom = dRect.Top + 1
        UnionRect uRect, uRect, dRect
    Next
    ' cache base rectangle top/left coords for eventual rendering offsets
    ptOffsets.X = uRect.Left: ptOffsets.Y = uRect.Top
    
    ' now we repeat, offseting rectangles for any brushes/pens in use
    For br = 1 To m_BrushCount
        With m_Brushes(br)
            If Not .FillHandle = 0& Then ' brush used
                If Not (.OffsetX = 0& And .OffsetY = 0&) Then
                    tRect = sRect
                    OffsetRect tRect, .OffsetX, .OffsetY
                    .UpdateGradientBounds = True
                    If .OutlineHandle = 0& Then
                        penWidth = 1&
                    Else 'the GDI+ brush class can have both a pen & brush set simultaneously
                        .GetOutlineAttributes 0&, , penWidth
                        OffsetRect tRect, .OffsetX, .OffsetY
                        tRect.Right = tRect.Right - .OffsetX
                        tRect.Bottom = tRect.Bottom - .OffsetY
                    End If
                Else
                    penWidth = 0&
                End If
            ElseIf Not .OutlineHandle = 0& Then ' pen used
                tRect = sRect
                .GetOutlineAttributes 0&, , penWidth
                If .OffsetX = 0& And .OffsetY = 0& Then
                    OffsetRect tRect, penWidth, penWidth
                Else
                    OffsetRect tRect, penWidth * .OffsetX, penWidth * .OffsetY
                End If
            Else
                penWidth = 0&
            End If
            If penWidth Then
                For mx = 0 To 3
                    Select Case mx
                    Case 0: rotPT.X = tRect.Left: rotPT.Y = tRect.Top       ' top left
                    Case 1: rotPT.X = tRect.Right: rotPT.Y = tRect.Top      ' top right
                    Case 2: rotPT.X = tRect.Right: rotPT.Y = tRect.Bottom   ' bottom right
                    Case 3: rotPT.X = tRect.Left: rotPT.Y = tRect.Bottom    ' bottom left
                    End Select
                    With m_Source
                        tmpPT.X = ((rotPT.X * .Transformation(0)) + (rotPT.Y * .Transformation(2)) + .Transformation(4))
                        tmpPT.Y = ((rotPT.X * .Transformation(1)) + (rotPT.Y * .Transformation(3)) + .Transformation(5))
                    End With
                    ' add this point to a union rectangle for calculating overall size
                    dRect.Left = tmpPT.X
                    dRect.Top = tmpPT.Y
                    dRect.Right = dRect.Left + 1
                    dRect.Bottom = dRect.Top + 1
                    UnionRect uRect, uRect, dRect
                Next
            End If
        End With
    Next
    ' adjust rendering offsets
    ptOffsets.X = ptOffsets.X + (uRect.Left - ptOffsets.X)
    ptOffsets.Y = ptOffsets.Y + (uRect.Top - ptOffsets.Y)
    
    ' position rectangle at 0,0 coords; inflate size by 2 pixel buffer
    OffsetRect uRect, -uRect.Left, -uRect.Top
    uRect.Right = uRect.Right + 3
    uRect.Bottom = uRect.Bottom + 3
    
    ' create a bitmap to render the path to
    With bmpi.bmiHeader
        .biSize = 40
        .biPlanes = 1
        .biWidth = uRect.Right
        .biHeight = uRect.Bottom
        
        If ImageType = waSaveAsBMP Then ' save as 24bpp bitmap
            .biBitCount = 24
            .biSizeImage = .biHeight * ((((.biWidth * .biBitCount) + &H1F&) And Not &H1F&) \ &H8&)
        Else
            .biBitCount = 32           ' GDI+ will do the saving; use 32bpp
        End If
    
    End With
    hObj = GetDC(GetDesktopWindow)
    dDC = CreateCompatibleDC(hObj)
    
    If dDC = 0& Or hObj = 0& Then
        ReleaseDC GetDesktopWindow(), hObj
        If bReplaceOverrides Then m_State = ((m_State Or ps_TransformOffsets) And Not ps_TransformMatrix)
        Exit Function
    End If
    hDib = CreateDIBSection(hObj, bmpi, 0&, pPtr, 0&, 0&)
    ReleaseDC GetDesktopWindow(), hObj
    If hDib = 0& Or pPtr = 0& Then
        DeleteDC dDC
        If bReplaceOverrides Then m_State = ((m_State Or ps_TransformOffsets) And Not ps_TransformMatrix)
        Exit Function
    End If
    
    ' select DIB into DC & fill with optional background color
    hObj = SelectObject(dDC, hDib)
    If BackColor < 0& Then
        br = CreateSolidBrush(vbWhite)
    Else
        br = CreateSolidBrush(BackColor)
    End If
    FillRect dDC, uRect, br
    DeleteObject br
                
    ' render path
    Me.Offset -ptOffsets.X + 1!, -ptOffsets.Y + 1!
    Me.Render dDC
    SelectObject dDC, hObj
    
    ' replace setting as needed
    If bReplaceOverrides Then m_State = (m_State Or ps_TransformOffsets)
    ' moving a path can affect some GDI+ brush objects; tell them to refresh as needed
    Me.Move ptDisplayLoc.X, ptDisplayLoc.Y
    For br = 1 To m_BrushCount
        m_Brushes(br).UpdateGradientBounds = True
    Next
        
    ' return image to byte array
    If ImageType = waSaveAsPNG Then
       SaveToImageArray = SaveToPNG(outArray(), uRect.Right, uRect.Bottom, pPtr)
    ElseIf ImageType = waSaveAsJPG Then
       SaveToImageArray = SaveToJPG(outArray(), uRect.Right, uRect.Bottom, pPtr, JPG_Quality)
    Else ' write out the bitmap
        ReDim outArray(0 To bmpi.bmiHeader.biSizeImage + 53) ' size return array & include header+14 byte fileheader
        CopyMemory outArray(0), &H4D42&, 2& ' BMP magic number
        CopyMemory outArray(2), (54& + bmpi.bmiHeader.biSizeImage), 4& ' overall file size
        ' next 4 bytes are reserved & must be zero; just skip them
        CopyMemory outArray(10), 54&, 4& ' offset to where pixels begin
        CopyMemory outArray(14), bmpi.bmiHeader.biSize, bmpi.bmiHeader.biSize ' the header
        CopyMemory outArray(54), ByVal pPtr, bmpi.bmiHeader.biSizeImage ' the pixels
        SaveToImageArray = True
    End If
        
    ' clean up
    DeleteObject hDib
    DeleteDC dDC
    
End Function


Private Function SaveToPNG(outStream() As Byte, ByVal Width As Long, ByVal Height As Long, bitsPtr As Long) As Boolean

    ' Function uses GDI+ to create a PNG stream

    Dim uEncCLSID(0 To 3) As Long
    Dim IIStream As IUnknown
    Dim hImage As Long
    
    Const PixelFormat32bppARGB As Long = &H26200A
    
    ' Note: GdipCreateBitmapFromGdiDib does not handle 32bpp DIBs correctly
    Call GdipCreateBitmapFromScan0(Width, Height, Width * 4&, PixelFormat32bppARGB, ByVal bitsPtr, hImage)
    
    If Not hImage = 0& Then
        On Error Resume Next
        ' retrieve information GDI+ will use for conversion
        If Not pvGetEncoderClsID("image/png", uEncCLSID) = -1 Then
            ' dib is bottom up, scan0 does top down, so flip it
            GdipImageRotateFlip hImage, 6
            ' Saving to stream/array. Create a null stream (IUnknown object)
            Erase outStream
            Call CreateStreamOnHGlobal(0, 1, IIStream)
            ' have GDI+ save the 32bpp image to the IUnknown in a PNG format
            If GdipSaveImageToStream(hImage, IIStream, uEncCLSID(0&), ByVal 0&) = 0& Then
                ' now we need to get that array to pass back to client
                ArrayFromStream IIStream, outStream()
                SaveToPNG = True
            End If
        End If
        GdipDisposeImage hImage
    End If
    
End Function

Private Function SaveToJPG(outStream() As Byte, ByVal Width As Long, ByVal Height As Long, bitsPtr As Long, Optional Quality As Long = 70) As Boolean

    ' Function uses GDI+ to create a JPG stream

    Dim uEncCLSID(0 To 3) As Long
    Dim aEncParams() As Byte
    Dim IIStream As IUnknown
    Dim uEncParams As EncoderParameters
    Dim hImage As Long
    
    Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    Const PixelFormat32bppARGB As Long = &H26200A
    
    ' Note: GdipCreateBitmapFromGdiDib does not handle 32bpp DIBs correctly
    Call GdipCreateBitmapFromScan0(Width, Height, Width * 4&, PixelFormat32bppARGB, ByVal bitsPtr, hImage)
    
    If Not hImage = 0& Then
        On Error Resume Next
        ' retrieve information GDI+ will use for conversion

        If Not pvGetEncoderClsID("image/jpeg", uEncCLSID) = -1 Then
            '-- Set encoder params. (Quality)
            uEncParams.Count = 1
            ReDim aEncParams(1 To Len(uEncParams))
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                Call CLSIDFromString(StrPtr(EncoderQuality), .GUID(0))
                If Quality < 30 Then
                    Quality = 30
                ElseIf Quality > 100 Then
                    Quality = 100
                End If
                .Value = VarPtr(Quality)
            End With
            Call CopyMemory(aEncParams(1), uEncParams, Len(uEncParams))
            
            GdipImageRotateFlip hImage, 6
            Erase outStream
            Call CreateStreamOnHGlobal(0, 1, IIStream)
            If GdipSaveImageToStream(hImage, IIStream, uEncCLSID(0&), aEncParams(1)) = 0& Then
                ' now we need to get that array to pass back to client
                ArrayFromStream IIStream, outStream()
                SaveToJPG = True
            End If
        End If
        GdipDisposeImage hImage
    End If

End Function

Private Function pvGetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
  ' supporting routine for SaveToPNG
  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0&) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For lIdx = 1& To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            CopyMemory ClassID(0), ICI(lIdx).ClassID(0), 16& ' Save the Class ID
            pvGetEncoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function pvPtrToStrW(ByVal lpsz As Long) As String
  ' supporting routine for SaveToPNG; converts String Pointer to String
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0&) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2&)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
    
End Function

Private Function ArrayFromStream(Stream As IUnknown, arrayBytes() As Byte) As Boolean

    ' Purpose: Return the array contained in an IUnknown interface
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    
    If Not Stream Is Nothing Then
    
        If GetHGlobalFromStream(ByVal ObjPtr(Stream), o_hMem) = 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                    ArrayFromStream = True
                End If
            End If
        End If

    End If
    
End Function

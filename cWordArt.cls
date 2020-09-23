VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWordArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



' Some more fun with GDI+ paths.
' Unicode compatible on Win2K & above if regional settings allow unicode support
' :: see the included PNG file: RegionSettings.png

' How to use it, the class needs several pieces of information, each of these must be called
' before any output is allowed:

' 1. Set the string to be output.
'    -- The font size is used for non-stretching perspectives
' 2. Create the Top and Bottom Guide Paths
'    -- Use the various AppendGuide_xxxx functions
' 3. Set any additional properties you wish
' 4. Call RenderPath to display it on screen

' Important comments regarding guide paths
' ==============================================================
' -- These are GDI+ paths and should not be closed paths with few exceptions (i.e., Circle/Oval)
'   :: guides must be created in the direction the text is expected to flow (left to right,
'       top to bottom for example). Exception. If unicode is RTL, then routines will detect
'       that and reverse the path internally. Assumption is system can display the unicode text
' -- The Top guide limits & contours the top edge of the characters
' -- The Bottom guide limits & contours the bottom edge of the characters
' -- The guides are just guides; there is no guarantee that the text will be perfectly pixel aligned to
'       either the top or bottom edges of the guide. However, alignment is near-perfect depending on
'       the perspective property settings and/or font choices. When rendering with outline pens,
'       the pens will always exceed the guides. Pens are decoration that are rendered around the
'       path; whereas brushes are contained within the path.

' Having 2 guides allows many shapes and also contours the text. Mixing them with the perspective
' settings and/or warp style properties can produce endless combinations of path results.

' Note: Ellipses and Circles.  If we look at a circle, say a 12-hour clock, where is zero degrees? Is it
' at the 12 hand, the 9 or the 3 hand?  GDI+ makes it the 3 hand and ellipses always travel clockwise.
' Non-RTL languages would look for text beginning at the 9 or 12 hand.  Keep this in mind when adding
' these to your path. You may have to opt for using the Arc shape which allows you to specify a starting
' angle and use of negative arc sweeps if necessary, when using in a RTL environment


' ---=====| API Constants & Enumerations |====---
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
Private Enum gdiPathPointType
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
Public Enum FillMode
    FillModeAlternate = &H0
    FillModeWinding = &H1
End Enum
Private Enum GpUnit
    UnitWorld = &H0
    UnitDisplay = &H1
    UnitPixel = &H2
    UnitPoint = &H3
    UnitInch = &H4
    UnitDocument = &H5
    UnitMillimeter = &H6
End Enum
Public Enum LinearGradientMode      ' gradient brush options
    LinearGradientModeHorizontal = 0
    LinearGradientModeVertical = 1
    LinearGradientModeForwardDiagonal = 2
    LinearGradientModeBackwardDiagonal = 3
    RadialGradientCenterOut = 4     ' << custom, not GDI+ standard
End Enum
Public Enum WarpModes       ' path warping options
    warpPerspective = 0
    warpBilinear = 1        ' default
End Enum


' ---=====| API DECLARATIONS |====---

' ---=====| GDI+ PATH RELATED |====---
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As FillMode, ByRef mpath As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipAddPathBeziers Lib "GdiPlus.dll" (ByVal Path As Long, ByRef pPoints As POINTF, ByVal Count As Long) As Long
Private Declare Function GdipAddPathLine2 Lib "GdiPlus.dll" (ByVal Path As Long, ByRef pPoints As POINTF, ByVal Count As Long) As Long
Private Declare Function GdipCreateMatrix Lib "GdiPlus.dll" (ByRef mMatrix As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "GdiPlus.dll" (ByVal mMatrix As Long, ByVal mOffsetX As Single, ByVal mOffsetY As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipDeleteMatrix Lib "GdiPlus.dll" (ByVal mMatrix As Long) As Long
Private Declare Function GdipTransformPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mMatrix As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipSetPathFillMode Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mFillMode As FillMode) As Long
Private Declare Function GdipGetPathData Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPathData As PATHDATA) As Long
Private Declare Function GdipGetPointCount Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePath2 Lib "GdiPlus.dll" (ByVal mPOINTFptr As Long, ByVal mBYTEptr As Long, ByVal mCount As Long, ByVal mFillMode As FillMode, ByRef mpath As Long) As Long
Private Declare Function GdipAddPathBezier Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single, ByVal mX3 As Single, ByVal mY3 As Single, ByVal mX4 As Single, ByVal mY4 As Single) As Long
Private Declare Function GdipAddPathPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mAddingPath As Long, ByVal mConnect As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipAddPathLine Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipAddPathArc Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipGetPathWorldBounds Lib "GdiPlus.dll" (ByVal Path As Long, ByRef Bounds As RECTF, ByVal matrix As Long, ByVal Pen As Long) As Long
Private Declare Function GdipClearPathMarkers Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipClonePath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mClonePath As Long) As Long
Private Declare Function GdipIsVisiblePathPoint Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mx As Single, ByVal mY As Single, ByVal mGraphics As Long, ByRef mResult As Long) As Long

' ---=====| GDI+ BRUSHES/PENS |====---
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long, ByRef mCount As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (ByRef pRect As RECTF, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Mode As Long, ByVal WrapMode As Long, ByRef lineGradient As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As GpUnit, ByRef mPen As Long) As Long
Private Declare Function GdipGetLineColors Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long) As Long
Private Declare Function GdipGetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long) As Long
Private Declare Function GdipGetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipSetPathGradientCenterPoint Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mPoints As Any) As Long

' ---=====| GDI+ DRAWING |====---
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMode As SmoothingMode) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal interpolation As InterpolationMode) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixOffsetMode As PixelOffsetMode) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long

' ---=====| NON-GDI+ API DECLARATIONS & CONSTANTS |====---
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


' ---=====| API USER-DEFINED TYPES |====---

Private Type PATHDATA           ' GDI+ UDT
    Count As Long
    PointsPTR As Long
    TypesPtr As Long
End Type
Private Type RECTF              ' GDI+ UDT
   X As Single
   Y As Single
   mWidth As Single
   mHeight As Single
End Type
Private Type POINTF             ' GDI+ UDT
    X As Single
    Y As Single
End Type

' ---=====| CUSTOM USER-DEFINED TYPES |====---
Private Type LINEARCURVEREF
    Distance(0 To 100) As Single    ' time to distance lookup. Must be 100 elements.
    Index As Long                   ' ^^ if changed, change NavigatePath to accommodate
End Type
Private Type BLOCKTEXTDATA          ' 1 structure per line of text
    yOffset As Single               ' difference between top of path characters & top of bounds
    Length As Single                ' adjusted width of display path
    blockHeight As Single           ' total height required for line of text
End Type
    

' ---====| CUSTOM PUBLIC ENUMERATIONS |====---
Public Enum eErrorCodes
    waErrInvalidGuide_Top = 1
    waErrInvalidGuide_Btm = 2
    waErrUnsupportedFont = 3
    waErrTextNotSet = 4
    waErrInvalidDC = 5
    waErrUnknown = 99
End Enum
Public Enum eAlignmentEnum     ' path alignment
    waAlignLeft = 0            ' ^^ can be combined with OR
    waAlignRight = 1
    waAlignCenterHorizontal = 2
    waAlignStretchHorizontal = 4   ' default
    waAlignStretchVertical = 8     ' default
    waAlignAllCharsSameSizeVertical = 16
End Enum
Public Enum ePathOrder
    TopLeft_Clockwise = 0 ' left-top, right-top, right-bottom, left-bottom
    TopLeft_zPattern = 1  ' left-top, right-top, left-bottom, right-bottom
End Enum
Public Enum eImageFormats
    waFormatJPG = 0
    waFormatBMP = 1
    waFormatPNG = 2
End Enum

' ---=====| CUSTOM CLASS-ONLY ENUMERATIONS |====---
Private Enum StatusEnum
    waStat_Text = 1
    waStat_GuideBtm = 2
    waStat_GuideTop = 4
    waStat_Ok = 7           ' each of the above are set
    waStat_GuideCombined = 8
    waStat_RebuildPath = 32
    waStat_DoFillBrush = 64
    waStat_DoShadowBrush = 128
End Enum

' ---=====| CLASS-LEVEL VARIABLES |====---
Private m_Font As StdFont           ' font used to create path
Private m_Text As String            ' text used for path
Private m_Style As WarpModes        ' (pending) warp style both for individual characters & entire path as needed
Private m_Alignment As eAlignmentEnum ' horizontal/vertical alignment options
Private m_IndentR As Long           ' left edge indentation
Private m_IndentL As Long           ' right edge indentation
Private m_Status As StatusEnum      ' flags indicating various states of the path
Private m_LineSpacing As Long       ' optional line spacing for multi-line captions. Default=3
Private m_CharSpacing As Long       ' optional inter-character spacing. Default is zero

Private m_Source As clsWApath       ' the text path
Private m_Guides As clsWApath       ' the combined top/bottom guides
Private m_guideTop As clsWApath     ' temporary top guide path
Private m_guideBtm As clsWApath     ' temporary bottom guide path

Private m_AnchorPts As POINTF      ' left/top screen position for rendered path

' Tracking related variables (tracking used for mouse-driven sizing/rotation)
Private WithEvents m_Tracker As clsWATracker
Attribute m_Tracker.VB_VarHelpID = -1
Private client_Tracker As ITrackingCallback
Attribute client_Tracker.VB_VarHelpID = -1
Private m_BackupPath As clsWApath  ' rendered path used for hit testing as needed
Private m_TrackingKey As String    ' optional key to be returned in tracking callbacks


' ---=====| INITIALIZING PROPERTIES |====---
Private Sub Class_Initialize()
    
    Set m_Source = New clsWApath
    Set m_Guides = New clsWApath
    Set Font = Nothing          ' create default font
    Me.WarpStyle = warpBilinear    ' default warp style
    Me.Alignment = waAlignStretchHorizontal Or waAlignStretchVertical
    Me.LineSpacing = 3

End Sub

' ---=====| CLEAN UP|====---
Private Sub Class_Terminate()
    
    ' Note: Hitting STOP or executing END statements prevents clean up
    ' of resources, leaks, and potential crashes
    If Not m_Tracker Is Nothing Then m_Tracker.TerminateTracking False
    Set client_Tracker = Nothing
    Set m_Tracker = Nothing
    Set m_Guides = Nothing
    Set m_Source = Nothing
    Set m_BackupPath = Nothing
End Sub

' ---=====| FONT SETTINGS |====---
Public Property Let Font(vbFont As StdFont)
    Set Font = vbFont
End Property
Public Property Set Font(ByVal vbFont As StdFont)
    ' Use true type fonts only
    Set m_Font = vbFont
    If vbFont Is Nothing Then       ' declare default font
        Set m_Font = New StdFont
        m_Font.name = "Tahoma"
        m_Font.Size = 24
    End If
    m_Status = (m_Status Or waStat_RebuildPath)
End Property
Public Property Get Font() As StdFont
    If m_Font Is Nothing Then Set Font = Nothing
    Set Font = m_Font
End Property

' ---=====| CAPTION SETTINGS |====---
Public Property Let Caption(ByVal sText As String)
    If m_Text <> sText Then
        ' remove any leading carriage returns
        Do
            If Left$(sText, 2) = vbCrLf Then
                sText = Mid$(sText, 3)
            Else
                Exit Do
            End If
        Loop
        m_Text = sText
        If m_Text = vbNullString Then
            m_Status = (m_Status And Not waStat_Text)
        Else
            m_Status = (m_Status Or waStat_Text Or waStat_RebuildPath)
        End If
    End If
End Property
Public Property Get Caption() As String
    Caption = m_Text
End Property

' ---=====| LINE SPACING SETTING |====---
Public Property Let LineSpacing(ByVal lSize As Long)
    If lSize > -1& Then
        If m_LineSpacing <> lSize Then
            m_LineSpacing = lSize
            m_Status = (m_Status Or waStat_RebuildPath)
        End If
    End If
End Property
Public Property Get LineSpacing() As Long
    LineSpacing = m_LineSpacing
End Property

' ---=====| CHARACTER SPACING SETTING |====---
Public Property Let CharacterSpacing(ByVal lSize As Long)
    ' When characters have thick pens drawing their edges, default spacing btwn the
    ' characters may be too small and the the borders can overlap. This option allows
    ' ability to spread out characters a few pixels more than normal. When set,
    ' the spacing is applied to both left & right edges of the character
    If lSize > -1& Then
        If m_CharSpacing <> lSize Then
            m_CharSpacing = lSize
            m_Status = (m_Status Or waStat_RebuildPath)
        End If
    End If
End Property
Public Property Get CharacterSpacing() As Long
    CharacterSpacing = m_CharSpacing
End Property

' ---=====| WARP STYLE SETTINGS (PENDING) |====---
Public Property Let WarpStyle(ByVal Style As WarpModes)
    ' Warp style applies to both path & individual characters
    If Not Style = warpPerspective Then Style = warpBilinear
    If Not Style = m_Style Then
        m_Style = Style
        m_Status = (m_Status Or waStat_RebuildPath)
    End If
End Property
Public Property Get WarpStyle() As WarpModes
    WarpStyle = m_Style
End Property

' ---=====| ALIGNMENT SETTINGS |====---
Public Property Let Alignment(ByVal pAnchor As eAlignmentEnum)
    ' When Left, Center or Right alignment is chosen, the actual
    ' font width is used if the text can be contained within the
    ' path. If not, the routines default to horizontal stretching
    If pAnchor >= waAlignLeft Then
        If pAnchor <> m_Alignment Then
            m_Alignment = pAnchor
            m_Status = (m_Status Or waStat_RebuildPath)
        End If
    End If
End Property
Public Property Get Alignment() As eAlignmentEnum
    Alignment = m_Alignment
End Property

' ---=====| LEFT INDENTATION SETTING |====---
Public Property Let IndentRight(ByVal vIndent As Single)
    ' Indentation offsets left edge of path & reduces overall path width
    If vIndent <> m_IndentR Then
        m_IndentR = vIndent
        m_Status = (m_Status Or waStat_RebuildPath)
    End If
End Property
Public Property Get IndentRight() As Single
    IndentRight = m_IndentR
End Property

' ---=====| RIGHT INDENTATION SETTING |====---
Public Property Let IndentLeft(ByVal vIndent As Single)
    ' Indentation offsets right edge of path & reduces overall path width
    If vIndent <> m_IndentL Then
        m_IndentL = vIndent
        m_Status = (m_Status Or waStat_RebuildPath)
    End If
End Property
Public Property Get IndentLeft() As Single
    IndentLeft = m_IndentL
End Property


' ---=====| PUBLIC FUNCTIONS |====---

Public Property Let Brush(ByVal Key As String, hBrush As clsGDIbrush)
    Set Brush(Key) = hBrush
End Property
Public Property Set Brush(ByVal Key As String, hBrush As clsGDIbrush)
    Set m_Source.Brush(Key) = hBrush
End Property
Public Property Get Brush(ByVal Key As String) As clsGDIbrush
    Set Brush = m_Source.Brush(Key)
End Property
Public Function BrushZOrder(ByVal Key As String, ByVal FirstBrush As Boolean) As Boolean
    BrushZOrder = m_Source.BrushZOrder(Key, FirstBrush)
End Function

Public Sub ResetGuides(Optional ByVal fFillType As FillMode = FillModeWinding)
    ' ---=====| Clears the Guides |====---
    
    ' tst code remove
    If Not m_Tracker Is Nothing Then Call Tracking_End(True)
    m_Guides.Clear
    Set m_guideBtm = Nothing
    Set m_guideTop = Nothing
    m_Status = ((m_Status And waStat_Text) Or waStat_DoFillBrush Or waStat_DoShadowBrush Or waStat_RebuildPath)
End Sub

Public Function SaveAsBMP(outArray() As Byte, Optional BkgColor As Long = vbWhite) As Boolean
    If Not m_Source Is Nothing Then
        SaveAsBMP = m_Source.SaveToImageArray(outArray(), waSaveAsBMP, BkgColor)
    End If
End Function
Public Function SaveAsPNG(outArray() As Byte, Optional BkgColor As Long = vbWhite) As Boolean
    If Not m_Source Is Nothing Then
        SaveAsPNG = m_Source.SaveToImageArray(outArray(), waSaveAsPNG, BkgColor)
    End If
End Function
Public Function SaveAsJPG(outArray() As Byte, Optional BkgColor As Long = vbWhite, Optional Quality As Long = 70) As Boolean
    If Not m_Source Is Nothing Then
        SaveAsJPG = m_Source.SaveToImageArray(outArray(), waSaveAsJPG, BkgColor, Quality)
    End If
End Function

Public Function AppendGuide_Line(ByVal TopGuide As Boolean, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Boolean
    
    ' ---=====| Adds a line to a path. If no path exists, it is created |====---
    
    Dim Path As clsWApath
    Set Path = AppendGuide(TopGuide, True)
    ' attempt to add the line
    If Path.Append_Line(X1, Y1, X2, Y2) Then
        AppendGuide TopGuide, False
        AppendGuide_Line = True
    End If
End Function

Public Function AppendGuide_Bezier(ByVal TopGuide As Boolean, ByVal X1 As Single, ByVal Y1 As Single, _
                    ByVal CtrlPtX1 As Single, ByVal CtrlPtY1 As Single, _
                    ByVal CtrlPtX2 As Single, ByVal CtrlPtY2 As Single, _
                    ByVal X2 As Single, ByVal Y2 As Single) As Boolean
    
    ' ---=====| Adds a bezier curve to a path. If no path exists, it is created |====---
    
    Dim Path As clsWApath
    Set Path = AppendGuide(TopGuide, True)
    ' attempt to add the curve
    If Path.Append_Bezier(X1, Y1, CtrlPtX1, CtrlPtY1, CtrlPtX2, CtrlPtY2, X2, Y2) Then
        AppendGuide TopGuide, False
        AppendGuide_Bezier = True
    End If
End Function

Public Function AppendGuide_Arc(ByVal TopGuide As Boolean, ByVal X As Single, Y As Single, _
                    ByVal arcWidth As Single, ByVal arcHeight As Single, _
                    ByVal StartAngle As Single, ByVal ArcAngle As Single) As Boolean
    
    ' ---=====| Adds an arc to a path. If no path exists, it is created |====---
    
    ' Notes
    ' :: for left to right start arc at 180 vs 0
    '   - For an upwards arc, make arcAngle positive: 180
    '   - For a downards arc, make arcAngle negative: -180
    ' :: for right to left, start arc at 0 vs 180
    '   - For an upwards arc, make arcAngle negative: -180
    '   - For a downards arc, make arcAngle positive: 180
    Dim Path As clsWApath
    Set Path = AppendGuide(TopGuide, True)
    ' attempt to add the arc
    If Path.Append_Arc(X, Y, arcWidth, arcHeight, StartAngle, ArcAngle) Then
        AppendGuide TopGuide, False
        AppendGuide_Arc = True
    End If

End Function

Public Function AppendGuide_Circle(ByVal TopGuide As Boolean, ByVal X As Single, Y As Single, _
                    ByVal Diameter As Single, Optional ByVal StartAtX As Boolean = True) As Boolean
                    
    ' ---=====| Adds a circle to a path. If no path exists, it is created |====---
    
    ' Notes
    ' :: clockwise circles from the X coordinate, set StartAtX = True
    ' :: clockwise circles from the X+Diameter coordinate, set StartAtX = False
    Dim Path As clsWApath, Angle As Single
    Set Path = AppendGuide(TopGuide, True)
    ' attempt to add the circle
    If StartAtX Then Angle = 180 Else Angle = 0
    If Path.Append_Arc(X, Y, Diameter, Diameter, Angle, 360) Then
        AppendGuide TopGuide, False
        AppendGuide_Circle = True
    End If

End Function

Public Function AppendGuide_EllipseSplit(ByVal X As Single, Y As Single, _
                    ByVal Width As Single, ByVal Height As Single, Optional ByVal StartAtX As Boolean = True) As Boolean
                    
    ' ---=====| Adds an arc to each the bottom & top guide path. If no path exists, it is created |====---
    
    ' Creates both the top and bottom guide paths.
    ' Top guide is assigned to top half of ellipse, bottom guide assigned to bottom half
                        
    ' Notes
    ' :: clockwise ellipse from the X coordinate, set StartAtX = True
    ' :: clockwise ellipse from the X+Width coordinate, set StartAtX = False
    Dim Angle As Single
    AppendGuide True, True
    AppendGuide False, True
    ' attempt to add the arcs
    If StartAtX Then Angle = 180 Else Angle = 0
    If m_guideTop.Append_Arc(X, Y, Width, Height, Angle, 180) Then
        AppendGuide True, False
        If m_guideBtm.Append_Arc(X, Y, Width, Height, Angle, -180) Then
            AppendGuide False, False
            AppendGuide_EllipseSplit = True
        End If
    End If

End Function

Public Function AppendGuide_Ellipse(ByVal TopGuide As Boolean, ByVal X As Single, Y As Single, _
                    ByVal Width As Single, ByVal Height As Single, Optional ByVal StartAtX As Boolean = True) As Boolean
                    
    ' ---=====| Adds an ellipse to a path. If no path exists, it is created |====---
    
    ' Notes
    ' :: clockwise ellipse from the X coordinate, set StartAtX = True
    ' :: clockwise ellipse from the X+Width coordinate, set StartAtX = False
    Dim Path As clsWApath, Angle As Single
    Set Path = AppendGuide(TopGuide, True)
    If StartAtX Then Angle = 180 Else Angle = 0
    If Path.Append_Arc(X, Y, Width, Height, Angle, 180) Then
        If Path.Append_Arc(X, Y, Width, Height, Abs(Angle - 180), 180) Then
            AppendGuide TopGuide, False
            AppendGuide_Ellipse = True
        End If
    End If

End Function

Public Function AppendGuide_GDIpPath(ByVal TopGuide As Boolean, ByVal thePath As Long, Optional ByVal JoinToExisting As Boolean = True) As Boolean
    
    ' ---=====| Adds a passed GDI+ path to a path. If the guide path does not exist, it is created |====---
    
    Dim Path As clsWApath, nrPts As Long
    GdipGetPointCount thePath, nrPts    ' validate the parameter
    If nrPts > 1 Then
        Set Path = AppendGuide(TopGuide, True)
        If Path.Append_Path(thePath, JoinToExisting) Then
            AppendGuide TopGuide, False
            AppendGuide_GDIpPath = True
        End If
    End If
    
End Function

Public Function AppendGuide_Wave(ByVal TopGuide As Boolean, ByVal X1 As Single, ByVal Y1 As Single, _
                            ByVal X2 As Single, ByVal Y2 As Single, ByVal CrestHeight As Long, _
                            Optional ByVal WaveCount As Long = 1) As Boolean
    
    ' ---=====| Creates waves. If the guide path does not exist, it is created |====---
    Dim Path As clsWApath
    
    If WaveCount < 1 Then Exit Function
    Set Path = AppendGuide(TopGuide, True)
    If Path.Append_Wave(X1, Y1, X2, Y2, CrestHeight, WaveCount) Then
        AppendGuide TopGuide, False
        AppendGuide_Wave = True
    End If

End Function

Public Function AppendGuide_Spline(ByVal TopGuide As Boolean, ByVal X1 As Single, ByVal Y1 As Single, _
                            ByVal X2 As Single, ByVal Y2 As Single, _
                            ByVal CrestLocPct As Single, ByVal CrestHeight As Single) As Boolean
    
    ' ---=====| Creates a spline. If the guide path does not exist, it is created |====---
    
    ' CrestLocPct is the percentage of the spline where the crest/arc apex exists
    ' The percentage is in a range from 0 to 1
    
    Dim Path As clsWApath, ptWave(0 To 3) As POINTF, ctrY As Single
    Dim Cx As Single, Cy As Single, T As Single
    
    Set Path = AppendGuide(TopGuide, True)
    ' :: start point
    ptWave(0).X = X1: ptWave(0).Y = Y1
    ' :: end point
    ptWave(3).X = X2: ptWave(3).Y = Y2
    ' :: control point #1
    If X1 = X2 Then ' vertical line
        Cy = Y2 - Y1
        ptWave(1).X = X1 + CrestHeight
        ptWave(1).Y = Y1 + CrestLocPct * Cy
    Else
        Cx = X2 - X1
        Cy = Y2 - Y1
        ptWave(1).X = X1 + CrestLocPct * Cx
        ptWave(1).Y = Y1 + CrestHeight + CrestLocPct * Cy
    End If
    ' :: control point #2 = #1
    ptWave(2) = ptWave(1)
    
    ' attempt to add wave to path
    If Path.Append_BeziersPTR(VarPtr(ptWave(0)), 4) Then
        AppendGuide TopGuide, False
        AppendGuide_Spline = True
    End If
    
End Function

Public Function Render(ByVal hDC As Long, _
                    Optional ByVal OffsetX As Single = 0!, _
                    Optional ByVal OffsetY As Single = 0!, _
                    Optional ByVal OutlineOnly As Boolean, _
                    Optional ByVal OutlinePenColor As Long) As Long
                    

    ' ---=====| Renders the path using pens/brushes, coordinates & properties previously established |====---
    
    ' NOTE: If this fails, it is because the text and/or path guides are not created

    If hDC = 0& Then Exit Function
    If (m_Status And waStat_Ok) <> waStat_Ok Then
        If (m_Status And waStat_GuideBtm) = 0& Then
            Render = waErrInvalidGuide_Btm
        ElseIf (m_Status And waStat_GuideTop) = 0& Then
            Render = waErrInvalidGuide_Top
        ElseIf (m_Status And waStat_Text) = 0& Then
            Render = waErrTextNotSet
        Else
            Render = waErrUnknown
        End If
        Exit Function
    End If
   
    Dim lErr As Long
    Dim hGraphics As Long
    
    ' rebuild the path if needed
    If (m_Status And waStat_RebuildPath) Then
        lErr = BuildDisplayPath
        If lErr Then
            Render = lErr
            Exit Function
        End If
    End If
    
    If OutlineOnly Then
        
        m_Source.RenderOutline_XOR hDC, OutlinePenColor, , psXOR_Dotted, OffsetX, OffsetY
    
    Else
        ' does this path need to be warped? Only true if destination coords explicitly set
        If GdipCreateFromHDC(hDC, hGraphics) = 0& Then
        
            GdipSetInterpolationMode hGraphics, InterpolationModeHighQualityBicubic
            GdipSetPixelOffsetMode hGraphics, PixelOffsetModeHighQuality
            GdipSetSmoothingMode hGraphics, SmoothingModeHighQuality
            
            m_Source.Render hDC, OffsetX, OffsetY, hGraphics
            
            ' clean up
            GdipDeleteGraphics hGraphics
        End If
        
    End If

End Function

Public Sub RenderGuides(ByVal hDC As Long, ByVal PenColor As Long, _
                    Optional ByVal Thickness As Long = 1, _
                    Optional ByVal PenOpacity As Long = 100, _
                    Optional OutLiineOnly As Boolean = False)
    
    ' ---=====| Renders path guides in passed pen color & thickness |====---
    
    If hDC = 0& Then Exit Sub
    If CombineGuides = False Then Exit Sub
    If (m_Status And waStat_RebuildPath) = waStat_RebuildPath Then BuildDisplayPath
    
    If OutLiineOnly Then
        m_Guides.RenderOutline_XOR hDC, PenColor, Thickness, psXOR_Dotted, 0, 0
    Else
        If PenColor < 0 Then PenColor = vbBlack
        If Thickness < 0 Then Thickness = 1
        
        m_Guides.Brush("Outline").SetOutlineAttributes PenColor, PenOpacity, Thickness, DashStyleDot
            
        m_Guides.Render hDC
        ' clean up
        Set m_Guides.Brush("Outline") = Nothing
    End If
    
End Sub


Public Function IsPathPoint(ByVal hDC As Long, ByVal X As Single, ByVal Y As Single) As Boolean

    IsPathPoint = m_Source.IsMouseOverPath(hDC, X, Y)
    
End Function

Public Function Offset(ByVal X As Single, ByVal Y As Single) As Boolean

    ' ---=====| Offsets the guide paths by passed X,Y offset values |====---
    
    If CombineGuides Then
        If (m_Status And waStat_RebuildPath) = waStat_RebuildPath Then BuildDisplayPath
        ' keep the guides & display paths aligned to each other
        m_Guides.Offset X, Y
        m_AnchorPts.X = m_AnchorPts.X + X
        m_AnchorPts.Y = m_AnchorPts.Y + Y
        AlignDisplayToGuide ' function forces guides & display to stay aligned
        Offset = True
    End If
    
End Function

Public Function Move(ByVal X As Single, ByVal Y As Single, Optional ByVal MoveOnCenter As Boolean) As Boolean

    ' ---=====| Moves the guide paths to passed X,Y values |====---

    If CombineGuides Then
    
        If (m_Status And waStat_RebuildPath) = waStat_RebuildPath Then BuildDisplayPath
        
        m_Guides.Move X, Y, , MoveOnCenter
        m_AnchorPts.X = m_Guides.DisplayLeft
        m_AnchorPts.Y = m_Guides.DisplayTop
        AlignDisplayToGuide ' function forces guides & display to stay aligned
        Move = True
    
    End If

End Function

Public Function GetBoundingRect(Left As Single, Top As Single, Width As Single, Height As Single) As Boolean
    
    ' ---=====| Returns a bounding rectangle around the guides |====---
    
    If CombineGuides Then
        If (m_Status And waStat_RebuildPath) = waStat_RebuildPath Then BuildDisplayPath
        m_Guides.GetBoundingBox True, Left, Top, Width, Height
        GetBoundingRect = True
    Else
        Width = 0!: Height = Width
        Left = Width: Top = Width
    End If

End Function

Public Function GetBoundingPoints_PTR(ByVal UseClipBox As Boolean, ByVal PointsPTR As Long, Optional Pt8Format As Boolean = False) As Long
    
    Dim pt8(0 To 11) As POINTF, p As Long
    Dim Cx As Single, Cy As Single
    
    If Not m_Guides.GetBoundingPoints_PTR(UseClipBox, True, VarPtr(pt8(8))) = 0 Then
        If Pt8Format Then
            For p = 0 To 7
                Select Case p
                Case 0
                    Cx = pt8(11).X - pt8(8).X: Cy = pt8(11).Y - pt8(8).Y
                    pt8(p).X = pt8(8).X + 0.5 * Cx
                    pt8(p).Y = pt8(8).Y + 0.5 * Cy
                Case 1: pt8(p) = pt8(8)
                Case 2
                    Cx = pt8(9).X - pt8(8).X: Cy = pt8(9).Y - pt8(8).Y
                    pt8(p).X = pt8(8).X + 0.5 * Cx
                    pt8(p).Y = pt8(8).Y + 0.5 * Cy
                Case 3: pt8(p) = pt8(9)
                Case 4
                    Cx = pt8(10).X - pt8(9).X: Cy = pt8(10).Y - pt8(9).Y
                    pt8(p).X = pt8(9).X + 0.5 * Cx
                    pt8(p).Y = pt8(9).Y + 0.5 * Cy
                Case 5: pt8(p) = pt8(10)
                Case 6
                    Cx = pt8(11).X - pt8(10).X: Cy = pt8(11).Y - pt8(10).Y
                    pt8(p).X = pt8(10).X + 0.5 * Cx
                    pt8(p).Y = pt8(10).Y + 0.5 * Cy
                Case 7: pt8(p) = pt8(11)
                End Select
            Next
            CopyMemory ByVal PointsPTR, pt8(0), 64&
            GetBoundingPoints_PTR = 8
        Else
            CopyMemory ByVal PointsPTR, pt8(8), 32&
            GetBoundingPoints_PTR = 4
        End If
    End If
End Function


Public Property Get Angle() As Single
    Angle = m_Guides.Angle
End Property
Public Property Let Angle(ByVal newAngle As Single)
    If newAngle <> m_Guides.Angle Then
        m_Guides.Angle = newAngle
        m_Source.Angle = newAngle
    End If
End Property

Public Property Get ScaleX() As Single
    ScaleX = m_Guides.ScaleX
End Property
Public Property Let ScaleX(ByVal theScale As Single)
    If theScale <> m_Guides.ScaleX Then
        m_Guides.ScaleX = theScale
        m_Source.ScaleX = theScale
        AlignDisplayToGuide
    End If
End Property
Public Property Get ScaleY() As Single
    ScaleY = m_Guides.ScaleY
End Property
Public Property Let ScaleY(ByVal theScale As Single)
    If theScale <> m_Guides.ScaleY Then
        m_Guides.ScaleY = theScale
        m_Source.ScaleY = theScale
        AlignDisplayToGuide
    End If
End Property

' ---=====| PRIVATE/HELPER FUNCTIONS |====---


Private Function BuildDisplayPath() As Long

    ' ---=====| Creates a display path with user-defined properties, font, & guide paths |====---

    If CombineGuides = False Then   ' can we even get started?
        BuildDisplayPath = waErrInvalidGuide_Top
        Exit Function
    End If
    
    
    ' A bit of a long routine, and very complicated but fairly easy to follow with the remarks
    Dim NextXbtm As Single, NextXtop As Single              ' position of next character in path
    Dim pathLenBtm As Single, pathLenTop As Single
    Dim pathScaleBtm As Single, pathScaleTop As Single
    Dim gRect As RECTF, fRect As RECTF                      ' path boundaries
    
    Dim lineData() As BLOCKTEXTDATA                         ' some metrics per line of text
    Dim lSubPath As Long
    
    Dim charSrcPT() As POINTF, charDstPT() As POINTF        ' final destination coords for a character
    Dim sText() As String                                   ' multiline text
    
    Dim I As Long, LineIndex As Long, LineCount As Long     ' loop variable
    Dim stringPath As clsWApath                             ' line of text as a path
    Dim charPath As clsWApath                               ' character path extracted from stringPath
    
    Dim ttlHeight As Single, ttlCtrlHt As Single            ' text height & the tallest character in the path
    Dim tmpVal1 As Single, tmpVal2 As Single                ' general purpose temp variable
    Dim WarpOffsetTop As Single, WarpOffsetBtm As Single    ' top/bottom Y offsets when warping to guide
    
    Dim charSpacingOffset As Long, isRTL As Boolean         ' positioning offsets
    Dim MarginBtm As Single, MarginTop As Single            ' more positioning offsets
    Dim fontStyle As eFontStyle, mIndex As Long             ' font style & path character/marker index
    
    sText = Split(m_Text, vbCrLf)               ' remove trailing blank lines, if any
    For LineIndex = UBound(sText) To 0 Step -1
        If Len(sText(LineIndex)) Then
            LineCount = LineIndex + 1
            Exit For
        End If
    Next
    If LineCount = 0& Then                      ' validate we have something to draw
        BuildDisplayPath = waErrTextNotSet
        Exit Function
    End If
    
    ' generate 1st line of text as a path
    With m_Font
        If .Bold Then fontStyle = fontStyle Or FontStyleBold
        If .Italic Then fontStyle = fontStyle Or FontStyleItalic
        ' underlining & strike through not supported; requires far more work
'        If .Underline Then fontStyle = fontStyle Or FontStyleUnderline
'        If .Strikethrough Then fontStyle = fontStyle Or FontStyleStrikeout
    End With
    
    Set stringPath = New clsWApath
    stringPath.Reset FillModeWinding
    If stringPath.Append_String(0, 0, sText(LineIndex), m_Font.name, m_Font.Size, fontStyle) = False Then
        ' typical error: not TrueType font
        BuildDisplayPath = waErrUnsupportedFont
        Exit Function
    End If
    
    ' call function to measure each path, create lookup tables & return path point data
    m_Guides.CachePathPoints = True             ' allows class to process path faster
    pathLenTop = m_Guides.Length(1)
    If pathLenTop = 0! Then
        m_Guides.CachePathPoints = False        ' abort; something is wrong
        BuildDisplayPath = waErrInvalidGuide_Top ' path has not points
        Exit Function
    End If
    pathLenBtm = m_Guides.Length(2)
    If pathLenBtm = 0! Then                     ' abort; something is wrong
        m_Guides.CachePathPoints = False        ' path has no points
        BuildDisplayPath = waErrInvalidGuide_Btm
        Exit Function
    End If
    
    ' create or reset the source path
    If Not m_Source Is Nothing Then
        m_Source.Reset
    Else
        Set m_Source = New clsWApath            ' should not be Nothing; but just in case
    End If
    m_Source.Angle = m_Guides.Angle             ' carry over properties from the Guides
    m_Source.ScaleX = m_Guides.ScaleX
    m_Source.ScaleY = m_Guides.ScaleY
    
    ' =================== MAIN WORK STARTS HERE ===================
    ' preprocess the lines, calculating widths, scales, and offsets
    ' =============================================================
    
    ' unfortunately, we will loop thru the text 2 times
    ' 1st time: get metrics of each line of text & some overall measurements
    ' 2rd time: calc scales & place the characters of each line within the guide
    
    ' Step 1: Get line metrics
    ' ---------------------------------------------------------
    ReDim lineData(0 To LineCount - 1)
    
    For LineIndex = LineCount - 1 To 0 Step -1
        If LineIndex < LineCount - 1 Then       ' 1st line already done as initial testing above
            stringPath.Reset FillModeWinding
            stringPath.Append_String 0!, 0!, sText(LineIndex), m_Font.name, m_Font.Size, fontStyle
        End If
        stringPath.CachePathPoints = True       ' allows class to process a bit faster
        
        With lineData(LineIndex)
            ' get overall dimensions of the string path
            stringPath.GetBoundingBox False, fRect.X, fRect.Y, fRect.mWidth, fRect.mHeight
            If fRect.mHeight Then
                ' non-blank line; calc total width needed for the path
                .Length = fRect.mWidth + fRect.X + m_IndentL + m_IndentR + m_CharSpacing * stringPath.MarkerCount * 2&
                .yOffset = fRect.mHeight    ' flag used for comparison in loop below
                For mIndex = 1 To stringPath.MarkerCount
                    ' get the bounds of the next character
                    stringPath.GetBoundingBox False, gRect.X, gRect.Y, gRect.mWidth, gRect.mHeight, mIndex
                    ' keep track of tallest character in the string
                    If gRect.mHeight + gRect.Y > .blockHeight Then .blockHeight = gRect.mHeight + gRect.Y
                    ' keep track of the smallest difference from top of character to top of block
                    If gRect.Y < .yOffset Then .yOffset = gRect.Y
                    If mIndex = 1 Then          ' test for RTL text
                        If gRect.X > fRect.X + gRect.mWidth Then isRTL = True
                    End If
                Next
                ' keep track of the largest line within the the total lines of text
                If .blockHeight > ttlCtrlHt Then ttlCtrlHt = .blockHeight
            End If
        End With
    Next
    
    ttlHeight = ttlCtrlHt * LineCount + (LineCount - 1) * m_LineSpacing ' total size of lines needed
    ReDim charSrcPT(0 To 3)                                 ' warping source points
    ReDim charDstPT(0 To 3)                                 ' warping destination points
    Set charPath = New clsWApath
    
    For LineIndex = 0 To LineCount - 1
        If Len(sText(LineIndex)) Then
            If LineIndex Then
                stringPath.Reset FillModeWinding
                stringPath.Append_String 0, 0, sText(LineIndex), m_Font.name, m_Font.Size, fontStyle
                stringPath.CachePathPoints = True
                GdipGetPathWorldBounds stringPath.handle, gRect, 0&, 0&
            End If
            With lineData(LineIndex)
                ' Calculate character width scales for the top and bottom guides
                If (m_Alignment And waAlignStretchHorizontal) = 0& Then ' non-stretching perspective
                    ' see if we can display the full text with the current font setting
                    ' subtract indentation & compare against text length
                    If pathLenBtm >= .Length And pathLenTop >= .Length Then
                        ' ok, it will fit -- set scales
                        If pathLenBtm > pathLenTop Then
                            pathScaleTop = 1!
                            pathScaleBtm = pathLenBtm / pathLenTop
                        Else
                            pathScaleBtm = 1!
                            pathScaleTop = pathLenTop / pathLenBtm
                        End If
                    Else    ' won't fit; fallback to stretch mode
                        pathScaleTop = 0!
                    End If
                Else
                    pathScaleTop = 0!
                End If
                If pathScaleTop = 0! Then ' scale must be stretched to display text
                    pathScaleBtm = pathLenBtm / .Length  ' bottom guide scale
                    pathScaleTop = pathLenTop / .Length  ' top guide scale
                    If isRTL Then
                        MarginBtm = m_IndentR * pathScaleBtm
                        MarginTop = m_IndentR * pathScaleTop
                    Else
                        MarginBtm = m_IndentL * pathScaleBtm
                        MarginTop = m_IndentL * pathScaleTop
                    End If
                Else
                    Select Case (m_Alignment And &H7)
                    Case 0 ' left aligned, nothing special unless RTL text
                        If isRTL Then
                            MarginBtm = pathLenBtm - (.Length - m_IndentR) * pathScaleBtm
                            MarginTop = pathLenTop - (.Length - m_IndentR) * pathScaleTop
                        Else
                            MarginBtm = m_IndentL * pathScaleBtm
                            MarginTop = m_IndentL * pathScaleTop
                        End If
                    Case 1 ' right aligned
                        If isRTL Then
                            MarginBtm = m_IndentR * pathScaleBtm
                            MarginTop = m_IndentR * pathScaleTop
                        Else
                            MarginBtm = pathLenBtm - (.Length - m_IndentL) * pathScaleBtm
                            MarginTop = pathLenTop - (.Length - m_IndentL) * pathScaleTop
                        End If
                    Case 2 ' centered
                        MarginBtm = (pathLenBtm - (.Length - m_IndentL - m_IndentR) * pathScaleBtm) / 2
                        MarginTop = (pathLenTop - (.Length - m_IndentL - m_IndentR) * pathScaleTop) / 2
                    End Select
                End If
            End With
            For mIndex = 1 To stringPath.MarkerCount
                lSubPath = stringPath.ClonePath(mIndex)
                charPath.Reset FillModeWinding
                charPath.Append_Path lSubPath, False
                GdipDeletePath lSubPath
                charPath.GetBoundingBox False, gRect.X, gRect.Y, gRect.mWidth, gRect.mHeight
                
                If gRect.mWidth Then
                    If isRTL Then ' RTL text with additional character spacing
                        charSpacingOffset = (stringPath.MarkerCount - mIndex) * m_CharSpacing * 2 + m_CharSpacing
                    ElseIf m_CharSpacing > 0& Then
                        charSpacingOffset = m_CharSpacing * mIndex * 2 - m_CharSpacing
                    End If
                    
                    NextXbtm = MarginBtm + (gRect.X + charSpacingOffset) * pathScaleBtm
                    NextXtop = MarginTop + (gRect.X + charSpacingOffset) * pathScaleTop
                
                    ' get character left coords in relation to the guide path
                    m_Guides.NavigatePath NextXtop, charDstPT(0).X, charDstPT(0).Y, 1
                    m_Guides.NavigatePath NextXbtm, charDstPT(2).X, charDstPT(2).Y, 2
                                
                    ' calculate character right edge in scale
                    NextXbtm = NextXbtm + gRect.mWidth * pathScaleBtm ' calc end of character
                    NextXtop = NextXtop + gRect.mWidth * pathScaleTop
                                
                    ' get character right edge in relation to the guide path
                    m_Guides.NavigatePath NextXtop, charDstPT(1).X, charDstPT(1).Y, 1
                    m_Guides.NavigatePath NextXbtm, charDstPT(3).X, charDstPT(3).Y, 2
                                
                    charSrcPT(0).X = gRect.X:                charSrcPT(0).Y = gRect.Y
                    charSrcPT(1).X = gRect.mWidth + gRect.X: charSrcPT(1).Y = charSrcPT(0).Y
                    charSrcPT(2).X = charSrcPT(0).X:         charSrcPT(2).Y = charSrcPT(0).Y + gRect.mHeight
                    charSrcPT(3).X = charSrcPT(1).X:         charSrcPT(3).Y = charSrcPT(2).Y
                            
                    ' calculate warp relative coordinates to form a new baseline for the font
                    ' The character requires warping/skewing to fit in the shape.
                    ' The calcs below help scale & position the glyphs in relation to the raw glyph & the destination warp/skew
                    If (m_Alignment And waAlignAllCharsSameSizeVertical) = 0& Then
                        ' PROPORTIONAL VERTICAL POSITIONING
                        WarpOffsetBtm = ((LineIndex + 1) * ttlCtrlHt + (LineIndex * m_LineSpacing) - (ttlCtrlHt - (gRect.Y + gRect.mHeight))) / ttlHeight
                        WarpOffsetTop = (gRect.Y - lineData(LineIndex).yOffset + (LineIndex * (ttlCtrlHt + m_LineSpacing))) / ttlHeight
                    Else
                        ' SAME HEIGHT CHARACTERS
                        WarpOffsetBtm = ((LineIndex + 1) * ttlCtrlHt + (LineIndex * m_LineSpacing)) / ttlHeight
                        WarpOffsetTop = (LineIndex * (ttlCtrlHt + m_LineSpacing)) / ttlHeight
                    End If
                    
                    tmpVal1 = charDstPT(2).X - charDstPT(0).X
                    tmpVal2 = charDstPT(2).Y - charDstPT(0).Y
                    charDstPT(2).X = charDstPT(0).X + WarpOffsetBtm * tmpVal1
                    charDstPT(2).Y = charDstPT(0).Y + WarpOffsetBtm * tmpVal2
                    charDstPT(0).X = charDstPT(0).X + WarpOffsetTop * tmpVal1
                    charDstPT(0).Y = charDstPT(0).Y + WarpOffsetTop * tmpVal2
                    
                    tmpVal1 = charDstPT(3).X - charDstPT(1).X
                    tmpVal2 = charDstPT(3).Y - charDstPT(1).Y
                    charDstPT(3).X = charDstPT(1).X + WarpOffsetBtm * tmpVal1
                    charDstPT(3).Y = charDstPT(1).Y + WarpOffsetBtm * tmpVal2
                    charDstPT(1).X = charDstPT(1).X + WarpOffsetTop * tmpVal1
                    charDstPT(1).Y = charDstPT(1).Y + WarpOffsetTop * tmpVal2
                            
                    ' send the character off and warp it to the calculated coordinates
                    charPath.WarpPath_PTR VarPtr(charSrcPT(0)), VarPtr(charDstPT(0))
                    
                    ' add the character to the guide path
                    m_Source.Append_Path charPath.handle, False
                End If
            Next
        End If
    Next
    
    ' clean up GDI+ resource -- no longer needed
    m_Guides.CachePathPoints = False
    Set stringPath = Nothing
    Set charPath = Nothing
    
    ' update status flags
    m_Status = (waStat_Ok Or waStat_GuideCombined Or waStat_GuideBtm Or waStat_GuideTop)
    
    Me.Move m_AnchorPts.X, m_AnchorPts.Y
    
End Function

Private Function CreateBGRobject(Color As Long, ByVal Opacity As Long) As Boolean

    ' ---=====| Converts RGB to GDI+ required BGRA |====---
    
    If Opacity < 1 Then
        Exit Function
    ElseIf Opacity > 100 Then
        Opacity = 100
    End If
    Opacity = (255 * Opacity) \ 100
    Color = ((Color And &HFF) * &H10000) Or _
        (((Color \ &H100) And &HFF) * &H100) Or ((Color \ &H10000) And &HFF)
    If Opacity > 127 Then
        Color = Color Or (&H80000000 Or ((Opacity - 128) * &H1000000))
    Else
        Color = (Opacity * &H1000000) Or Color
    End If
    CreateBGRobject = True
    
End Function


Private Function CombineGuides() As Boolean

    ' ---=====| Combines two paths into one, marking each for easy reference |====---

    If (m_Status And waStat_GuideCombined) = waStat_GuideCombined Then
        CombineGuides = True    ' combined, nothing to do
    Else
        ' see if both top & bottom guides are created
        If Not (m_Status And (waStat_GuideBtm Or waStat_GuideTop)) = 0& Then
            
            If Not m_Tracker Is Nothing Then Call Tracking_End(True)
            m_Status = (m_Status Or waStat_GuideCombined)   ' update state
            
            ' clear any existing markers that may exist on the temp paths
            GdipClearPathMarkers m_guideBtm.handle
            GdipClearPathMarkers m_guideTop.handle
            
            ' add these paths to a single path, destroy the temp paths
            ' We add markers, 1st path will be marker #1 & second is marker #2
            m_Guides.Append_Path m_guideTop.handle, False, True
            Set m_guideTop = Nothing
            m_Guides.Append_Path m_guideBtm.handle, False, True
            Set m_guideBtm = Nothing
            
            CombineGuides = True
            
        End If
    End If
    
End Function

Private Function AppendGuide(TopGuide As Boolean, bPrepGuide As Boolean) As clsWApath

    ' ---=====| Ensures path being updated is ready, update flags/state as needed |====---

    If bPrepGuide Then
        Dim pClone As Long
        
        If Not m_Tracker Is Nothing Then Call Tracking_End(True)
        
        If (m_Status And waStat_GuideCombined) Then ' combined guides already exist
            ' we will separate them while the guide paths are being modified
            ' and when done, they will be combined again
            Set m_guideBtm = New clsWApath          ' create new class instances
            Set m_guideTop = New clsWApath
            pClone = m_Guides.ClonePath(1)          ' get top guide from combined path
            m_guideTop.Append_Path pClone, False    ' add it to temporary class
            GdipDeletePath pClone                   ' delete clone
            pClone = m_Guides.ClonePath(2)          ' do the same for bottom guide
            m_guideBtm.Append_Path pClone, False
            GdipDeletePath pClone
            m_Guides.Reset FillModeWinding
            
        End If
        ' create and/or return the correct temp guide path
        If TopGuide Then
            If m_guideTop Is Nothing Then Set m_guideTop = New clsWApath
            Set AppendGuide = m_guideTop
        Else
            If m_guideBtm Is Nothing Then Set m_guideBtm = New clsWApath
            Set AppendGuide = m_guideBtm
        End If
        m_Status = ((m_Status And Not waStat_GuideCombined) Or waStat_RebuildPath)
    
    Else
        ' set flags after path successfully appended/modified
        If TopGuide Then
            m_Status = (m_Status Or waStat_GuideTop)
        Else
            m_Status = (m_Status Or waStat_GuideBtm)
        End If
        m_Status = (m_Status Or waStat_DoFillBrush Or waStat_DoShadowBrush)
    
    End If

End Function

Private Sub AlignDisplayToGuide()

    ' keep the guides & display paths aligned to each other
    
    Dim X As Single, Y As Single, Cx As Single, Cy As Single
    
    With m_Guides
        ' get the tightest bounding rectangle around the anchor path
        .GetBoundingBox True, X, Y, Cx, Cy
        
        ' the nonAnchor path will have its scale & rotation position set by
        ' using the following 6 coords (2 pts for rotation & 1 pt for scaling).
        ' This overrides the path class using default calculations to position
        ' the path when scaling/rotating and "locks" the nonAnchor path with
        ' the Anchor path
        m_Source.OverrideWorldTransform Cx / 2 + X, Cy / 2 + Y, _
            m_AnchorPts.X - X - Cx * (1 - .ScaleX) / 2, _
            m_AnchorPts.Y - Y - Cy * (1 - .ScaleY) / 2, _
            m_AnchorPts.X - X + X * (1 - .ScaleX), _
            m_AnchorPts.Y - Y + Y * (1 - .ScaleY)
    End With

End Sub

Public Function Tracking_Begin(ByVal Key As String, ByVal hostContainer As Object, CallbackImplementation As ITrackingCallback, ByVal currentX As Single, ByVal currentY As Single)
    
    If CallbackImplementation Is Nothing Then Exit Function
    If hostContainer Is Nothing Then Exit Function
    
    Dim pClone As Long, tTracker As clsWATracker
    Dim srcLeft As Single, srcTop As Single, OffsetX As Single, OffsetY As Single
    Dim gudLeft As Single, gudTop As Single, gudClipLeft As Single, gudClipTop As Single
    
    
    m_TrackingKey = Key
    Set m_BackupPath = Nothing
    pClone = m_Guides.ClonePath
    Set m_BackupPath = New clsWApath
    m_BackupPath.Append_Path pClone, False
    GdipDeletePath pClone
    
    With m_Guides
        m_BackupPath.Angle = .Angle
        m_BackupPath.ScaleX = .ScaleX
        m_BackupPath.ScaleY = .ScaleY
        m_BackupPath.Move m_AnchorPts.X, m_AnchorPts.Y
    End With
    
    Set client_Tracker = CallbackImplementation
    m_BackupPath.CachePathPoints = True
    
    Set tTracker = New clsWATracker
    Set m_Tracker = tTracker.CreateTracker(hostContainer, m_Tracker)
    If Not m_Tracker Is Nothing Then
        If m_Tracker.InitializeTracking(currentX, currentY, m_BackupPath) = False Then
            Set m_BackupPath = Nothing
            Set m_Tracker = Nothing
            Set client_Tracker = Nothing
        End If
    End If
End Function

Public Sub Tracking_End(ByVal Cancel As Boolean)
    If Not m_Tracker Is Nothing Then
        Call m_Tracker.TerminateTracking(Cancel)
    End If
End Sub

' WITH EVENTS CALLBACKS
Private Sub m_Tracker_GetBoundingPoints(Points() As Single)
    ReDim Points(0 To 1, 0 To 3)
    m_BackupPath.GetBoundingPoints_PTR True, True, VarPtr(Points(0, 0))
End Sub

Private Sub m_Tracker_GetDimensions(Left As Single, Top As Single, Width As Single, Height As Single)
    m_BackupPath.GetBoundingBox True, 0!, 0!, Width, Height
    Left = m_BackupPath.DisplayLeft
    Top = m_BackupPath.DisplayTop
End Sub

Private Sub m_Tracker_MoveObject(X As Single, Y As Single)
    m_BackupPath.Move X, Y
End Sub

Private Sub m_Tracker_RenderObject(ByVal hDC As Long, ByVal bErase As Boolean)
    m_BackupPath.RenderOutline_XOR hDC, vbBlue, , psXOR_Dotted
End Sub

Private Sub m_Tracker_Rotation(Angle As Single, bSet As Boolean)
    If bSet Then m_BackupPath.Angle = Angle Else Angle = m_BackupPath.Angle
End Sub

Private Sub m_Tracker_ScaleRatios(ScaledX As Single, ScaledY As Single, ByVal bSet As Boolean)
    If bSet Then
        m_BackupPath.ScaleX = ScaledX
        m_BackupPath.ScaleY = ScaledY
    Else
        ScaledX = m_BackupPath.ScaleX
        ScaledY = m_BackupPath.ScaleY
    End If
End Sub

' callback events from the clsWATracker class
Private Sub m_Tracker_TrackingChanged(X As Single, Y As Single, Shift As Integer, Mode As eTrackingModes, Cancel As Boolean)
    client_Tracker.TrackingPointChanged m_TrackingKey, X, Y, Shift, Mode, Cancel
End Sub

Private Sub m_Tracker_TrackingStarted(X As Single, Y As Single, Mode As eTrackingModes)
    client_Tracker.TrackingStarted m_TrackingKey, X, Y, Mode
End Sub

Private Sub m_Tracker_TrackingTerminated(Cancel As Boolean)

    If Cancel = False Then
        With m_BackupPath
            m_Guides.Angle = .Angle
            m_Guides.ScaleX = .ScaleX
            m_Guides.ScaleY = .ScaleY
            
            m_Source.ScaleX = .ScaleX
            m_Source.ScaleY = .ScaleY
            m_Source.Angle = .Angle
            m_AnchorPts.X = .DisplayLeft
            m_AnchorPts.Y = .DisplayTop
        End With
        Me.Move m_AnchorPts.X, m_AnchorPts.Y
    Else
        m_Tracker.TerminateTracking False
        Set m_Tracker = Nothing
        Set m_BackupPath = Nothing
    End If
    If Not client_Tracker Is Nothing Then client_Tracker.TrackingTerminated m_TrackingKey, Cancel
    If Cancel Then Set client_Tracker = Nothing
    
End Sub

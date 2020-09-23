VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDIbrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is still a work in progress

' This class wraps many of the common GDI+ brush/pen calls.
' The Texture brush is basically GDI+ version of GDI's Pattern brush. This class
' requires the texture to be a bitmap within a stdPicture object. Updates will
' allow the texture to be imported from resource file, disk file and possibly
' others. However, it is important that while the brush exists, the image that
' the texture brush uses does not get destroyed -- else crash potential.  Therefore,
' the requirement to pass it as a stdPicture should help prevent accidentally
' releasing the image before the texture brush is destroyed.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColors As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (ByRef pRect As RECTF, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Mode As Long, ByVal WrapMode As Long, ByRef lineGradient As Long) As Long
Private Declare Function GdipSetLineColors Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColor1 As Long, ByVal mColor2 As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipGetPathWorldBounds Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mBounds As RECTF, ByVal mMatrix As Long, ByVal mPen As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashStyle As Long) As Long
Private Declare Function GdipSetPenDashCap197819 Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashCap As Long) As Long
Private Declare Function GdipCreateHatchBrush Lib "GdiPlus.dll" (ByVal mHatchStyle As Long, ByVal mForecol As Long, ByVal mBackcol As Long, ByRef mBrush As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipCreateTexture Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mWrapMode As Long, ByRef mTexture As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal mImage As Long) As Long

Public Enum eTextureWrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum


Public Enum eHatchStyles
    HatchStyleHorizontal = &H0
    HatchStyleVertical = &H1
    HatchStyleForwardDiagonal = &H2
    HatchStyleBackwardDiagonal = &H3
    HatchStyleCross = &H4
    HatchStyleDiagonalCross = &H5
    HatchStyle05Percent = &H6
    HatchStyle10Percent = &H7
    HatchStyle20Percent = &H8
    HatchStyle25Percent = &H9
    HatchStyle30Percent = &HA
    HatchStyle40Percent = &HB
    HatchStyle50Percent = &HC
    HatchStyle60Percent = &HD
    HatchStyle70Percent = &HE
    HatchStyle75Percent = &HF
    HatchStyle80Percent = &H10
    HatchStyle90Percent = &H11
    HatchStyleLightDownwardDiagonal = &H12
    HatchStyleLightUpwardDiagonal = &H13
    HatchStyleDarkDownwardDiagonal = &H14
    HatchStyleDarkUpwardDiagonal = &H15
    HatchStyleWideDownwardDiagonal = &H16
    HatchStyleWideUpwardDiagonal = &H17
    HatchStyleLightVertical = &H18
    HatchStyleLightHorizontal = &H19
    HatchStyleNarrowVertical = &H1A
    HatchStyleNarrowHorizontal = &H1B
    HatchStyleDarkVertical = &H1C
    HatchStyleDarkHorizontal = &H1D
    HatchStyleDashedDownwardDiagonal = &H1E
    HatchStyleDashedUpwardDiagonal = &H1F
    HatchStyleDashedHorizontal = &H20
    HatchStyleDashedVertical = &H21
    HatchStyleSmallConfetti = &H22
    HatchStyleLargeConfetti = &H23
    HatchStyleZigZag = &H24
    HatchStyleWave = &H25
    HatchStyleDiagonalBrick = &H26
    HatchStyleHorizontalBrick = &H27
    HatchStyleWeave = &H28
    HatchStylePlaid = &H29
    HatchStyleDivot = &H2A
    HatchStyleDottedGrid = &H2B
    HatchStyleDottedDiamond = &H2C
    HatchStyleShingle = &H2D
    HatchStyleTrellis = &H2E
    HatchStyleSphere = &H2F
    HatchStyleSmallGrid = &H30
    HatchStyleSmallCheckerBoard = &H31
    HatchStyleLargeCheckerBoard = &H32
    HatchStyleOutlinedDiamond = &H33
    HatchStyleSolidDiamond = &H34
End Enum
Public Enum eDashCap
    DashCapFlat = &H0
    DashCapRound = &H2
    DashCapTriangle = &H3
End Enum
Public Enum ePenStyle
    DashStyleSolid = &H0
    DashStyleDash = &H1
    DashStyleDot = &H2
    DashStyleDashDot = &H3
    DashStyleDashDotDot = &H4
End Enum
Public Enum eGradientBrushStyles ' gradient brush options
    LinearHorizontal = 0
    LinearVertical = 1
    LinearForwardDiagonal = 2
    LinearBackwardDiagonal = 3
    PathRadialCenterOut = 4         ' << custom, not GDI+ standard
End Enum


Private Const UnitPixel = &H2
Private Type RECTF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private m_Brush As Long             ' brush handle
Private m_Pen As Long               ' pen handle
Private m_FillColor As Long         ' brush primary color
Private m_Gradient As Long          ' brush gradient color
Private m_Style As eGradientBrushStyles ' <8 gradients, 8 solid brush, 16+ hatch, 128+ texture
Private m_PenStyle As ePenStyle     ' pen style
Private m_PenCap As eDashCap        ' pen cap style
Private m_OffsetX As Single         ' brush/pen offset horizontal
Private m_OffsetY As Single         ' brush/pen offset vertical
Private m_OutlineColor As Long      ' pen color
Private m_OutlineWidth As Long      ' pen width
Private m_UpdateRqd As Boolean      ' flag indicating brush bounds need to be updated
Private m_Texture As StdPicture
Private m_Image As Long

' ---=====| Brush Handle - Read Only |====---
Public Property Get FillHandle() As Long
Attribute FillHandle.VB_MemberFlags = "200"
    FillHandle = m_Brush
End Property
' ---=====| Pen Handle - Read Only |====---
Public Property Get OutlineHandle() As Long
    OutlineHandle = m_Pen
End Property

' ---=====| Update Flag |====---
Public Property Get UpdateGradientBounds() As Boolean
    ' Note: Designed for use with the Path class, can be used without it.
    ' Whenever a new path is created, call this property and set it to True.
    ' When the gradient style changes and is in conflict with the type
    ' style previously set, this flag is automatically changed.
    ' You should query this property before rendering the path and if
    ' it returns True, call AssignGradientBrush routine to update the brush
    
    ' Jump to end of routine for recommended usage with a GDI+ rendering routine
    UpdateGradientBounds = m_UpdateRqd
End Property
Public Property Let UpdateGradientBounds(Update As Boolean)
    If m_Style < 8 Then m_UpdateRqd = Update
End Property

' ---=====| Offset Options |====---
Public Property Get OffsetX() As Single
    OffsetX = m_OffsetX
End Property
Public Property Let OffsetX(ByVal Offset As Single)
    m_OffsetX = Offset
End Property
Public Property Get OffsetY() As Single
    OffsetY = m_OffsetY
End Property
Public Property Let OffsetY(ByVal Offset As Single)
    m_OffsetY = Offset
End Property

' ---=====| Set Offsets in one call |====---
Public Sub SetOffsets(ByVal OffsetX As Single, ByVal OffsetY As Single)
    m_OffsetX = OffsetX
    m_OffsetY = OffsetY
End Sub

' ---=====| Brush Destruction |====---
Public Sub RemoveFill()
    If m_Brush Then
        GdipDeleteBrush m_Brush
        m_Brush = 0&
        m_FillColor = 0&            ' flag used to indicate do not re-create brush
        m_UpdateRqd = False
    End If
End Sub

' ---=====| Pen Destruction |====---
Public Sub RemoveOutline()
    If m_Pen Then
        GdipDeletePen m_Pen
        m_Pen = 0&
    End If
End Sub

' ---=====| Brush Updating |====---
' This should be called whenever the UpdateGradientBounds flag is True
Public Function AssignGradientBrush(ByVal thePath As Long) As Boolean
                                    
    If Not thePath = 0& Then
        If Not m_FillColor = 0& Then
        
            If m_Style >= 8 Then
                m_UpdateRqd = False
            Else
                Dim pRect As RECTF, nrColors As Long
                If m_Brush Then
                    GdipDeleteBrush m_Brush
                    m_Brush = 0&
                End If
                If m_Gradient = 0& Then m_Gradient = m_FillColor
                
                If m_Style = PathRadialCenterOut Then
                
                    GdipCreatePathGradientFromPath thePath, m_Brush
                    If Not m_Brush = 0& Then
                        nrColors = 1&
                        GdipSetPathGradientSurroundColorsWithCount m_Brush, m_Gradient, nrColors
                        GdipSetPathGradientCenterColor m_Brush, m_FillColor
                        m_UpdateRqd = False
                        AssignGradientBrush = True
                    End If
                    
                Else
                
                    GdipGetPathWorldBounds thePath, pRect, 0&, 0&
                    GdipCreateLineBrushFromRect pRect, m_FillColor, m_Gradient, 0&, 0&, m_Brush
                    If Not m_Brush = 0& Then
                        m_UpdateRqd = False
                        AssignGradientBrush = True
                    End If
                    
                End If
            
            End If
        End If
    End If
    
End Function

' ---=====| Pen Creation |====---
Public Sub SetOutlineAttributes(ByVal Color As Long, Optional ByVal Opacity As Long = 100&, _
                                Optional ByVal Width As Long = 1&, _
                                Optional ByVal Style As ePenStyle = DashStyleSolid, _
                                Optional ByVal Cap As eDashCap = DashCapFlat, _
                                Optional ByVal ColorBGRA As Variant)
    
    ' Note: Opacity values are between 0 and 100
    ' If the optional BGRA color provided, it must be in GDI+ format & will be used
    Me.RemoveOutline
    If Width < 1 Then Width = 1
    m_OutlineWidth = Width
    If Style < DashStyleSolid Or Style > DashStyleDashDotDot Then Style = DashStyleSolid
    m_PenStyle = Style
    If Cap < DashCapFlat Or Cap > DashCapTriangle Then Cap = DashCapFlat
    m_PenCap = Cap
    If IsMissing(ColorBGRA) Then
        GdipCreatePen1 ConvertColor(Color, Opacity, m_OutlineColor), m_OutlineWidth, UnitPixel, m_Pen
    Else
        m_OutlineColor = CLng(Val(ColorBGRA))
        GdipCreatePen1 m_OutlineColor, m_OutlineWidth, UnitPixel, m_Pen
    End If
    If m_Pen Then
        GdipSetPenDashStyle m_Pen, m_PenStyle
        GdipSetPenDashCap197819 m_Pen, m_PenCap
    End If
    
End Sub

' ---=====| Pen Attribute Retrieval |====---
Public Sub GetOutlineAttributes(Color As Long, Optional Opacity As Long, _
                            Optional Width As Long, _
                            Optional Style As ePenStyle, _
                            Optional Cap As eDashCap, _
                            Optional ColorBGRA As Long)

    ' Note: Opacity values are between 0 and 100
    ConvertGDIpColor Color, Opacity, m_OutlineColor
    Width = m_OutlineWidth
    Style = m_PenStyle
    Cap = m_PenCap
    ColorBGRA = m_OutlineColor

End Sub


' ---=====| Gradient Style Options |====---
Public Property Let GradientStyle(ByVal newStyle As eGradientBrushStyles)
    
    If newStyle >= LinearHorizontal And newStyle <= PathRadialCenterOut Then
        RemoveTextureBrush
        If newStyle <> m_Style Then
            If newStyle = PathRadialCenterOut Or m_Style = PathRadialCenterOut Then
                ' changing to/from path gradient brush, need path data or bounds
                m_UpdateRqd = True
                m_Style = newStyle
            ElseIf m_Style >= 8 Then ' was solid fill/hatch, need rect bounds
                m_UpdateRqd = True
                m_Style = newStyle
            Else    ' no update, simply change gradient direction
                m_Style = newStyle
                RefreshGradientBrush
            End If
        End If
    End If
    
End Property
Public Property Get GradientStyle() As eGradientBrushStyles
    GradientStyle = m_Style
End Property

' ---=====| Gradient Colors |====---
Public Sub SetGradientAttributes(ByVal PrimaryColor As Long, ByVal GradientColor As Long, _
                                Optional ByVal PrimaryOpacity As Long = 100&, Optional ByVal GradientOpacity As Long = 100&, _
                                Optional ByVal PrimaryColorBGRA As Variant, _
                                Optional ByVal GradientColorBGRA As Variant)

    ' Note: Opacity values are between 0 and 100
    ' If the optional BGRA colors provided, they must be in GDI+ format & will be used
    If IsMissing(PrimaryColorBGRA) Then
        ConvertColor PrimaryColor, PrimaryOpacity, m_FillColor
    Else
        m_FillColor = CLng(Val(PrimaryColorBGRA))
    End If
    If IsMissing(GradientColorBGRA) Then
        ConvertColor GradientColor, GradientOpacity, m_Gradient
    Else
        m_Gradient = CLng(Val(GradientColorBGRA))
    End If
    ' recreate the brush with changed colors as needed
    If m_UpdateRqd = False Then RefreshGradientBrush
        
End Sub

' ---=====| Gradient Colors Retrieval |====---
Public Sub GetGradientAttributes(PrimaryColor As Long, GradientColor As Long, _
                                Optional PrimaryOpacity As Long, Optional GradientOpacity As Long, _
                                Optional PrimaryColorBGRA As Long, Optional GradientColorBGRA As Long)

    ' Note: Opacity values are between 0 and 100
    If m_Style < 8 Then
        ConvertGDIpColor PrimaryColor, PrimaryOpacity, m_FillColor
        ConvertGDIpColor GradientColor, GradientOpacity, m_Gradient
        PrimaryColorBGRA = m_FillColor
        GradientColorBGRA = m_Gradient
    End If
    
End Sub

' ---=====| Solid Brush Color |====---
Public Sub SetSolidFillAttributes(ByVal Color As Long, Optional ByVal Opacity As Long = 100&, _
                                  Optional ByVal ColorBGRA As Variant)

    ' Note: Opacity values are between 0 and 100
    ' If the optional BGRA color provided, it must be in GDI+ format & will be used
    Me.RemoveFill
    RemoveTextureBrush
    m_Style = 8 ' solid
    m_UpdateRqd = False
    If IsMissing(ColorBGRA) Then
        GdipCreateSolidFill ConvertColor(Color, Opacity, m_FillColor), m_Brush
    Else
        m_FillColor = CLng(Val(ColorBGRA))
        GdipCreateSolidFill m_FillColor, m_Brush
    End If
    
End Sub

' ---=====| Solid Brush Color Retrieval |====---
Public Sub GetSolidFillAttributes(Color As Long, Optional Opacity As Long, Optional ColorBGRA As Long)

    ' Note: Opacity values are between 0 and 100
    If m_Style = 8 Then
        ConvertGDIpColor Color, Opacity, m_FillColor
        ColorBGRA = m_FillColor
    End If
End Sub

' ---=====| Hatch Brush Attributes |====---
Public Sub SetHatchBrushAttributes(ByVal ForeColor As Long, Optional ByVal BackColor As Long, _
                                    Optional ByVal Style As eHatchStyles = HatchStyleHorizontal, _
                                    Optional ByVal ForeOpacity As Long = 100&, _
                                    Optional ByVal BackOpacity As Long = 100&, _
                                    Optional ByVal ForeColorBGRA As Variant, _
                                    Optional ByVal BackColorBGRA As Variant)

    ' Note: Opacity values are between 0 and 100
    ' If the optional BGRA colors provided, they must be in GDI+ format & will be used
    
    Me.RemoveFill
    RemoveTextureBrush
    If Style < HatchStyleHorizontal Or Style > HatchStyleSolidDiamond Then Style = HatchStyleHorizontal
    m_Style = 16 + Style ' hatch = 16+
    m_UpdateRqd = False
    If IsMissing(BackColorBGRA) Then
        ConvertColor BackColor, BackOpacity, m_FillColor
    Else
        m_FillColor = CLng(Val(BackColorBGRA))
    End If
    If IsMissing(ForeColorBGRA) Then
        ConvertColor ForeColor, ForeOpacity, m_Gradient
    Else
        m_Gradient = CLng(Val(ForeColorBGRA))
    End If
    GdipCreateHatchBrush Style, m_Gradient, m_FillColor, m_Brush
    
End Sub

' ---=====| Hatch Brush Attributes Retrieval |====---
Public Sub GetHatchBrushAttributes(ForeColor As Long, ByVal BackColor As Long, _
                                    Style As eHatchStyles, _
                                    Optional ForeOpacity As Long, Optional BackOpacity As Long, _
                                    Optional ForeColorBGRA As Long, _
                                    Optional BackColorBGRA As Long)

    ' Note: Opacity values are between 0 and 100
    If m_Style >= 16 And m_Style < 128 Then
        ConvertGDIpColor ForeColor, ForeOpacity, m_Gradient
        ConvertGDIpColor BackColor, BackOpacity, m_FillColor
        Style = m_Style - 16
        BackColorBGRA = m_FillColor
        ForeColorBGRA = m_Gradient
    End If
    
End Sub

' ---=====| Hatch Brush Attributes Retrieval |====---
Public Sub SetTextureBrush(Texture As StdPicture, Optional ByVal WrapMode As eTextureWrapMode = WrapModeTile)

    If m_Brush Then Me.RemoveFill
    If Not Texture Is Nothing Then
        If Texture.Type = vbPicTypeBitmap Then
            Set m_Texture = Texture
            If m_Image Then
                GdipDisposeImage m_Image
                m_Image = 0&
            End If
            GdipCreateBitmapFromHBITMAP Texture.Handle, 0&, m_Image
            If m_Image Then
                GdipCreateTexture m_Image, WrapMode, m_Brush
            End If
        End If
    End If
    m_Style = 128 + WrapMode

End Sub

' ---=====| Texture Brush Attributes Retrieval |====---
Public Sub GetTextureBrush(Texture As StdPicture, WrapMode As eTextureWrapMode)
    
    If m_Style >= 128 Then
        Set Texture = m_Texture
        WrapMode = m_Style - 128
    End If
    
End Sub

' ---=====| Internal Use Only |====---
' Updates gradient brush colors as needed
Private Sub RefreshGradientBrush()

    Dim nrColors As Long
    If m_Brush Then
        If m_Brush Then
            GdipDeleteBrush m_Brush
            m_Brush = 0&
        End If
        If m_Style = PathRadialCenterOut Then
            nrColors = 1&
            GdipSetPathGradientSurroundColorsWithCount m_Brush, m_FillColor, nrColors
            GdipSetPathGradientCenterColor m_Brush, m_Gradient
        Else
            GdipSetLineColors m_Brush, m_FillColor, m_Gradient
        End If
    End If

End Sub

' ---=====| Destroys texture brush GDI+ image object |====---
Private Sub RemoveTextureBrush()
    
    If m_Image Then
        GdipDisposeImage m_Image
        m_Image = 0&
    End If
    If Not m_Texture Is Nothing Then Set m_Texture = Nothing
    
End Sub

' ---=====| Convert RGB Colors to GDI+ Color Format |====---
Private Function ConvertColor(Color As Long, Opacity As Long, propVal As Long) As Long

    Dim BGRA(0 To 3) As Byte
    If Opacity < -100 Then
        Opacity = 100
    ElseIf Opacity > 100 Then
        Opacity = 100
    End If
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory propVal, BGRA(0), 4&
    ConvertColor = propVal

End Function

' ---=====| Convert GDI+ Color Format to RGB Color + Alpha |====---
Private Sub ConvertGDIpColor(Color As Long, Opacity As Long, propVal As Long)

    Dim BGRA(0 To 3) As Byte
    CopyMemory BGRA(0), propVal, 4&
    Opacity = (BGRA(3) * 100&) \ 255
    Color = ((BGRA(0) * &H10000) Or (BGRA(1) * &H100&) Or BGRA(2))

End Sub

' ---=====| Initial Property Values |====---
Private Sub Class_Initialize()
    m_Style = 8    ' default is solid brush style
End Sub
' ---=====| Clean Up |====---
Private Sub Class_Terminate()
    
    Me.RemoveFill
    Me.RemoveOutline
    If m_Image Then GdipDisposeImage m_Image

End Sub



' **************************************************************
'   EXAMPLE OF USING THIS CLASS IN A GDI+ RENDERING ROUTINE
' **************************************************************

'    m_Brushes() is an array of this class. Multiple brushes can be
'    used in the same rendering; and if done, you can apply them
'    by enumerating the array. Don't allow uninitialized classes in the
'    array, or if you do, test to ensure: Not m_Brushes(br) Is Nothing

'    ' reassign brush boundaries/path as needed & then render
'    Dim br As Long, shMx As Long
'    For br = 0 To nrBrushes - 1
'        With m_Brushes(br)
'            ' update the path bounds if needed
'            If .UpdateGradientBounds Then .AssignGradientBrush [yourPathHandle]
'            If Not (.FillHandle = 0& And .OutlineHandle = 0&) Then
'                ' if the brush/pen is to be offset, do that now
'                If Not (.OffsetX = 0! And .OffsetY = 0!) Then
'                    GdipCreateMatrix shMx
'                    GdipTranslateMatrix shMx, .OffsetX, .OffsetY, MatrixOrderPrepend
'                    GdipTransformPath [yourPathHandle], shMx
'                End If
'
'                ' now fill and/or outline the path
'                If Not .FillHandle = 0& Then GdipFillPath hGraphics, .FillHandle, [yourPathHandle]
'                If Not .OutlineHandle = 0& Then GdipDrawPath hGraphics, .OutlineHandle, [yourPathHandle]
'
'                ' reset the path coordinates if shifted above
'                If Not shMx = 0& Then
'                    GdipTranslateMatrix shMx, .OffsetX * -2, .OffsetY * -2, MatrixOrderPrepend
'                    GdipTransformPath [yourPathHandle], shMx
'                    GdipDeleteMatrix shMx
'                    shMx = 0&
'                End If
'
'            End If
'        End With
'    Next


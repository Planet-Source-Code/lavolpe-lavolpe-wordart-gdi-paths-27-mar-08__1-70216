VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWATracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Something I am trying to formalize into a generic object tracker.
' Ideally, this will eventually be able to track anything: pictures, labels, etc.
' But for this project, it works strictly with public events contained within
' the path class. Still in the workings. Needs more work, but usable for this project

' THIS CLASS IS NEVER EXPOSED TO THE USER. IT SHOULD BE CALLED ONLY BY THE
' OBJECT INTERFACE (CLASS/FORM/whatever) THAT IS CONTROLLING OBJECTS TO BE TRACKED

' To use this properly
' In the controlling object, declare as :: Private WithEvents myTracker As clsWATracker
' To initialize tracking, from the controlling object:
'       Dim tmpTracker As New clsWATracker
'       Set myTracker = tmpTracker.CreateTracker(hostObjectUsedForDrawing, myTracker)
'       If myTracker Is Nothing Then ... error occurred; invalid HostObjectForDrawing
' To begin tracking, from the controlling object: InitializeTracking X,Y, drawingObject
' To terminate tracking & remove handles, from controlling object: TerminateTracking
' Now get feedback from the tracker
'  :: Respond to the Public Events
'   -- TrackingStarted triggered when InitializeTracking called or when a tracking handle is clicked
'   -- TrackingChanged triggered whenever handle is dragged
'   -- TrackingTerminated triggered when mouse is released or ESC pressed
'       :: can also be triggered when this class' TerminateTracking is called & optional parameter is True
' To send feedback to the user, ensure they pass you an instance of ITrackingCallback, then...
'   relay to the user the appropriate ITrackingCallback event related to the Public event you received

' In the test project provided, look at cWordArt's Tracking_Begin & Tracking_End functions to see
'   how the class initializes/terminates tracking
' Also look at how it relays the events to the frmWordArt when these Public events are triggered

' public events sent back to owning class
Public Event TrackingStarted(X As Single, Y As Single, Mode As eTrackingModes)
Public Event TrackingChanged(X As Single, Y As Single, Shift As Integer, Mode As Long, Cancel As Boolean)
Public Event TrackingTerminated(Cancel As Boolean)

' Public events used to get/set object information
Public Event RenderObject(ByVal hDC As Long, ByVal bErase As Boolean)
Public Event Rotation(Angle As Single, bSet As Boolean)
Public Event ScaleRatios(ScaledX As Single, ScaledY As Single, ByVal bSet As Boolean)
Public Event GetDimensions(Left As Single, Top As Single, Width As Single, Height As Single)
Public Event GetBoundingPoints(Points() As Single)
Public Event MoveObject(X As Single, Y As Single)


Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_LBUTTONDOWN As Long = &H201

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' GDI+ UDTs
Private Type RECTF
    X As Single
    Y As Single
    Width As Single
    Height As Single
End Type
Private Type POINTF
    X As Single
    Y As Single
End Type

Private m_XScaleRatio As Single     ' used to allow aspect ratio sizing
Private m_MousePos As POINTF        ' used to track last mouse coordinates
Private m_PointOrigin As POINTF     ' original mouse coords when tracking originated

Private m_Radian As Single          ' Angle converted to radian
Private m_ScaleBox As POINTF        ' used to track scale size changes
Private m_SrcBox As POINTF          ' scaled object dimensions when tracking begins
Private m_Angle As Single           ' object angle

Private PI As Single                ' used in a few routines; cached vs recalculating
Private m_IsTracking As Boolean     ' whether we are tracking or not
Private m_XorFlag As Boolean

Private theHandles(0 To 10) As clsTrackingHandle ' tracking handle classes
Private WithEvents theTimer As VB.Timer          ' currently not used
Attribute theTimer.VB_VarHelpID = -1
Private theContainer As Object      ' the container containing the drawing object (form, picbox)
Private m_ParentHwnd As Long        ' the handle to where creation of handles is cached

Private Const HANDLESIZE As Long = 7        ' tracking handle size - change as needed; keep odd vs even values
Private Const ANGLERESTRICTION As Single = 15! ' when restricting rotation angle, angle to restrict to

Implements ITrackingCallback        ' callback for the clsTrackingHandle class

Public Function CreateTracker(ByVal Host As Object, currentTracker As clsWATracker) As clsWATracker

    ' Function must be called before InitializeTracking method is called.
    ' If successful, it will return a clsWATracker object ready to be used to begin tracking
    ' If it fails the return value is Nothing

    If Not Host Is Nothing Then
        On Error Resume Next
        
        Dim p As Long, propVal As Long
        Dim tmpObject As Object
        Dim sName As String
        Dim newTracker As clsWATracker
        Dim bGetHandleRefs As Boolean
        
        ' find the form that contains the Host object
        m_ParentHwnd = 0&
        Set theContainer = Host
        Do
            If TypeOf Host Is Control Then
                Set Host = theContainer.Container
                If Err Then Exit Do ' can't think of what kind of error here; but safety so infinite loop isn't created
            Else
                Exit Do
            End If
        Loop
        If Err Then
            Set theContainer = Nothing
            Err.Clear
            Exit Function
        End If
        
        propVal = theContainer.hDC ' will trigger error if passed container has no hDC property
        m_ParentHwnd = Host.hwnd
        If Err Then             ' don't use in an object without exposed hDC property (i.e., frames)
            Err.Clear           ' can't use this in a windowless container usercontrol
            Set theContainer = Nothing
            m_ParentHwnd = 0&
            Exit Function
        End If
        
        propVal = GetProp(m_ParentHwnd, "waTracker.v1")
        If propVal = 0& Then            ' first time this function ever called
            Set currentTracker = Nothing
            Set newTracker = Me         ' or it was released & need to check if handles exist
            SetProp m_ParentHwnd, "waTracker.v1", ObjPtr(Me)
        
            propVal = GetProp(m_ParentHwnd, "waTracker.v1.Handle")
            If propVal = 0& Then                   ' tracking handles never created, do so now
                
                sName = "waHndl" & ObjPtr(Me) & ObjPtr(theContainer)
                For p = 0 To 10
                    Set theHandles(p) = New clsTrackingHandle
                    If TypeOf theContainer Is Control Then
                        If theHandles(p).CreateHandle(HANDLESIZE, theContainer.Parent, sName & p, p, Me, theContainer) = False Then Exit For
                    Else
                        If theHandles(p).CreateHandle(HANDLESIZE, theContainer.Parent, sName & p, p, Me, theContainer.Parent) = False Then Exit For
                    End If
                    If Err Then Exit For
                    If p Mod 2 Then ' odd handles vs even
                        theHandles(p).This.BackColor = vbCyan
                    Else
                        theHandles(p).This.BackColor = vbYellow
                    End If
                    Select Case p
                    Case 1: theHandles(p).This.ToolTipText = "Left"
                    Case 2: theHandles(p).This.ToolTipText = "T/L"
                    Case 3: theHandles(p).This.ToolTipText = "Top"
                    Case 4: theHandles(p).This.ToolTipText = "T/R"
                    Case 5: theHandles(p).This.ToolTipText = "Right"
                    Case 6: theHandles(p).This.ToolTipText = "B/R"
                    Case 7: theHandles(p).This.ToolTipText = "Bottom"
                    Case 8: theHandles(p).This.ToolTipText = "B/L"
                    Case 9, 10: theHandles(p).This.BackColor = vbRed
                            theHandles(p).This.ToolTipText = "Rotate"
                    End Select
                    If p = 0 Then
                        SetProp m_ParentHwnd, "waTracker.v1.Handle", ObjPtr(theHandles(p).This)
                    Else
                        SetProp theHandles(p - 1).This.hwnd, "waTracker.v1.Next", ObjPtr(theHandles(p).This)
                    End If
                Next
                If Err Then
                    Err.Clear ' failed to create handles
                    Set newTracker = Nothing
                End If
                
            Else        ' tracking handles exist, reuse them
                bGetHandleRefs = True
            End If
            
        ElseIf propVal <> ObjPtr(currentTracker) Then
        
            ' another class is currently tracking
            CopyMemory tmpObject, propVal, 4&
            Set newTracker = tmpObject
            CopyMemory tmpObject, 0&, 4&
            newTracker.TerminateTracking True ' send it a cancel event
            Set newTracker = Me
            bGetHandleRefs = True
            
        Else
            
            Set newTracker = currentTracker
            m_ParentHwnd = 0&
            
        End If
        
        If bGetHandleRefs Then
            Set theHandles(0) = New clsTrackingHandle
            CopyMemory tmpObject, propVal, 4&
            Set theHandles(0).This = tmpObject
            CopyMemory tmpObject, 0&, 4&
            Set theHandles(0).CallBackObject = Me
            theHandles(0).Mode = waMoving
            theHandles(0).This.ZOrder
            
            For p = 1 To 10
                propVal = GetProp(theHandles(p - 1).This.hwnd, "waTracker.v1.Next")
                Set theHandles(p) = New clsTrackingHandle
                If propVal Then
                    CopyMemory tmpObject, propVal, 4&
                    Set theHandles(p).This = tmpObject
                    CopyMemory tmpObject, 0&, 4&
                    theHandles(p).This.ZOrder
                    Set theHandles(p).CallBackObject = Me
                    theHandles(p).Mode = p
                End If
            Next
        End If
    
    End If
    
    If newTracker Is Nothing Then
        RemoveProp m_ParentHwnd, "waTracker.v1"
        m_ParentHwnd = 0&
        Set theContainer = Nothing
    End If
    
    Set CreateTracker = newTracker
    
End Function

Public Sub TerminateTracking(Optional ByVal bRaiseEvent As Boolean)
    ' terminates tracking, unloads tracking handles & optionally fires an event

    m_IsTracking = False
    ShowHandles False
    Erase theHandles
    If m_ParentHwnd Then
        RemoveProp m_ParentHwnd, "waTracker.v1"
        m_ParentHwnd = 0&
    End If
    If bRaiseEvent Then RaiseEvent TrackingTerminated(True)

End Sub

Public Function InitializeTracking(ByVal X As Single, ByVal Y As Single, theGuides As clsWApath) As Boolean

    ' Initializes tracking defaulting to Move while mouse is down
    ' Call this on the MouseDown event
    
    If theGuides Is Nothing Then Exit Function

    Dim mPt As POINTAPI
    
    ' this should be called due to a mouse down event
    ' validate that...
    If GetCapture = 0& Then
        MoveWindow theHandles(0).This.hwnd, -100, -100, HANDLESIZE, HANDLESIZE, 0&
        PositionHandles
        ShowHandles True
    Else
        ' move mouse capture to one of our handles
        GetCursorPos mPt
        ReleaseCapture
        ScreenToClient theContainer.hwnd, mPt
        MoveWindow theHandles(0).This.hwnd, mPt.X - 1, mPt.Y - 1, HANDLESIZE, HANDLESIZE, 0&
        SendMessage theHandles(0).This.hwnd, WM_LBUTTONDOWN, 0&, ByVal 0&
    End If
    InitializeTracking = True

End Function

Private Sub ShowHandles(MakeVisible As Boolean)
    
    ' hides/shows sizing handles
    If Not theHandles(0) Is Nothing Then
        Dim I As Long
        For I = 1 To UBound(theHandles)
            theHandles(I).This.Visible = MakeVisible
        Next
    End If
    
End Sub

Private Function PositionHandles() As Boolean

    ' positions sizing handles around the tracked object
    
    Dim p As Long, hPts(0 To 10) As POINTF
    Dim Cx As Single, Cy As Single
    Dim srcPts() As Single
    
    
    ' need 11 points: 8 for sizing, 1 for moving, 2 for rotation
    If Not theContainer Is Nothing Then
        RaiseEvent GetBoundingPoints(srcPts)
        For p = 1 To 9
            Select Case p
            Case 1
                Cx = srcPts(0, 3) - srcPts(0, 0): Cy = srcPts(1, 3) - srcPts(1, 0)
                hPts(p).X = srcPts(0, 0) + 0.5 * Cx
                hPts(p).Y = srcPts(1, 0) + 0.5 * Cy
            Case 2:
                hPts(p).X = srcPts(0, 0)
                hPts(p).Y = srcPts(1, 0)
            Case 3
                Cx = srcPts(0, 1) - srcPts(0, 0): Cy = srcPts(1, 1) - srcPts(1, 0)
                hPts(p).X = srcPts(0, 0) + 0.5 * Cx
                hPts(p).Y = srcPts(1, 0) + 0.5 * Cy
            Case 4:
                hPts(p).X = srcPts(0, 1)
                hPts(p).Y = srcPts(1, 1)
            Case 5
                Cx = srcPts(0, 2) - srcPts(0, 1): Cy = srcPts(1, 2) - srcPts(1, 1)
                hPts(p).X = srcPts(0, 1) + 0.5 * Cx
                hPts(p).Y = srcPts(1, 1) + 0.5 * Cy
            Case 6:
                hPts(p).X = srcPts(0, 2)
                hPts(p).Y = srcPts(1, 2)
            Case 7
                Cx = srcPts(0, 3) - srcPts(0, 2): Cy = srcPts(1, 3) - srcPts(1, 2)
                hPts(p).X = srcPts(0, 2) + 0.5 * Cx
                hPts(p).Y = srcPts(1, 2) + 0.5 * Cy
            Case 8:
                hPts(p).X = srcPts(0, 3)
                hPts(p).Y = srcPts(1, 3)
            Case 9:
                If hPts(4).Y < hPts(5).Y Then
                    hPts(p).X = hPts(4).X '- HANDLESIZE * 2
                    hPts(p).Y = hPts(4).Y - HANDLESIZE * 2
                    hPts(p + 1).X = hPts(8).X '+ HANDLESIZE * 2
                    hPts(p + 1).Y = hPts(8).Y + HANDLESIZE * 2
                Else
                    hPts(p).X = hPts(4).X '- HANDLESIZE * 2
                    hPts(p).Y = hPts(4).Y + HANDLESIZE * 2
                    hPts(p + 1).X = hPts(8).X '+ HANDLESIZE * 2
                    hPts(p + 1).Y = hPts(8).Y - HANDLESIZE * 2
                End If
            End Select
        Next
        
        For p = 1 To 10
            ' use APIs to avoid ScaleMode issues
            MoveWindow theHandles(p).This.hwnd, hPts(p).X - HANDLESIZE \ 2 - 1, hPts(p).Y - HANDLESIZE \ 2 - 1, HANDLESIZE, HANDLESIZE, 0
        Next
        
        PositionHandles = True
    End If

ExitRoutine:
    If Err Then
        Err.Clear
        ShowHandles False
    End If
End Function

Private Function CalculateAngle(ByVal X As Single, ByVal Y As Single, RestrictAngle As Boolean) As Single
    
    ' returns an Angle vs Radian to be written to the object's class property
    
    Dim Radian As Single
    Dim xDiff As Single
    Dim yDiff As Single
    Dim Angle As Single
    
    xDiff = -X
    yDiff = -Y
    
    If Abs(xDiff) < 0.00000001 Then xDiff = 0!
    If Abs(yDiff) < 0.00000001 Then yDiff = 0!
    
    If xDiff <> 0! Then
        Radian = Atn(yDiff / xDiff)
        If xDiff < 0 Then Radian = Radian + PI
    Else
        If yDiff < 0 Then
            Radian = 3 * PI / 2 '90
        Else
            Radian = PI / 2 '270
        End If
    End If
    
    If Radian < 0! Then Radian = Radian + PI * 2
    Angle = Radian * (180 / PI)
    
    ' restrict rotation to n degree increments
    If RestrictAngle Then
        Angle = Int((Angle \ ANGLERESTRICTION) * ANGLERESTRICTION)
    End If
        
    CalculateAngle = Angle

End Function

Private Sub ITrackingCallback_TrackingPointChanged(ByVal Key As String, ByVal X As Single, ByVal Y As Single, ByVal Shift As Integer, ByVal TrackingMode As eTrackingModes, Cancel As Boolean)

    ' A bit cumbersome and complicated, but this routine rubberbands both
    ' rotated and non-rotated objects, in both aspect & non-aspection ratios.
    
    Dim ChangeInX As Single, ChangeInY As Single
    Dim tScaleDiffX As Single, tScaleDiffY As Single
    Dim newCtrX As Single, newCtrY As Single
    Dim bCancel As Boolean, angleRadian As Single
    Dim Angle As Single, ScaledX As Single, ScaledY As Single
    
    ' forward message to user & honor Cancel if passed
    RaiseEvent TrackingChanged(X, Y, Shift, TrackingMode, bCancel)
    
    If bCancel = False Then
    
        Shift = (Shift And vbShiftMask) ' is shift pressed?
        ' When True & rotation, rotation is restricted to 30 degree angles else any angle
        ' When True & sizing, sizing is restricted to aspect ratio & no-mirroring allowed
        '       else sizing can be done at any aspect ratio & mirroring allowed
                
        If TrackingMode >= waRotate Then ' rotation, get angle from mouse coords
            ChangeInX = CalculateAngle(m_PointOrigin.X - (X - m_MousePos.X), m_PointOrigin.Y - (Y - m_MousePos.Y), (Shift = vbShiftMask))
            If TrackingMode > waRotate Then ChangeInX = ChangeInX + 180
            If (ChangeInX = m_Angle) Then ' no change in angle
                bCancel = True
            Else    ' erase previous outline, set angle & redraw outline
                m_Angle = ChangeInX
                m_XorFlag = m_XorFlag Xor True
                RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
                RaiseEvent Rotation(ChangeInX, True)
                m_XorFlag = m_XorFlag Or True
                RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
                bCancel = True
            End If
        Else
            ' erase previous outline
            m_XorFlag = m_XorFlag Or True
            RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
            
            If TrackingMode = waMoving Then ' moving: erase set new position & redraw outline
                RaiseEvent MoveObject(m_PointOrigin.X + (X - m_MousePos.X), m_PointOrigin.Y + (Y - m_MousePos.Y))
                m_XorFlag = m_XorFlag Xor True
                RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
                bCancel = True
            End If
        End If
        
        If Not bCancel Then
    
            angleRadian = -m_Radian ' use reverse rotation
            
            Select Case TrackingMode
            Case waWE ' resizing west to east only ->
                ' determine change in X only
                ChangeInX = X * Cos(angleRadian) + (-Sin(angleRadian) * Y)
                ChangeInX = (ChangeInX - (m_MousePos.X * (Cos(angleRadian)) + (-Sin(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInY = ChangeInX * m_XScaleRatio
                Else
                    ' calculate new center point, rotation cursor to the right in -> direction
                    angleRadian = (-m_Angle + 180) * PI / 180
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInX / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInX / 2)
                End If
            
            Case waNWSE ' resizing northwest to southeast only
                ' determine change in X & Y
                ChangeInX = X * Cos(angleRadian) + (-Sin(angleRadian) * Y)
                ChangeInY = X * (Sin(angleRadian)) + (Cos(angleRadian) * Y)
                ChangeInX = (ChangeInX - (m_MousePos.X * (Cos(angleRadian)) + (-Sin(angleRadian) * m_MousePos.Y)))
                ChangeInY = (ChangeInY - (m_MousePos.X * (Sin(angleRadian)) + (Cos(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInX = ChangeInY / m_XScaleRatio
                    ChangeInY = ChangeInX * m_XScaleRatio
                Else
                    ' calculate new center point, rotationg Y down & then rotation X right ->
                    angleRadian = (-m_Angle + 90) * PI / 180
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInY / 2)
                    angleRadian = (-m_Angle + 180) * PI / 180
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInX / 2)
                End If
                
            Case waNS ' resizing north to south only
                ' determine change in Y only
                ChangeInY = X * (Sin(angleRadian)) + (Cos(angleRadian) * Y)
                ChangeInY = ChangeInY - (m_MousePos.X * (Sin(angleRadian)) + (Cos(angleRadian) * m_MousePos.Y))
                
                If Shift = vbShiftMask Then
                    ChangeInX = ChangeInY / m_XScaleRatio
                Else
                    ' calculate new center point, rotating up & applying to both X & Y
                    angleRadian = (-m_Angle + 90) * PI / 180
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInY / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInY / 2)
                End If
                
            Case waNESW ' resizing northeast to southwest only
                ' determine change in Y only
                ChangeInX = X * Cos(angleRadian) + (-Sin(angleRadian) * Y)
                ChangeInY = X * (Sin(angleRadian)) + (Cos(angleRadian) * Y)
                ChangeInX = -(ChangeInX - (m_MousePos.X * (Cos(angleRadian)) + (-Sin(angleRadian) * m_MousePos.Y)))
                ChangeInY = (ChangeInY - (m_MousePos.X * (Sin(angleRadian)) + (Cos(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInX = ChangeInY / m_XScaleRatio
                    ChangeInY = ChangeInX * m_XScaleRatio
                Else
                    ' calculate new center point, rotating X left
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInX / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInX / 2)
                    ' now rotate Y down
                    angleRadian = (-m_Angle + 90) * PI / 180
                    newCtrX = newCtrX - Cos(angleRadian) * (ChangeInY / 2)
                    newCtrY = newCtrY - Sin(-angleRadian) * (ChangeInY / 2)
                End If
                
            Case waEW ' resizing east to west only
                ' determine change in X only
                ChangeInX = X * Cos(angleRadian) + (-Sin(angleRadian) * Y)
                ChangeInX = -(ChangeInX - (m_MousePos.X * (Cos(angleRadian)) + (-Sin(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInY = ChangeInX * m_XScaleRatio
                Else
                    ' calculate new center point, rotating X & Y left
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInX / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInX / 2)
                End If
            
            Case waSENW ' resizing southeast to northwest only
                ' determine change in X & Y
                ChangeInX = X * Cos(angleRadian) + (-Sin(angleRadian) * Y)
                ChangeInY = X * (Sin(angleRadian)) + (Cos(angleRadian) * Y)
                ChangeInX = -(ChangeInX - (m_MousePos.X * (Cos(angleRadian)) + (-Sin(angleRadian) * m_MousePos.Y)))
                ChangeInY = -(ChangeInY - (m_MousePos.X * (Sin(angleRadian)) + (Cos(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInX = ChangeInY / m_XScaleRatio
                    ChangeInY = ChangeInX * m_XScaleRatio
                Else
                    ' calculate new center point, rotating X left
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInX / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInX / 2)
                    ' now rotate Y up
                    angleRadian = (-m_Angle - 90) * PI / 180
                    newCtrX = newCtrX - Cos(angleRadian) * (ChangeInY / 2)
                    newCtrY = newCtrY - Sin(-angleRadian) * (ChangeInY / 2)
                End If
                
            Case waSN ' resizing south to north only
                ' determine change in Y only
                ChangeInY = X * (Sin(angleRadian)) + (Cos(angleRadian) * Y)
                ChangeInY = -(ChangeInY - (m_MousePos.X * (Sin(angleRadian)) + (Cos(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInX = ChangeInY / m_XScaleRatio
                Else
                    ' calculate new center point, rotating both X & Y up
                    angleRadian = (-m_Angle - 90) * PI / 180
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInY / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInY / 2)
                End If
                
            Case waSWNE ' resizing southwest to northeast only
                
                ' determine change in X & Y
                ChangeInX = X * Cos(angleRadian) + (-Sin(angleRadian) * Y)
                ChangeInY = X * (Sin(angleRadian)) + (Cos(angleRadian) * Y)
                ChangeInX = (ChangeInX - (m_MousePos.X * (Cos(angleRadian)) + (-Sin(angleRadian) * m_MousePos.Y)))
                ChangeInY = -(ChangeInY - (m_MousePos.X * (Sin(angleRadian)) + (Cos(angleRadian) * m_MousePos.Y)))
                If Shift = vbShiftMask Then
                    ChangeInX = ChangeInY / m_XScaleRatio
                    ChangeInY = ChangeInX * m_XScaleRatio
                Else
                    ' calculate new center point, rotating X to the right ->
                    angleRadian = (-m_Angle + 180) * PI / 180
                    newCtrX = m_PointOrigin.X - Cos(angleRadian) * (ChangeInX / 2)
                    newCtrY = m_PointOrigin.Y - Sin(-angleRadian) * (ChangeInX / 2)
                    ' now rotate Y up
                    angleRadian = (-m_Angle - 90) * PI / 180
                    newCtrX = newCtrX - Cos(angleRadian) * (ChangeInY / 2)
                    newCtrY = newCtrY - Sin(-angleRadian) * (ChangeInY / 2)
                End If
                
            Case Else
                'Case waMoving & waRotate ' moving & rotation :: handled above
            End Select
        
            If Not bCancel Then
                
                ' calculate new scale
                tScaleDiffX = m_ScaleBox.X - ChangeInX
                tScaleDiffY = m_ScaleBox.Y - ChangeInY
                RaiseEvent ScaleRatios(tScaleDiffX / m_SrcBox.X, tScaleDiffY / m_SrcBox.Y, True)
                
                ' now determnine X,Y coordinates for the non-rotated, scaled object
                If Shift = vbShiftMask Then
                    RaiseEvent MoveObject(m_PointOrigin.X - (tScaleDiffX / 2), m_PointOrigin.Y - (tScaleDiffY / 2))
                Else
                    RaiseEvent MoveObject(newCtrX - (tScaleDiffX / 2), newCtrY - (tScaleDiffY / 2))
                End If
                m_XorFlag = m_XorFlag Xor True
                RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
                
            End If
            
        End If
        
    End If

End Sub

Private Sub ITrackingCallback_TrackingStarted(ByVal Key As String, ByVal X As Single, ByVal Y As Single, ByVal TrackingMode As eTrackingModes)

    
    ' using APIs below for positioning to avoid ScaleMode differences
    
    Dim Radius As Single
    Dim mPt As POINTAPI
    Dim mRect As RECT
    Dim ScalerX As Single, ScalerY As Single
    
    ReleaseCapture
    ' hide handles when tracking begins
    ShowHandles False
    
    ' retrieve current object dimensions
    RaiseEvent GetDimensions(m_PointOrigin.X, m_PointOrigin.Y, m_SrcBox.X, m_SrcBox.Y)
    RaiseEvent ScaleRatios(ScalerX, ScalerY, False)
    RaiseEvent Rotation(m_Angle, False)
    
    ' calculate scaled object
    m_ScaleBox.X = m_SrcBox.X * ScalerX
    m_ScaleBox.Y = m_SrcBox.Y * ScalerY
    
    PI = Atn(1) * 4
    m_Radian = m_Angle * PI / 180
    
    m_XScaleRatio = m_SrcBox.Y / m_SrcBox.X ' number of pixels*Y to keep horizontal proportion
    m_IsTracking = False    ' reset
    
    If TrackingMode >= waRotate Then 'rotation
        
        Radius = m_ScaleBox.X / 2
        If TrackingMode = waRotate Then
            m_PointOrigin.X = -Cos(m_Radian) * Radius   ' calculate X,Y coords from zero center point
            m_PointOrigin.Y = -Sin(m_Radian) * Radius   ' to circle from any radius.
        Else
            m_PointOrigin.X = Cos(m_Radian) * Radius ' calculate X,Y coords from zero center point
            m_PointOrigin.Y = Sin(m_Radian) * Radius ' to circle from any radius.
        End If
        
    ElseIf TrackingMode > waMoving Then
        
        ' determine center of object;
        ' use API to get position vs .Left,.Top to avoid container ScaleMode issues
        GetWindowRect theHandles(1).This.hwnd, mRect
        mPt.X = mRect.Left: mPt.Y = mRect.Top
        ScreenToClient theContainer.hwnd, mPt
        Radius = m_ScaleBox.X / 2
        m_PointOrigin.X = (mPt.X + HANDLESIZE \ 2 + 1) + (Cos(m_Radian) * Radius)
        m_PointOrigin.Y = (mPt.Y + HANDLESIZE \ 2 + 1) + (Sin(m_Radian) * Radius)
        
    End If
    
    ' get ready to start tracking.
    ' We will move, not hide, the active handle. When we move them, the current X,Y coords
    ' on the handle will change, so we need to calc where that will be because we want
    ' the m_MousePos to hold the starting X,Y coordinates.
    ' Why not hide the handles? Because then they will not receive keyboard events
    ' and may not always receive mouse events, which we want
    GetWindowRect theHandles(TrackingMode).This.hwnd, mRect
    mPt.X = mRect.Left: mPt.Y = mRect.Top
    ScreenToClient theContainer.hwnd, mPt
    m_MousePos.X = mPt.X + 100
    m_MousePos.Y = mPt.Y + 100
    
    MoveWindow theHandles(TrackingMode).This.hwnd, -100, -100, HANDLESIZE, HANDLESIZE, 0&
    ' send notification to client
    RaiseEvent TrackingStarted(m_MousePos.X, m_MousePos.Y, TrackingMode)
    ' render first outline
    m_XorFlag = False
    RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
    
    ' now trigger a second start event. The handles class will simply call a change event
    ' vs calling 2 or more start events when a start event already occurred
    SendMessage theHandles(TrackingMode).This.hwnd, WM_LBUTTONDOWN, 0&, ByVal 0&
    
End Sub

Private Sub ITrackingCallback_TrackingTerminated(ByVal Key As String, ByVal Canceled As Boolean)
    m_XorFlag = m_XorFlag Or True
    m_IsTracking = False
    If Canceled = False Then
        PositionHandles
        ShowHandles True
    End If
    RaiseEvent RenderObject(theContainer.hDC, m_XorFlag)
    RaiseEvent TrackingTerminated(Canceled)
End Sub

Private Sub Class_Terminate()
    If m_ParentHwnd Then RemoveProp m_ParentHwnd, "waTracker.v1"
    If Not theHandles(0) Is Nothing Then
        ShowHandles False
        Erase theHandles
    End If
    Set theContainer = Nothing
End Sub


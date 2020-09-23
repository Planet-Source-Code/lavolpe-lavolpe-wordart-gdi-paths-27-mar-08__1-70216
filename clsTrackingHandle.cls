VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrackingHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long

Private m_Mode As Long

' cached values to be passed when Shift key toggled
    Private m_LastX As Single
    Private m_LastY As Single
    Private m_Shift As Integer

Private m_isActive As Boolean               ' whether tracking or not
Private m_CallBack As ITrackingCallback     ' reference back to clsWATracker

Private WithEvents theHandle As VB.PictureBox ' our handle
Attribute theHandle.VB_VarHelpID = -1

Public Function CreateHandle(ByVal HANDLESIZE As Long, hostContainerParent As Object, ByVal handleName As String, ByVal Mode As Long, CallBackObj As ITrackingCallback, hostContainer As Object) As Boolean

    If hostContainerParent Is Nothing Then Exit Function
    
    On Error Resume Next
    Set theHandle = hostContainerParent.Controls.Add("VB.PictureBox", handleName, hostContainer)
    m_isActive = False
    
ExitRoutine:
    If Err Then
        Err.Clear
        Stop
        Resume
    Else
        m_Mode = Mode
        With theHandle
            .ScaleMode = vbPixels
            .Width = HANDLESIZE
            .Height = HANDLESIZE
            .Appearance = 0
            .BorderStyle = 1
            .Enabled = True
            .MousePointer = vbCrosshair
            .ZOrder
        End With
        Set m_CallBack = CallBackObj
        CreateHandle = True
    End If
End Function

Public Property Get This() As VB.PictureBox
    Set This = theHandle
End Property
Public Property Set This(theObject As VB.PictureBox)
    Set m_CallBack = Nothing
    Set theHandle = theObject
End Property

Public Property Get CallBackObject() As ITrackingCallback
    Set CallBackObject = m_CallBack
End Property
Public Property Set CallBackObject(theObject As ITrackingCallback)
    Set m_CallBack = theObject
End Property

Public Property Get Mode() As eTrackingModes
    Mode = m_Mode
End Property
Public Property Let Mode(newMode As eTrackingModes)
    m_Mode = newMode
End Property

Private Sub Class_Terminate()
    Set m_CallBack = Nothing
    Set theHandle = Nothing
End Sub

Private Sub theHandle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        m_isActive = False
        If GetCapture = theHandle.hwnd Then ReleaseCapture
        Call m_CallBack.TrackingTerminated(vbNullString, True)
    ElseIf m_isActive Then
        If KeyCode = vbKeyShift Then
            If (m_Shift And vbShiftMask) = 0& Then
                m_Shift = m_Shift Or vbShiftMask
                Call m_CallBack.TrackingPointChanged(vbNullString, m_LastX, m_LastY, m_Shift, m_Mode, False)
            End If
        End If
    End If
End Sub

Private Sub theHandle_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_isActive Then
        If KeyCode = vbKeyShift Then
            If (m_Shift And vbShiftMask) = vbShiftMask Then
                m_Shift = (m_Shift And Not vbShiftMask)
                Call m_CallBack.TrackingPointChanged(vbNullString, m_LastX, m_LastY, m_Shift, m_Mode, False)
            End If
        End If
    End If
End Sub

Private Sub theHandle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If Not m_CallBack Is Nothing Then
            m_LastX = x: m_LastY = y: m_Shift = Shift
            If m_isActive Then
                Call m_CallBack.TrackingPointChanged(vbNullString, x, y, Shift, m_Mode, False)
            Else
                m_isActive = True
                Call m_CallBack.TrackingStarted(vbNullString, x, y, m_Mode)
            End If
        End If
    End If
End Sub

Private Sub theHandle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ExitRoutine
    Dim bCancel As Boolean
    If m_isActive Then
        m_LastX = x: m_LastY = y: m_Shift = Shift
        Call m_CallBack.TrackingPointChanged(vbNullString, x, y, Shift, m_Mode, bCancel)
        If bCancel Then m_isActive = False
    End If
    
ExitRoutine:
    If Err Then
        Err.Clear
        m_isActive = False
    End If
End Sub

Private Sub theHandle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_isActive Then
        If Button = vbLeftButton Then
            m_isActive = False
            Call m_CallBack.TrackingTerminated(vbNullString, False)
        End If
    End If
End Sub



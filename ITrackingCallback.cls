VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITrackingCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTrackingModes
    waMoving = 0
    waWE = 1
    waNWSE = 2
    waNS = 3
    waNESW = 4
    waEW = 5
    waSENW = 6
    waSN = 7
    waSWNE = 8
    waRotate = 9
    waWarp = 10
End Enum

' FEEDBACK RELATED TO TRACKING PATH POSITIONING/SIZING/ROTATION
' By using an Implementation, you can get feedback from an array of cWordArt classes

' If you want to use the tracker to allow users to modify the cWordArt object's size,
' position, rotation, you must Implement this class in the form that holds the path's renderings

' If you do not want this tracking ability in the project, you need to tweak the project
' - Remove the ITrackingCallback class
' - Remove the clsWATracker class
' - Remove the clsTrackingHandle class
' - Within the cWordArt class, remove the following routines:
'       Private Sub m_Tracker_Terminated(Cancel As Boolean, PostNotify As Boolean)
'       Public Function Tracking_Begin(ByVal hostContainer As Object, CallbackImplementation As ITrackingCallback, ByVal currentX As Single, ByVal currentY As Single) as Boolean
'       Public Sub Tracking_End()
' - Also remove the following declarations
'       Private WithEvents m_Tracker As clsWATracker
'       Private client_Tracker As ITrackingCallback
'       Private m_BackupPath As clsWApath
' - Also remove the lines containing m_Tracker from: Class_Initialize & Class_Terminate
' - Also remove the lines containing m_BackupPath from: Class_Initialize & Class_Terminate


Public Sub TrackingTerminated(ByVal Key As String, ByVal Canceled As Boolean)
    ' method is triggered when tracking is finished.
    ' Cancel=True if tracking terminated other than by a Mouse Up event
    ' When this happens, ideally you want to revert back to previous saved state. The cWordArt does this
End Sub

Public Sub TrackingStarted(ByVal Key As String, ByVal X As Single, ByVal Y As Single, ByVal TrackingMode As eTrackingModes)
    ' method is triggered when a Mouse Down occurs on one of the sizing handles
    ' When this happens, you should prepare your DC for drawing.
    ' -- At a minimum, set AutoRedraw=False to allow XOR rendering of the guides
End Sub

Public Sub TrackingPointChanged(ByVal Key As String, ByVal X As Single, ByVal Y As Single, ByVal Shift As Integer, ByVal TrackingMode As eTrackingModes, ByRef Cancel As Boolean)
    ' method is triggered when a tracking point is changed.
    ' The parameters have not much meaning to the user, except the last one:
    ' If you pass Cancel for whatever reasons, the changed tracking is not processed -- it is bypassed
End Sub


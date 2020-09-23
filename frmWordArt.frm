VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWordArt 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   " Save As - Examples "
      Height          =   1515
      Left            =   6135
      TabIndex        =   25
      Top             =   5430
      Width           =   2640
      Begin VB.OptionButton optSaveAs 
         Caption         =   "JPG"
         Height          =   225
         Index           =   1
         Left            =   1710
         TabIndex        =   27
         Top             =   525
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optSaveAs 
         Caption         =   "BMP"
         Height          =   225
         Index           =   2
         Left            =   930
         TabIndex        =   33
         Top             =   525
         Width           =   1050
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save To File"
         Height          =   375
         Left            =   615
         TabIndex        =   32
         Top             =   1110
         Width           =   1455
      End
      Begin VB.CommandButton cmdBkgColor 
         Caption         =   "..."
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Tag             =   "16777215"
         ToolTipText     =   "Optional Bkg Color"
         Top             =   255
         Width           =   480
      End
      Begin VB.TextBox txtJPGquality 
         Height          =   300
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "80"
         Top             =   795
         Width           =   570
      End
      Begin VB.OptionButton optSaveAs 
         Caption         =   "PNG"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Bkg Color (default White)"
         Height          =   255
         Index           =   7
         Left            =   675
         TabIndex        =   31
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "JPG Quality (30-100)"
         Height          =   240
         Index           =   6
         Left            =   150
         TabIndex        =   29
         Top             =   840
         Width           =   1725
      End
   End
   Begin VB.ComboBox cboSample 
      Height          =   315
      ItemData        =   "frmWordArt.frx":0000
      Left            =   165
      List            =   "frmWordArt.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5730
      Width           =   3525
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Position Test"
      Height          =   435
      Left            =   2595
      TabIndex        =   18
      Top             =   6105
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008080&
      Caption         =   "Rotation Test"
      Height          =   435
      Left            =   1425
      TabIndex        =   17
      Top             =   6105
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00008080&
      Caption         =   "Path Tracing/Tracking Test"
      Height          =   405
      Left            =   1410
      TabIndex        =   16
      Top             =   6540
      Width           =   2280
   End
   Begin VB.TextBox txtInfo 
      Height          =   1380
      Left            =   3750
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmWordArt.frx":00A2
      Top             =   5520
      Width           =   2340
   End
   Begin VB.CommandButton Command4 
      Caption         =   "design time testing"
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   4515
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   135
      ScaleHeight     =   5205
      ScaleWidth      =   5970
      TabIndex        =   11
      Top             =   210
      Width           =   6000
      Begin VB.Image Image1 
         Height          =   720
         Left            =   5205
         Picture         =   "frmWordArt.frx":00A8
         Top             =   4470
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7995
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Tahoma"
      FontSize        =   22
   End
   Begin VB.Frame Frame1 
      Caption         =   " Options "
      Height          =   4110
      Left            =   6225
      TabIndex        =   0
      Top             =   150
      Width           =   2355
      Begin VB.ComboBox cboUnicode 
         Height          =   315
         ItemData        =   "frmWordArt.frx":05AC
         Left            =   90
         List            =   "frmWordArt.frx":05C5
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1860
         Width           =   2115
      End
      Begin VB.ComboBox cboVertical 
         Height          =   315
         ItemData        =   "frmWordArt.frx":060D
         Left            =   90
         List            =   "frmWordArt.frx":0617
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1305
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   1395
         TabIndex        =   3
         Top             =   3540
         Width           =   870
      End
      Begin VB.CheckBox chkOpts 
         Caption         =   "Add Emboss"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   3495
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CheckBox chkOpts 
         Caption         =   "Add Shadow"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   3255
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CommandButton cmdFontAttr 
         Caption         =   "..."
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   2970
         Width           =   375
      End
      Begin VB.TextBox txtDisplay 
         Height          =   600
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmWordArt.frx":0640
         Top             =   2295
         Width           =   2115
      End
      Begin VB.ComboBox cboJustify 
         Height          =   315
         ItemData        =   "frmWordArt.frx":0664
         Left            =   90
         List            =   "frmWordArt.frx":0674
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   765
         Width           =   2130
      End
      Begin VB.CheckBox chkOpts 
         Caption         =   "Warp Perspective"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   255
         Width           =   2175
      End
      Begin VB.CheckBox chkOpts 
         Caption         =   "Show Guides"
         Height          =   270
         Index           =   4
         Left            =   105
         TabIndex        =   1
         Top             =   3810
         Width           =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "Other Languages "
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   22
         Top             =   1650
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Vertical Perspective"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1095
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Font Attributes"
         Height          =   240
         Index           =   2
         Left            =   510
         TabIndex        =   8
         Top             =   3015
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Horizontal Perspective"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   555
         Width           =   2100
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Some Samples to Play With"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   24
      Top             =   5520
      Width           =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "Click on object to resize/rotate.  Holding shift while sizing keeps aspect ratio"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   15
      Top             =   15
      Width           =   5955
   End
   Begin VB.Label Label2 
      Caption         =   "Adding embossing will activate optional extra pixel spacing to prevent char edges from touching"
      Height          =   1035
      Left            =   6255
      TabIndex        =   14
      Top             =   4350
      Width           =   2295
   End
End
Attribute VB_Name = "frmWordArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal mtoken As Long)
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (ByRef mtoken As Long, ByRef mInput As GdiplusStartupInput, ByRef mOutput As Any) As Long
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Function GdipWidenPath Lib "GdiPlus.dll" (ByVal mNativePath As Long, ByVal mPen As Long, ByVal mMatrix As Long, ByVal mFlatness As Single) As Long
Private Declare Function GdipWindingModeOutline Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mMatrix As Long, ByVal mFlatness As Single) As Long

' simple unicode text support (read only)
' Clipboard functions:
Private Declare Function OpenClipboard Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "USER32" () As Long
Private Declare Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "USER32" (ByVal wFormat As Long) As Long

' Memory functions:
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const CF_UNICODETEXT = 13


Private myWordArt As cWordArt
Private gdiToken As Long
Private theText As String

Private m_IsTracking As Boolean
Implements ITrackingCallback

Private Sub cboJustify_Click()
    
    If cboJustify.Tag = vbNullString Then
        If cboJustify.ListIndex = 3 Then
            myWordArt.Alignment = waAlignStretchHorizontal Or cboVertical.ListIndex * waAlignAllCharsSameSizeVertical
        Else
            myWordArt.Alignment = cboJustify.ListIndex Or cboVertical.ListIndex * waAlignAllCharsSameSizeVertical
        End If
        
        If m_IsTracking Then myWordArt.Tracking_End True
        Call Command1_Click
    
    Else
        cboJustify.Tag = vbNullString
    End If
End Sub

Private Sub cboSample_Click()
    If cboSample.Tag = vbNullString Then
        Call ShowSample
    Else
        cboSample.Tag = vbNullString
    End If
End Sub

Private Sub cboUnicode_Click()
    If cboUnicode.Tag = vbNullString Then
        If cboUnicode.ListIndex = 0& Then
            myWordArt.Caption = txtDisplay.Text
        Else
            myWordArt.Caption = GetUnicodeSample(cboUnicode.ListIndex)
        End If
        If m_IsTracking Then myWordArt.Tracking_End True
        Call Command1_Click
    Else
        cboUnicode.Tag = vbNullString
    End If
End Sub

Private Sub cboVertical_Click()
    If cboVertical.Tag = vbNullString Then
        If cboJustify.ListIndex = 3 Then
            myWordArt.Alignment = waAlignStretchHorizontal Or cboVertical.ListIndex * waAlignAllCharsSameSizeVertical
        Else
            myWordArt.Alignment = cboJustify.ListIndex Or cboVertical.ListIndex * waAlignAllCharsSameSizeVertical
        End If
        If m_IsTracking Then myWordArt.Tracking_End True
        Call Command1_Click
    Else
        cboVertical.Tag = vbNullString
    End If
End Sub

Private Sub cmdBkgColor_Click()
    On Error GoTo ExitRoutine
    CommonDialog1.Color = Val(cmdBkgColor.Tag)
    CommonDialog1.ShowColor
    cmdBkgColor.Tag = CommonDialog1.Color
ExitRoutine:
End Sub

Private Sub cmdFontAttr_Click()
    On Error GoTo ExitRoutine
    With CommonDialog1
        .Flags = cdlCFScalableOnly Or cdlCFTTOnly Or cdlCFScreenFonts
        .FontName = myWordArt.Font.name
        .FontItalic = myWordArt.Font.Italic
        .FontBold = myWordArt.Font.Bold
        .FontSize = myWordArt.Font.Size
    End With
    CommonDialog1.ShowFont
    Dim tFont As StdFont
    Set tFont = myWordArt.Font
    With CommonDialog1
        tFont.name = .FontName
        tFont.Bold = .FontBold
        tFont.Italic = .FontItalic
        tFont.Size = .FontSize
    End With
    Set myWordArt.Font = tFont
    myWordArt.LineSpacing = tFont.Size / 2
    If m_IsTracking Then myWordArt.Tracking_End True
    Call Command1_Click
ExitRoutine:
End Sub

Private Sub cmdSave_Click()
    Dim fnr As Integer
    Dim sFile As String, sExt As String
    Dim arrBytes() As Byte
    
    On Error Resume Next
    
    If optSaveAs(0) = True Then
        sExt = ".png"
    ElseIf optSaveAs(1) = True Then
        sExt = ".jpg"
    Else
        sExt = ".bmp"
    End If
    
    sFile = App.Path & "\_SamplePath" & sExt
    If Len(Dir$(sFile)) Then
        Kill sFile
        If Err Then
            MsgBox Err.Description, vbOKOnly, "Error"
            Err.Clear
            Exit Sub
        End If
    End If
    Select Case True
    Case optSaveAs(0) ' png
        If myWordArt.SaveAsPNG(arrBytes, Val(cmdBkgColor.Tag)) = False Then Exit Sub
    Case optSaveAs(1) ' jpg
        If myWordArt.SaveAsJPG(arrBytes, Val(cmdBkgColor.Tag), Val(txtJPGquality)) = False Then Exit Sub
    Case Else
        If myWordArt.SaveAsBMP(arrBytes, Val(cmdBkgColor.Tag)) = False Then Exit Sub
    End Select
    
    fnr = FreeFile()
    Open sFile For Binary As #fnr
    Put #fnr, 1, arrBytes()
    Close #fnr
    If Err Then
        MsgBox Err.Description, vbOKOnly, "Error"
        Err.Clear
    Else
        MsgBox "File saved to project folder as: _SamplePath" & sExt
    End If
    
End Sub

Private Sub Command2_Click()

    ' rotation test. A test to ensure rotation and scaling performed as expected

    If m_IsTracking Then myWordArt.Tracking_End True

    Dim AngleStart As Single, AngleStop As Single
    Dim ScaleStep As Single, AngleStep As Single
    Dim ScalerX As Single, ScalerY As Single
    Dim Cx As Single, Cy As Single
    Dim I As Long, Angle As Single
    Dim oldInfo As String
    
    oldInfo = txtInfo.Text
    txtInfo = "This was a test I used to see if my scaling/rotation/offsetting algorithms worked correctly." _
        & vbNewLine & vbNewLine & "Note. Sometimes IDE may stall while running test routines"
    txtInfo.Refresh
    
    AngleStep = 20
    ScaleStep = AngleStep / 360
    AngleStop = 360
    
    myWordArt.GetBoundingRect 0!, 0!, Cx, Cy
    ScalerX = 1!: ScalerY = 1!
    
    For I = 1 To 3
        For Angle = AngleStart To AngleStop Step AngleStep
            
            picCanvas.Cls
            With myWordArt
                .Angle = Angle
                .ScaleX = ScalerX
                .ScaleY = ScalerY
                If chkOpts(4) Then
                    .RenderGuides picCanvas.hDC, RGB(128, 128, 128)
                End If
                .Render picCanvas.hDC
            End With
            picCanvas.Refresh
            Sleep 1
            ScalerX = ScalerX - ScaleStep
            ScalerY = ScalerX
            myWordArt.Move (picCanvas.ScaleWidth - Cx * ScalerX) / 2, (picCanvas.ScaleHeight - Cy * ScalerY) / 2
            
        Next
        AngleStart = AngleStop
        If I = 1 Then
            Sleep 250
            AngleStop = AngleStart - 720
            AngleStep = -AngleStep
            ScaleStep = -ScaleStep
        ElseIf I = 2 Then
            DoEvents
            Sleep 1500
            ScaleStep = -ScaleStep
            AngleStep = -AngleStep
            AngleStop = AngleStart + 360
        ElseIf ScalerX <> 1! Then
            ' return to scale of 1:1, may be off by a fraction of a percent
            myWordArt.Move (picCanvas.ScaleWidth - Cx) / 2, (picCanvas.ScaleHeight - Cy) / 2
            picCanvas.Cls
            If chkOpts(4) Then
                myWordArt.RenderGuides picCanvas.hDC, RGB(128, 128, 128)
            End If
            myWordArt.Render picCanvas.hDC
            picCanvas.Refresh
        End If
    Next
    myWordArt.Angle = 0&
    myWordArt.ScaleX = 1
    myWordArt.ScaleY = 1
    txtInfo.Text = oldInfo
    
End Sub

Private Sub Command3_Click()

    ' positioning test
    ' This test was just to make sure I can position a scaled/flipped path where I wanted
    
    If m_IsTracking Then myWordArt.Tracking_End True

    Dim I As Long, Scaler As Single
    Dim Cx As Single, Cy As Single
    Dim penWidth As Long
    Dim oldInfo As String
    
    oldInfo = txtInfo.Text
    txtInfo = "This was a test I used to ensure I can place a path, on screen, where I wanted to." _
        & vbNewLine & vbNewLine & "Note. Sometimes IDE may stall while running test routines"
    txtInfo.Refresh
    
    ' should consider shadow & pen/brush offsets when positioning
    ' if you don't want them cropped by the edges of a DC or another object
    
    If chkOpts(2) = 1 Then penWidth = 2
    
    myWordArt.GetBoundingRect 0!, 0!, Cx, Cy
    For Scaler = 1 To 0.5 Step -0.5
        For I = 1 To 5
            Select Case I
                Case 1:
                    myWordArt.ScaleX = Scaler: myWordArt.ScaleY = Scaler
                    myWordArt.Move penWidth, penWidth
                Case 2:
                    myWordArt.ScaleX = -Scaler: myWordArt.ScaleY = -Scaler
                    myWordArt.Move picCanvas.ScaleWidth - penWidth, Cy * Scaler + penWidth
                Case 3:
                    myWordArt.ScaleX = Scaler: myWordArt.ScaleY = Scaler
                    myWordArt.Move picCanvas.ScaleWidth - Cx * Scaler - penWidth, picCanvas.ScaleHeight - Cy * Scaler - penWidth
                Case 4:
                    myWordArt.ScaleX = -Scaler: myWordArt.ScaleY = -Scaler
                    myWordArt.Move Scaler * Cx + penWidth, picCanvas.ScaleHeight - penWidth
                Case 5:
                    myWordArt.ScaleX = Scaler: myWordArt.ScaleY = Scaler
                    myWordArt.Move (picCanvas.ScaleWidth - Cx * Scaler) / 2, (picCanvas.ScaleHeight - Cy * Scaler) / 2
            End Select
            
            picCanvas.Cls
            If chkOpts(4) Then
                myWordArt.RenderGuides picCanvas.hDC, RGB(128, 128, 128)
            End If
            myWordArt.Render picCanvas.hDC
            picCanvas.Refresh
            Sleep 1000
        Next
    Next
    
    myWordArt.ScaleX = 1: myWordArt.ScaleY = 1
    myWordArt.Move (picCanvas.ScaleWidth - Cx) / 2, (picCanvas.ScaleHeight - Cy) / 2
    picCanvas.Cls
    If chkOpts(4) Then
        myWordArt.RenderGuides picCanvas.hDC, RGB(128, 128, 128)
    End If
    myWordArt.Render picCanvas.hDC
    picCanvas.Refresh
    txtInfo.Text = oldInfo
    
End Sub

Private Sub Command4_Click()
    
'    Dim b() As Byte
'    Dim c As clsWApath
'
'    Set c = myWordArt.SourceClass
'    c.SaveToImageArray b(), waSaveAsJPG, , 100
'
'    Dim fnr As Integer
'    fnr = FreeFile()
'    Open "C:\_Testpath.jpg" For Binary As #fnr
'    Put #fnr, 1, b()
'    Close #fnr
    
'    myWordArt.Tracking_Begin "", picCanvas, Me, 0, 0
    
    ' I use this button to test ideas and logic changes
'    myWordArt.Caption = ""
'    myWordArt.Caption = GetUnicodeSample(1)
'    Call Command1_Click
'
End Sub

Private Sub Command5_Click()

    ' Path tracing test.  In order to position things along a path, I have to be
    ' sure I can find any point on the path. This test was to visually confirm that
    
    ' The class' NavigatePath routine has an optional UseTransformation parameter.
    ' Pass that parameter as True if navigating a scaled or rotated path

    Dim tPath As clsWApath
    Dim X As Single, Y As Single, T As Single
    Dim StartX As Single, StartY As Single
    Dim oldInfo As String
    
    If m_IsTracking Then myWordArt.Tracking_End True
    
    oldInfo = txtInfo.Text
    txtInfo = "This was a test I used to ensure I can detect points on a path." _
        & vbNewLine & vbNewLine & "Note. Sometimes IDE may stall while running test routines"
    txtInfo.Refresh
    
    Const Delay As Long = 15
    Const Radius As Long = 5
    
    ' add a bunch of path items
    Set tPath = New clsWApath
    tPath.Append_Line 0, 0, 240, 0
    tPath.Append_Line 240, 0, 300, 20
    tPath.Append_Wave 300, 20, 300, 300, 30, 3
    tPath.Append_Line 300, 300, 200, 275
    tPath.Append_Ellipse 100, 250, 100, 50
    tPath.Append_Line 200, 275, 100, 275
    tPath.Append_Bezier 100, 275, 0, 200, 300, 100, 50, 100
    tPath.Append_Wave 50, 100, 0, 0, 50, 2
    
    picCanvas.Cls
    ' going to reference points a lot, have class cache until we release or terminate path
    tPath.CachePathPoints = True
    
    ' paths are always 0,0 regardless of path item positions so set it
    ' Note: DisplayWidth/DisplayHeight are scaled if applicable & can be negative if using negative scales
    tPath.Move (picCanvas.ScaleWidth - tPath.DisplayWidth) / 2, (picCanvas.ScaleHeight - tPath.DisplayHeight) / 2
    
    ' give path an outline pen
    tPath.Brush("outline").SetOutlineAttributes vbRed, 60
    tPath.Render picCanvas.hDC
    picCanvas.Refresh
    
    picCanvas.AutoRedraw = False
    picCanvas.DrawMode = 10 ' xor pen
    
    ' get first path point in relation to DC/display coordinates
    tPath.GetFirstDisplayPoint StartX, StartY
    ' show marker on point, pause & erase it
    picCanvas.Circle (StartX, StartY), Radius, vbBlue
    Sleep 500
    picCanvas.Circle (StartX, StartY), Radius, vbBlue
    
    ' now trace the path at 5 pixel steps
    For T = 5 To tPath.Length - 5 Step 5
        tPath.NavigatePath T, X, Y              ' << not using optional transformation parameter; requires offset
        picCanvas.Circle (X + StartX, Y + StartY), Radius, vbBlue
        Sleep Delay
        picCanvas.Circle (X + StartX, Y + StartY), Radius, vbBlue
    Next
    
    ' show the last point (which is also the first point in this example)
    ' but use class function anyway so it can be tested
    tPath.GetLastDisplayPoint X, Y
    picCanvas.Circle (X, Y), Radius, vbBlue
    Sleep 500
    picCanvas.Circle (X, Y), Radius, vbBlue
    
    ' now trace the path backwards
    For T = tPath.Length - 5 To 5 Step -5
        tPath.NavigatePath T, X, Y, , True  ' << using optional transformation parameter; should return display x,y coords; no offset
        picCanvas.Circle (X, Y), Radius, vbBlack
        Sleep Delay
        picCanvas.Circle (X, Y), Radius, vbBlack
    Next
    
    ' go the the first point & pause
    tPath.NavigatePath 0, X, Y, , True
    picCanvas.Circle (X, Y), Radius, vbBlack
    Sleep 600
    
    picCanvas.AutoRedraw = True
    picCanvas.DrawMode = 13 ' set back to default draw mode
    Call Command1_Click ' redraw previous screen
    txtInfo.Text = oldInfo
    
End Sub

Private Sub Form_Load()
    
    picCanvas.ScaleMode = vbPixels
    picCanvas.AutoRedraw = True
    
    Dim gdiUDT As GdiplusStartupInput
    gdiUDT.GdiplusVersion = 1
    GdiplusStartup gdiToken, gdiUDT, ByVal 0&
    If gdiToken = 0& Then
        MsgBox "Failed to initialize GDI+.  Unloading application", vbExclamation + vbOKOnly, "System Error"
        Unload Me
    End If
    
    Set myWordArt = New cWordArt
    cboJustify.Tag = "no action"    ' prevent triggering display update -- form now shown yet
    cboVertical.Tag = "no aciton"
    cboUnicode.Tag = "no action"
    cboJustify.ListIndex = 3&
    cboVertical.ListIndex = 0&
    cboUnicode.ListIndex = 0&
    
    With myWordArt.Brush("Shadow")
        .SetSolidFillAttributes RGB(128, 128, 128), 60 ' add a sample shadow brush
        .SetOffsets 6, 5
    End With
    With myWordArt.Brush("Emboss")
        .SetOutlineAttributes vbWhite, , 2
        .SetOffsets -1, -1
    End With
    
    myWordArt.Font.name = "Times New Roman"
    myWordArt.Font.Size = 16
    myWordArt.LineSpacing = 3
    myWordArt.CharacterSpacing = 1
    myWordArt.Caption = txtDisplay.Text
    
    Show
    cboSample.ListIndex = 0&
    
    MsgBox "The text box has minimal unicode support." & vbCrLf & _
        "Delete everything out of the text box and paste unicode inside of it, " & vbCrLf & _
        "then simply tab out of the text box and unicode should be displayed on screen", vbInformation + vbOKOnly
        
End Sub

Private Sub chkOpts_Click(Index As Integer)
    
    If m_IsTracking Then myWordArt.Tracking_End True
    
    Select Case Index
    Case 0
        If chkOpts(Index).Value Then myWordArt.WarpStyle = warpPerspective Else myWordArt.WarpStyle = warpBilinear
    Case 1, 2 ' shadow & embossing
        If chkOpts(Index) = 0& Then
            If Index = 1 Then
                myWordArt.Brush("Shadow").RemoveFill
            Else
                myWordArt.Brush("Emboss").RemoveOutline
                myWordArt.CharacterSpacing = 0&
            End If
        Else
            On Error Resume Next
            CommonDialog1.ShowColor
            If Err Then ' user canceled dialog
                Err.Clear
                chkOpts(Index) = 0
            Else
                On Error GoTo 0
                If Index = 1 Then
                    myWordArt.Brush("Shadow").SetSolidFillAttributes CommonDialog1.Color, 60
                    myWordArt.Brush("Shadow").SetOffsets 6, 5
                Else
                    myWordArt.Brush("Emboss").SetOutlineAttributes CommonDialog1.Color, , 2
                    myWordArt.Brush("Emboss").SetOffsets -1, -1
                    myWordArt.CharacterSpacing = 1
                End If
            End If
        End If
    Case 4 ' guides on/off
        If chkOpts(Index) Then
            With picCanvas
                .Cls
                .ForeColor = RGB(128, 128, 128)
                picCanvas.Line (.ScaleWidth / 2, 0)-(.ScaleWidth / 2, .ScaleHeight)
                picCanvas.Line (0, .ScaleHeight / 2)-(.ScaleWidth, .ScaleHeight / 2)
                .Picture = .Image
            End With
        Else
            Set picCanvas.Picture = Nothing
        End If
    End Select
    
    Call Command1_Click
ExitRoutine:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set myWordArt = Nothing
    ' ^^ contains GDI+ objects.
    ' If left undestroyed, crash potential for your app when GDI+ is unloaded
    If gdiToken Then GdiplusShutdown gdiToken
End Sub

Private Sub ITrackingCallback_TrackingPointChanged(ByVal Key As String, ByVal X As Single, ByVal Y As Single, ByVal Shift As Integer, ByVal TrackingMode As eTrackingModes, Cancel As Boolean)
    ' just fyi info here
End Sub

Private Sub ITrackingCallback_TrackingStarted(ByVal Key As String, ByVal X As Single, ByVal Y As Single, ByVal TrackingMode As eTrackingModes)
    picCanvas.Cls
    picCanvas.AutoRedraw = False
    picCanvas.Refresh
End Sub

Private Sub ITrackingCallback_TrackingTerminated(ByVal Key As String, ByVal Canceled As Boolean)
    picCanvas.Cls
    picCanvas.AutoRedraw = True
    If Canceled Then m_IsTracking = False
    Call Command1_Click
End Sub

Private Sub Command1_Click()

    Dim errString As String
    
    picCanvas.Cls
    If chkOpts(4) Then myWordArt.RenderGuides picCanvas.hDC, RGB(65, 65, 65)
    Select Case myWordArt.Render(picCanvas.hDC)
        Case 0 ' successful
        Case waErrInvalidGuide_Top: errString = "Top guide not set or has not valid path"
        Case waErrInvalidGuide_Btm: errString = "Bottom guide not set or has not valid path"
        Case waErrUnsupportedFont: errString = "Not able to process the chosen font"
        Case waErrInvalidDC: errString = "GDI+ failed to recognized DC"
        Case waErrTextNotSet: errString = "The display text is not set"
        Case waErrUnknown: errString = "Error occurred: Unknown Error Code"
    End Select
    picCanvas.Refresh
    If Len(errString) Then MsgBox errString, vbExclamation + vbOKOnly, "Error"

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        If picCanvas.MousePointer = vbSizeAll Then ' currently over path
            m_IsTracking = True
            myWordArt.Tracking_Begin "1", picCanvas, Me, X, Y
        ElseIf m_IsTracking Then
            myWordArt.Tracking_End False
            m_IsTracking = False
        End If
    End If
    
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If myWordArt.IsPathPoint(picCanvas.hDC, X, Y) Then
        picCanvas.MousePointer = vbSizeAll
    Else
        picCanvas.MousePointer = vbDefault
    End If
End Sub

Private Sub txtDisplay_Change()
    txtDisplay.Tag = "chg"
End Sub

Private Sub txtDisplay_LostFocus()
    If txtDisplay.Tag = "chg" Then
        txtDisplay.Tag = vbNullString
        
        Dim p As Long, uniCount As Long, sText As String
        For p = 1 To Len(txtDisplay.Text)
            If Mid$(txtDisplay.Text, p, 1) = "?" Then uniCount = uniCount + 1
        Next
        If uniCount > p \ 2 Then ' probably unicode,
            ' let's see if there is unicode on the clipboard
            sText = GetUnicodeText(Me.hWnd)
            If sText = vbNullString Then sText = txtDisplay.Text
            myWordArt.Caption = sText
        Else
            myWordArt.Caption = txtDisplay.Text
        End If
        ShowSample
    End If
End Sub

Private Sub ShowSample()
    
    With myWordArt
    
        .ResetGuides
        
        Select Case cboSample.ListIndex
        Case -1
        Case 0, 11 ' wave over wave
            With .Brush("Fill")
                .SetTextureBrush Image1, WrapModeTile
                .RemoveOutline
            End With
            .IndentLeft = 0: .IndentRight = 0
            If cboSample.ListIndex = 0 Then
                .AppendGuide_Wave True, 0, 20, 255, 20, -20     ' negative crests up, positive dips
                .AppendGuide_Wave False, 0, 80, 255, 80, -20
                txtInfo = "Wave over Wave" & vbNewLine & _
                    " -- Tiled bitmap brush, no outlilne pen"
            Else
                .AppendGuide_Line True, 120, 255, 120, 0
                .AppendGuide_Line False, 180, 255, 180, 0
                txtInfo = "90 Degree Parallel Lines" & vbNewLine & _
                    " -- Tiled bitmap brush, no outlilne pen"
            End If
        Case 1 ' wave over line
            With .Brush("Fill")
                .SetGradientAttributes RGB(64, 64, 255), RGB(192, 192, 255)
                .GradientStyle = LinearForwardDiagonal
                .SetOutlineAttributes vbRed
            End With
            .IndentLeft = 0: .IndentRight = 0
            .AppendGuide_Wave True, 0, 20, 300, 20, -10, 2
            .AppendGuide_Line False, 0, 80, 300, 80
            txtInfo = "Wave Over Horizontal Line" & vbNewLine & _
                " -- Diagonal Gradeint brush " & vbNewLine & " -- Red single pixel outline"
            
        Case 5, 6 ' circles
            .IndentLeft = 0: .IndentRight = 10  ' use left or right indent to prevent text end touching text start
            .AppendGuide_Circle True, 20, 0, 180
            If cboSample.ListIndex = 5 Then
                With .Brush("Fill")
                    .SetGradientAttributes vbYellow, vbMagenta
                    .GradientStyle = LinearVertical
                    .SetOutlineAttributes vbBlack
                End With
                .AppendGuide_Circle False, 70, 50, 80
                txtInfo = "Circle within a Circle" & vbNewLine & _
                    " -- Vertical Gradeint brush " & vbNewLine & " -- Black single pixel outline" & _
                    vbNewLine & " -- Uses right indentation to prevent end of line touching start of line"
            Else
                With .Brush("Fill")
                    .SetGradientAttributes vbBlue, vbCyan
                    .GradientStyle = LinearHorizontal
                    .SetOutlineAttributes vbBlack
                End With
                .AppendGuide_Circle False, 60, 40, 60
                txtInfo = "Offset Circle within a Circle" & vbNewLine & _
                    " -- Horizontal Gradeint brush " & vbNewLine & " -- Black single pixel outline" & _
                    vbNewLine & " -- Uses right indentation to prevent end of line touching start of line"
            End If
            
        Case 3 ' spline over spline
            With .Brush("Fill")
                .SetOutlineAttributes vbBlue, , , DashStyleDot
                .RemoveFill ' no fill, just outline for this example
            End With
            .IndentLeft = 0: .IndentRight = 0
            .AppendGuide_Spline True, 0, 60, 300, 80, 0.5, 20
            .AppendGuide_Spline False, 0, 120, 300, 100, 0.5, 20
            txtInfo = "Spline (curve) over Spline" & vbNewLine & _
                " -- Not filled" & vbNewLine & " -- Blue dotted outline pen"
            
        Case 4 ' line over spline
            With .Brush("Fill")
                .RemoveOutline ' no outline, just filled
                .SetGradientAttributes RGB(255, 64, 64), RGB(255, 192, 192)
                .GradientStyle = LinearForwardDiagonal
            End With
            .IndentLeft = 0: .IndentRight = 0
            .AppendGuide_Line True, 0, 60, 300, 60
            .AppendGuide_Spline False, 0, 120, 300, 90, 0.66, 20
            txtInfo = "Horizontal Line over Spline (Curve)" & vbNewLine & _
                " -- Diagonal Gradeint brush " & vbNewLine & " -- No outline"
        
        Case 2 'wave over arc
            With .Brush("Fill")
                .SetHatchBrushAttributes vbBlack, vbCyan, HatchStyleWave
                .SetOutlineAttributes vbBlue
            End With
            .IndentLeft = 0: .IndentRight = 0
            .AppendGuide_Wave True, 0, 40, 300, 40, -20
            .AppendGuide_Arc False, 0, 60, 300, 60, 180, -180
            txtInfo = "Wave over Arc" & vbNewLine & _
                " -- Hatch-Style brush " & vbNewLine & " -- Blue single pixel outline"
        
        Case 8 ' arc over arc
            With .Brush("Fill")
                .SetGradientAttributes RGB(255, 64, 64), RGB(255, 212, 212)
                .GradientStyle = LinearForwardDiagonal
                .SetOutlineAttributes vbRed
            End With
            .IndentLeft = 11: .IndentRight = 11 ' use indents to prevent edges from being to squashed
            .AppendGuide_EllipseSplit 0, 0, 300, 100
            txtInfo = "Oval (Arc over Arc)" & vbNewLine & _
                " -- Diagonal Gradeint brush " & vbNewLine & " -- Red single pixel outline" & _
                    vbNewLine & " -- Uses indentation to prevent far left/right characters from being too scrunched"
            
        Case 10 ' tent (arch over line)
            With .Brush("Fill")
                .GradientStyle = LinearBackwardDiagonal
                .SetGradientAttributes vbBlue, vbCyan
                .SetOutlineAttributes vbBlack
            End With
            .AppendGuide_Spline True, 60, 100, 320, 100, 0.5, -100 ' using spline to create arch
            .AppendGuide_Line False, 40, 120, 340, 120
            .IndentLeft = 15: .IndentRight = 15
            txtInfo = "Arc over Line" & vbNewLine & _
                " -- Diagonal Gradeint brush " & vbNewLine & " -- Black single pixel outline" & _
                    vbNewLine & " -- Uses indentation to prevent far left/right characters from being too scrunched"
            
        Case 7 ' circle in oval
            With .Brush("Fill")
                .SetHatchBrushAttributes vbBlue, vbCyan, HatchStyleTrellis
                .SetOutlineAttributes vbBlue
            End With
            .IndentLeft = 0: .IndentRight = 10 ' use left or right indent to prevent text end touching text start
            .AppendGuide_Ellipse True, 0, 0, 300, 150
            .AppendGuide_Circle False, 120, 45, 60
            txtInfo = "Oval (Circle in Oval)" & vbNewLine & _
                " -- Hatch style brush " & vbNewLine & " -- Blue single pixel outline" & _
                    vbNewLine & " -- Uses right indentation to prevent end of line touching start of line"
            
        Case 9 ' double arch
            With .Brush("Fill")
                .SetSolidFillAttributes vbYellow
                .SetOutlineAttributes vbMagenta
            End With
            .IndentLeft = 15: .IndentRight = 15
            .AppendGuide_Arc True, 0, -60, 300, 150, 180, 180
            .AppendGuide_Arc False, 0, 0, 300, 150, 180, 180
            txtInfo = "Arc over Arc" & vbNewLine & _
                " -- Solid yellow brush " & vbNewLine & " -- Magenta single pixel outline" & _
                    vbNewLine & " -- Uses indentation to prevent far left/right characters from being too scrunched"
        
        End Select
        
        ' The move command allows the option to center move.
        ' So instead of the following 2 lines, we can just call one
        '   .GetBoundingRect X, Y, Cx, Cy
        '   .Move (picCanvas.ScaleWidth - Cx) / 2, (picCanvas.ScaleHeight - Cy) / 2

        .Move picCanvas.ScaleWidth / 2, picCanvas.ScaleHeight / 2, True
        
    End With
    
    Call Command1_Click
    
End Sub

Private Function GetUnicodeSample(WhichOne As Long) As String

    Dim sUniText As String
    Dim uniT() As Integer
    
    Select Case WhichOne
    Case 1
        ReDim uniT(0 To 7)   'ARABIC
        uniT(0) = 1606
        uniT(1) = 1589
        uniT(2) = 1603
        uniT(3) = 32
        uniT(4) = 1607
        uniT(5) = 1606
        uniT(6) = 1575
    Case 2
        ReDim uniT(0 To 7)   'TRADITIONAL CHINESE
        uniT(0) = 20320
        uniT(1) = 30340
        uniT(2) = 25991
        uniT(3) = 23383
        uniT(4) = 22312
        uniT(5) = -28647
        uniT(6) = -30495
    Case 3
        ReDim uniT(0 To 18)   'GREEK
        uniT(0) = 932
        uniT(1) = 959
        uniT(2) = 32
        uniT(3) = 954
        uniT(4) = 949
        uniT(5) = 953
        uniT(6) = 956
        uniT(7) = 949
        uniT(8) = 957
        uniT(9) = 959
        uniT(10) = 32
        uniT(11) = 963
        uniT(12) = 945
        uniT(13) = 962
        uniT(14) = 32
        uniT(15) = 949
        uniT(16) = 948
        uniT(17) = 969
    Case 4
        ReDim uniT(0 To 11)   'JAPANESE
        uniT(0) = 12371
        uniT(1) = 12371
        uniT(2) = 12395
        uniT(3) = 12354
        uniT(4) = 12394
        uniT(5) = 12383
        uniT(6) = 12398
        uniT(7) = 12486
        uniT(8) = 12461
        uniT(9) = 12473
        uniT(10) = 12488
    Case 5
        ReDim uniT(0 To 12)   'KOREAN
        uniT(0) = -14868
        uniT(1) = -20944
        uniT(2) = -14896
        uniT(3) = 32
        uniT(4) = -21056
        uniT(5) = -10920
        uniT(6) = -14504
        uniT(7) = 32
        uniT(8) = -11955
        uniT(9) = -15708
        uniT(10) = -11592
        uniT(11) = -18052
    Case 6
        ReDim uniT(0 To 15)   'RUSSIAN
        uniT(0) = 1042
        uniT(1) = 1072
        uniT(2) = 1096
        uniT(3) = 32
        uniT(4) = 1090
        uniT(5) = 1077
        uniT(6) = 1082
        uniT(7) = 1089
        uniT(8) = 1090
        uniT(9) = 32
        uniT(10) = 1079
        uniT(11) = 1076
        uniT(12) = 1077
        uniT(13) = 1089
        uniT(14) = 1100
    End Select
    
    sUniText = String$(UBound(uniT) + 1, 0)
    CopyMemory ByVal StrPtr(sUniText), uniT(0), (UBound(uniT) + 1) * 2
    GetUnicodeSample = sUniText

End Function

Private Function GetUnicodeText(hWnd As Long) As String

' Returns a byte array containing binary data on the clipboard for
' format lFormatID:
Dim hMem As Long, lSize As Long, lPtr As Long
Dim sReturn As String
    
    If OpenClipboard(hWnd) Then
    
        If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then Exit Function
    
        hMem = GetClipboardData(CF_UNICODETEXT)
        ' If success:
        If (hMem <> 0) Then
            ' Get the size of this memory block:
            lSize = GlobalSize(hMem)
            ' Get a pointer to the memory:
            lPtr = GlobalLock(hMem)
            If (lSize > 0) Then
                ' Resize the byte array to hold the data:
                sReturn = String$(lSize \ 2 + 1, 0)
                ' Copy from the pointer into the array:
                CopyMemory ByVal StrPtr(sReturn), ByVal lPtr, lSize
            End If
            ' Unlock the memory block:
            GlobalUnlock hMem
            ' Success:
            GetUnicodeText = sReturn
            ' Don't free the memory - it belongs to the clipboard.
        End If
        
        CloseClipboard
    End If
    
End Function


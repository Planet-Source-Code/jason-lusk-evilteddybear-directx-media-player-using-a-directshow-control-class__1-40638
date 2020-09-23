VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDShowControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objBasicAudio As IBasicAudio
Private m_objBasicVideo As IBasicVideo
Private m_objMediaEvent As IMediaEvent
Private m_objVideoWindow As IVideoWindow
Private m_objMediaControl As IMediaControl
Private m_objMediaPosition As IMediaPosition

Private Const WS_VISIBLE = &H10000000
Dim ORGMwidth As Integer
Dim ORGMheight As Integer
Dim TargetHeight  As Integer
Dim TargetWidth  As Integer
Dim TargetLeft  As Integer
Dim TargetTop  As Integer
Public Duration As Double

Public Sub DS_Remove()
    On Local Error GoTo RemoveDShowError
    
    'If a MediaControl instance exists, then stop it from playing
    If ObjPtr(m_objMediaControl) > 0 Then
        m_objMediaControl.Stop
    End If
    'If a VideoWindow instance exists, then remove the link to the render target
    If ObjPtr(m_objVideoWindow) > 0 Then
        m_objVideoWindow.Owner = 0
    End If
    'Destroy all objects
    If ObjPtr(m_objBasicAudio) > 0 Then Set m_objBasicAudio = Nothing
    If ObjPtr(m_objBasicVideo) > 0 Then Set m_objBasicVideo = Nothing
    If ObjPtr(m_objMediaControl) > 0 Then Set m_objMediaControl = Nothing
    If ObjPtr(m_objVideoWindow) > 0 Then Set m_objVideoWindow = Nothing
    If ObjPtr(m_objMediaPosition) > 0 Then Set m_objMediaPosition = Nothing
    Exit Sub
RemoveDShowError:
    Err.Clear
    Exit Sub
End Sub
Public Function DS_CorrectAspect()

Dim ret As Long
Dim AreaWid As Long
Dim AreaHgt As Long
Dim ImgWid As Long
Dim ImgHgt As Long
Dim HScl As Double
Dim WScl As Double
Dim h As Long
Dim w As Long
Dim CenterTop As Long
Dim CenterLeft As Long
On Error Resume Next

'Get Sizes
'AreaWid = frmMain.picTarget.Width
'AreaHgt = frmMain.picTarget.Height
AreaWid = TargetWidth
AreaHgt = TargetHeight
ImgWid = ORGMwidth
ImgHgt = ORGMheight

'Get Aspect
If (ImgHgt <> AreaHgt) Or (ImgWid <> AreaWid) Then
    HScl = AreaHgt / ImgHgt
    WScl = AreaWid / ImgWid
    If WScl < HScl Then
       w = AreaWid
       h = Int(ImgHgt * WScl)
    Else
       h = AreaHgt
       w = Int(ImgWid * HScl)
    End If
       If h <= 0 Then h = 1
       If w <= 0 Then w = 1
    Else
       w = ImgWid
       h = ImgHgt
    End If

'Center Movie In Aspect
CenterLeft = (AreaWid - w) / 2
CenterTop = (AreaHgt - h) / 2

'Resize Movie to Aspect
'VideoWindow.Left = 0 'CenterLeft
'VideoWindow.Top = 0 'CenterTop
'VideoWindow.Height = h
'VideoWindow.Width = w

m_objVideoWindow.SetWindowPosition CenterLeft, CenterTop, w, h

End Function
Public Function DS_StretchMovie()

Dim AreaWid As Long
Dim AreaHgt As Long
Dim AreaLeft As Long
Dim AreaTop As Long
On Error Resume Next

'Get Sizes
AreaWid = TargetWidth
AreaHgt = TargetHeight
AreaTop = 0
AreaLeft = 0
m_objVideoWindow.SetWindowPosition AreaLeft, AreaTop, AreaWid, AreaHgt
End Function
Public Function DS_ChangeTarget(rectTop As Integer, _
                                rectLeft As Integer, _
                                rectWidth As Integer, _
                                rectHeight As Integer, _
                                hWndTarget As Long)
    
    TargetHeight = rectHeight
    TargetWidth = rectWidth
    TargetTop = rectTop
    TargetLeft = rectLeft
    
    With m_objVideoWindow
        .WindowStyle = WS_VISIBLE 'WS_VISIBLE = &H10000000
        .Owner = hWndTarget
        .Top = TargetTop
        .Left = TargetLeft
        .Width = TargetWidth
        .Height = TargetHeight
    End With
DS_CorrectAspect
    
End Function
Public Sub DS_OpenFile(Filename As String, _
                       rectTop As Integer, _
                       rectLeft As Integer, _
                       rectWidth As Integer, _
                       rectHeight As Integer, _
                       hWndTarget As Long)
    On Local Error GoTo OpenFileError
    
    'First clean up
    Call DS_Remove
    
    Set m_objMediaControl = New FilgraphManager
    Call m_objMediaControl.RenderFile(Filename)
    
    Set m_objBasicAudio = m_objMediaControl
    m_objBasicAudio.Volume = 0 'Loudest
    m_objBasicAudio.Balance = 0 'Centered
    
    Set m_objVideoWindow = m_objMediaControl

    ORGMwidth = m_objVideoWindow.Width
    ORGMheight = m_objVideoWindow.Height
    
    TargetHeight = rectHeight
    TargetWidth = rectWidth
    TargetTop = rectTop
    TargetLeft = rectLeft
    
  
    
    With m_objVideoWindow
        .WindowStyle = WS_VISIBLE 'WS_VISIBLE = &H10000000
        .Owner = hWndTarget
        .Top = rectTop
        .Left = rectLeft
        .Width = rectWidth
        .Height = rectHeight
    End With
    
DS_CorrectAspect
    
    Set m_objMediaEvent = m_objMediaControl
           
    Set m_objMediaPosition = m_objMediaControl
    m_objMediaPosition.Rate = 1 'Normal forward playback speed
    
    Exit Sub
OpenFileError:
    Err.Clear
    Resume Next
End Sub

Public Sub DS_Play()
    'Play if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then
        If CLng(m_objMediaPosition.CurrentPosition) = CLng(m_objMediaPosition.Duration) Then
            m_objMediaPosition.CurrentPosition = 0
        End If
        Call m_objMediaControl.Run
    End If
End Sub

Public Sub DS_Pause()
    'Pause if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaControl) > 0 Then
        Call m_objMediaControl.Pause
    End If
End Sub

Public Sub DS_Stop()
    'Stop if DShow is initialized and a file is loaded
    If (ObjPtr(m_objMediaControl) > 0) And (ObjPtr(m_objMediaPosition) > 0) Then
        Call m_objMediaControl.Stop
        m_objMediaPosition.CurrentPosition = 0
    End If
End Sub
Public Sub DS_Seek(Amount As Double)
    'Seek if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then
        'if position is out of bounds then correct
        If m_objMediaPosition.CurrentPosition + Amount < 0 Then
            m_objMediaPosition.CurrentPosition = 0
        ElseIf m_objMediaPosition.CurrentPosition + Amount > _
               m_objMediaPosition.Duration Then
            m_objMediaPosition.CurrentPosition = m_objMediaPosition.Duration
        Else
            'else just set the new position
            m_objMediaPosition.CurrentPosition = m_objMediaPosition.CurrentPosition + Amount
        End If
    End If
End Sub

Public Function DS_StillPlaying(msTimeout As Long) As Boolean
    On Local Error Resume Next
    Dim EvCode As Long
    'Check only if DShow is initialized and a file is loaded
    If ObjPtr(m_objMediaEvent) > 0 Then
        m_objMediaEvent.WaitForCompletion msTimeout, EvCode
        If EvCode = 0 Then
            DS_StillPlaying = True
        Else
            DS_StillPlaying = False
        End If
    End If
End Function

Public Function DS_GetRate() As Double
    'Get rate if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then DS_GetRate = m_objMediaPosition.Rate
End Function

Public Sub DS_SetRate(newRate As Double)
    'rate must not be <= 0!
    If newRate <= 0 Then newRate = 0.1
    'Set rate if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then m_objMediaPosition.Rate = newRate
End Sub

Public Function DS_GetDuration() As Double
    'Get duration if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then DS_GetDuration = m_objMediaPosition.Duration
End Function

Public Function DS_GetPosition() As Double
    'Get position if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then DS_GetPosition = m_objMediaPosition.CurrentPosition
End Function

Public Sub DS_SetPosition(newPosition As Double)
    'Set position if DShow is initialized and file is loaded
    If ObjPtr(m_objMediaPosition) > 0 Then
        'if newPosition is out of bounds then correct
        If newPosition < 0 Then newPosition = 0
        If newPosition > m_objMediaPosition.Duration Then newPosition = m_objMediaPosition.Duration
        m_objMediaPosition.CurrentPosition = newPosition
    End If
End Sub
Public Function DS_GetVolume() As Long
    'Get volume if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then DS_GetVolume = m_objBasicAudio.Volume
End Function

Public Sub DS_SetVolume(newVolume As Long)
    'Volume must be between 0 (loudest) and -10000 (disabled)
    If newVolume > 0 Then newVolume = 0
    If newVolume < -10000 Then newVolume = -10000
    'Set volume if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then m_objBasicAudio.Volume = newVolume
End Sub

Public Function DS_GetBalance() As Long
    'Get balance if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then DS_GetBalance = m_objBasicAudio.Balance
End Function

Public Sub DS_SetBalance(newBalance As Long)
    'balance must be between -10000 (left) and +10000 (right)
    If newBalance < -10000 Then newBalance = -10000
    If newBalance > 10000 Then newBalance = 10000
    'Set balance if DShow is initialized and file has audiostream
    If ObjPtr(m_objBasicAudio) > 0 Then m_objBasicAudio.Balance = newBalance
End Sub



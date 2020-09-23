VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectShow Demo"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   7560
      TabIndex        =   26
      Top             =   2040
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   7560
      Max             =   0
      Min             =   -5000
      TabIndex        =   25
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2nd Location"
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   6000
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1st Location"
      Height          =   375
      Left            =   7560
      TabIndex        =   23
      Top             =   6360
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   7560
      ScaleHeight     =   1905
      ScaleWidth      =   1965
      TabIndex        =   22
      Top             =   3960
      Width           =   1995
   End
   Begin VB.CommandButton cmdStretch 
      Caption         =   "clDShow.DS_StretchMovie"
      Height          =   315
      Left            =   5280
      TabIndex        =   21
      Top             =   420
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clDShow.DS_CorrectAspect"
      Height          =   315
      Left            =   5280
      TabIndex        =   20
      Top             =   60
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cdOpenFile 
      Left            =   8760
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSeekFor 
      Caption         =   "Seek >>"
      Height          =   315
      Left            =   8640
      TabIndex        =   18
      Top             =   1620
      Width           =   915
   End
   Begin VB.CommandButton cmdSeekBack 
      Caption         =   "<< Seek"
      Height          =   315
      Left            =   7560
      TabIndex        =   17
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmdRateAdd 
      Caption         =   "Rate +"
      Height          =   315
      Left            =   8640
      TabIndex        =   16
      Top             =   1200
      Width           =   915
   End
   Begin VB.CommandButton cmdRateSub 
      Caption         =   "Rate -"
      Height          =   315
      Left            =   7560
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdBalRgt 
      Caption         =   "Right Spk"
      Height          =   315
      Left            =   8640
      TabIndex        =   14
      Top             =   780
      Width           =   915
   End
   Begin VB.CommandButton cmdBalLft 
      Caption         =   "Left Spk"
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton cmdBalCen 
      Caption         =   "Center Balance"
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   420
      Width           =   1995
   End
   Begin VB.CommandButton cmdAddVol 
      Caption         =   "Volume +"
      Height          =   315
      Left            =   8640
      TabIndex        =   11
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdSubVol 
      Caption         =   "Volume -"
      Height          =   315
      Left            =   7560
      TabIndex        =   10
      Top             =   60
      Width           =   975
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   6720
   End
   Begin VB.PictureBox picTarget 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6360
      Left            =   60
      ScaleHeight     =   422
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   494
      TabIndex        =   7
      Top             =   780
      Width           =   7440
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "clDShow.DS_Stop"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   420
      Width           =   2175
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "clDShow.DS_Pause"
      Height          =   315
      Left            =   2700
      TabIndex        =   3
      Top             =   420
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "clDShow.DS_Play"
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   60
      Width           =   2175
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "clDShow.DS_OpenFile"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "<-- Render target   for video output               "
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblVolume 
      Caption         =   "Volume: "
      Height          =   195
      Left            =   7560
      TabIndex        =   9
      Top             =   3300
      Width           =   1995
   End
   Begin VB.Label lblBalance 
      Caption         =   "Balance:"
      Height          =   195
      Left            =   7560
      TabIndex        =   8
      Top             =   3060
      Width           =   1995
   End
   Begin VB.Label lblRate 
      Caption         =   "Rate:"
      Height          =   195
      Left            =   7560
      TabIndex        =   6
      Top             =   2820
      Width           =   1995
   End
   Begin VB.Label lblPosition 
      Caption         =   "Position:"
      Height          =   195
      Left            =   7560
      TabIndex        =   5
      Top             =   2580
      Width           =   1995
   End
   Begin VB.Label lblDuration 
      Caption         =   "Duration:"
      Height          =   195
      Left            =   7560
      TabIndex        =   1
      Top             =   2340
      Width           =   1995
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim clDShow As clsDShowControl


Private Sub Command1_Click()
If cdOpenFile.Filename = "" Then Exit Sub
clDShow.DS_CorrectAspect
End Sub

Private Sub cmdStretch_Click()
If cdOpenFile.Filename = "" Then Exit Sub
clDShow.DS_StretchMovie
End Sub

Private Sub Command2_Click()
If cdOpenFile.Filename = "" Then Exit Sub
Call clDShow.DS_ChangeTarget(0, 0, picTarget.Width, picTarget.Height, picTarget.hWnd)
End Sub

Private Sub Command3_Click()
If cdOpenFile.Filename = "" Then Exit Sub
Call clDShow.DS_ChangeTarget(0, 0, Picture1.Width, Picture1.Height, Picture1.hWnd)
clDShow.DS_CorrectAspect
End Sub

Private Sub Form_Load()
    'create a new instance of our controller class
    Set clDShow = New clsDShowControl

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cleanup DShow
    clDShow.DS_Remove
    'destroy our controller class
    Set clDShow = Nothing
End Sub

Private Sub cmdOpenFile_Click()
    'open a common dialog to search for a multi media file
    cdOpenFile.ShowOpen
    'open a file
    clDShow.DS_OpenFile cdOpenFile.Filename, 0, 0, _
                        picTarget.Width, picTarget.Height, picTarget.hWnd
    'start the timer, that updates the labels
    tmrTimer.Enabled = True
End Sub

Private Sub HScroll1_Change()
Call clDShow.DS_SetVolume(HScroll1.Value)
End Sub
Private Sub HScroll1_Scroll()
Call clDShow.DS_SetVolume(HScroll1.Value)
End Sub

Private Sub HScroll2_Change()
If HScroll2.Max <> clDShow.DS_GetDuration Then
HScroll2.Max = clDShow.DS_GetDuration
End If
Call clDShow.DS_SetPosition(HScroll2.Value)
End Sub

Private Sub lblPosition_Change()
If HScroll2.Max <> clDShow.DS_GetDuration Then
HScroll2.Max = clDShow.DS_GetDuration
End If
'HScroll2.Value = Round(1, clDShow.DS_GetPosition)
End Sub

Private Sub tmrTimer_Timer()
    'Update all label controls
    lblDuration.Caption = "Duration: " & clDShow.DS_GetDuration
    lblPosition.Caption = "Position: " & clDShow.DS_GetPosition
    lblRate.Caption = "Rate: " & clDShow.DS_GetRate
    lblVolume.Caption = "Volume: " & clDShow.DS_GetVolume
    lblBalance.Caption = "Balance: " & clDShow.DS_GetBalance
End Sub

Private Sub cmdPlay_Click()
    'start playback
    clDShow.DS_Play
End Sub

Private Sub cmdPause_Click()
    'pause playback
    clDShow.DS_Pause
End Sub

Private Sub cmdStop_Click()
    'stop playback
    clDShow.DS_Stop
End Sub

Private Sub cmdSubVol_Click()
    'lower Volume
    Call clDShow.DS_SetVolume(clDShow.DS_GetVolume - 1000)
End Sub

Private Sub cmdAddVol_Click()
    'raise Volume
    Call clDShow.DS_SetVolume(clDShow.DS_GetVolume + 1000)
End Sub

Private Sub cmdBalCen_Click()
    'center balance
    clDShow.DS_SetBalance 0
End Sub

Private Sub cmdBalLft_Click()
    'balance sound to left only
    clDShow.DS_SetBalance -10000
End Sub

Private Sub cmdBalRgt_Click()
    'balance sound to right only
    clDShow.DS_SetBalance 10000
End Sub

Private Sub cmdRateSub_Click()
    'lower the playback rate
    Call clDShow.DS_SetRate(clDShow.DS_GetRate - 0.1)
End Sub

Private Sub cmdRateAdd_Click()
    'raise the playback rate
    Call clDShow.DS_SetRate(clDShow.DS_GetRate + 0.1)
End Sub

Private Sub cmdSeekBack_Click()
    'Seek backwards
    clDShow.DS_Seek -1
End Sub

Private Sub cmdSeekFor_Click()
    'Seek backwards
    clDShow.DS_Seek 1
End Sub
Public Sub CorrectAspect()

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
AreaWid = TargetWidth.ScaleWidth
AreaHgt = TargetHeight.ScaleHeight
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

VideoWindow.SetWindowPosition CenterLeft, CenterTop, w, h


'    ret = mciSendString("put " & MediaPath & " window at " & CenterLeft & " " & CenterTop & " " & _
'                            w & " " & h, 0&, 0&, 0&)

End Sub

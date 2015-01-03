VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form TimerMainWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Countdown Timer"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   FillColor       =   &H00C0C0C0&
   Icon            =   "Timer(source).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox LoopTmr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4245
      TabIndex        =   30
      Top             =   1725
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   195
      TabIndex        =   29
      Top             =   2430
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   5000
      Scrolling       =   1
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "Countdown"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      Caption         =   "Target"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   615
      Width           =   1215
   End
   Begin VB.CommandButton StartStop 
      Caption         =   "Start"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox UsrCountDown 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox UsrCountDown 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   17
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox UsrCountDown 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox TxtLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3120
      MousePointer    =   1  'Arrow
      TabIndex        =   15
      Text            =   "(yyyy)"
      Top             =   900
      Width           =   405
   End
   Begin VB.TextBox TxtLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   4200
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Text            =   "(dd)"
      Top             =   900
      Width           =   285
   End
   Begin VB.TextBox TxtLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   3660
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      Text            =   "(MM)"
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox UsrTarget 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox UsrTarget 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox UsrTarget 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1800
   End
   Begin VB.TextBox TxtLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   2505
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Text            =   "(ss)"
      Top             =   900
      Width           =   285
   End
   Begin VB.TextBox TxtLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1995
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Text            =   "(mm)"
      Top             =   900
      Width           =   345
   End
   Begin VB.TextBox TxtLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   1485
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Text            =   "(hh)"
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox UsrTarget 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox UsrTarget 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox UsrTarget 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Cover 
      BackStyle       =   0  'Transparent
      Height          =   945
      Left            =   120
      TabIndex        =   27
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label LoopLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Loop"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3870
      TabIndex        =   31
      Top             =   1740
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Source 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Since: N/A"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      ToolTipText     =   "Timestamp when countdown started"
      Top             =   1350
      Width           =   1530
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Date"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   150
      Width           =   1455
   End
   Begin VB.Label DispUnits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "sec"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   3060
      TabIndex        =   26
      Top             =   2235
      Width           =   495
   End
   Begin VB.Label DispUnits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "min"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   2265
      TabIndex        =   25
      Top             =   2235
      Width           =   495
   End
   Begin VB.Label DispUnits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "hours"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   1020
      TabIndex        =   24
      Top             =   2235
      Width           =   975
   End
   Begin VB.Label PosNeg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   645
      TabIndex        =   23
      Top             =   1680
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   4560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label CurrentTm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Top             =   105
      Width           =   1455
   End
   Begin VB.Label CurrentTm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   105
      Width           =   1575
   End
   Begin VB.Label VersionLbl 
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2655
      Width           =   2415
   End
   Begin VB.Label Display1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000:00:00"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   4455
   End
End
Attribute VB_Name = "TimerMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TimerYes As Boolean 'Indicates if timer is currently running
Dim StrCaption As String 'Holds the default caption string
Dim StrDefDisp As String 'Holds the default display string
Dim TargetDate As Date
Dim HMS(3) As Long 'Holds total seconds and Hours, Min, Sec
Dim MaxDays As Double 'Holds the countdown limit
Dim SoundPath(2) As String

Private Sub Form_Load()

TimerYes = False
LoopTmr.Value = 0
Call ButtonCaption
Call CenterFormOnScreen

StrCaption = "Countdown Timer"
VersionLbl = "Thomas Anag. 20090316"
StrDefDisp = "000:00:00"
Call cover_DblClick 'reset display
Call SetSourceDate

MaxDays = 999 / 24 + 59 / 1440 + 59 / 86400 'set maximum countdown limit (in days)


SoundPath(0) = App.Path & "\timer_01.wav"
SoundPath(1) = App.Path & "\timer_02.wav"
SoundPath(2) = App.Path & "\timer_03.wav"


CurrentTm(0).Caption = Format(Date, "yyyy/MM/dd")
CurrentTm(1).Caption = Format(CStr(Time), "hh:mm:ss")

'Populate the user fields
UsrTarget(0).Text = Hour(Time)
UsrTarget(1).Text = Minute(Time)
UsrTarget(2).Text = Second(Time)
UsrTarget(3).Text = Month(Date)
UsrTarget(4).Text = Day(Date)
UsrTarget(5).Text = Year(Date)
UsrCountDown(0).Text = "000"
UsrCountDown(1).Text = "00"
UsrCountDown(2).Text = "00"

Call FormatFields
Call SetDefaultColors
Call CalcTargetDate

Timer1.Interval = 1000 'set timer at 1 sec intervals (1000* 1/1000sec)

End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0:
            LoopLbl.Visible = False
            LoopTmr.Visible = False
            Option1(0).Value = True
            Option1(1).Value = False
        Case 1:
            LoopLbl.Visible = True
            LoopTmr.Visible = True
            Option1(0).Value = False
            Option1(1).Value = True
    End Select
End Sub

'The button
Private Sub startstop_Click()
    If (TimerYes = True) Then
        TimerYes = False
    Else
        TimerYes = True
        If (Option1(0).Value = True) Then Call UpdateCdown Else Call UpdateTarget
    End If
    Call ButtonCaption
    Call SetDefaultColors
    Call SetSourceDate
    Call Timer1_Timer
End Sub

'Button Caption toggle
Sub ButtonCaption()
    If (TimerYes = True) Then StartStop.Caption = "Stop" Else StartStop.Caption = "Start"
End Sub

Sub SetSourceDate()
    If (TimerYes = True) Then
        Source.Tag = Now
        Source.Caption = Format(CStr(Source.Tag), "hh:mm:ss") & " - " & Format(Date, "yy/MM/dd")
    Else
        Source.Tag = 0
        Source.Caption = "Since: N/A"
    End If
End Sub

'The default formating for all user fields
Sub FormatFields()
    UsrTarget(0).Text = Format(Val(UsrTarget(0).Text), "00")
    UsrTarget(1).Text = Format(Val(UsrTarget(1).Text), "00")
    UsrTarget(2).Text = Format(Val(UsrTarget(2).Text), "00")
    UsrTarget(3).Text = Format(Val(UsrTarget(3).Text), "00")
    UsrTarget(4).Text = Format(Val(UsrTarget(4).Text), "00")
    UsrTarget(5).Text = Format(Val(UsrTarget(5).Text), "0000")
    
    UsrCountDown(0).Text = Format(Val(UsrCountDown(0).Text), "000")
    UsrCountDown(1).Text = Format(Val(UsrCountDown(1).Text), "00")
    UsrCountDown(2).Text = Format(Val(UsrCountDown(2).Text), "00")
End Sub

'The units under the target fields (click to reset the field to NOW
Private Sub TxtLabel_Click(Index As Integer)
Select Case Index
Case 0:
    UsrTarget(Index).Text = Hour(Time)
Case 1:
    UsrTarget(Index).Text = Minute(Time)
Case 2:
    UsrTarget(Index).Text = Second(Time)
Case 3:
    UsrTarget(Index).Text = Month(Date)
Case 4:
    UsrTarget(Index).Text = Day(Date)
Case 5:
    UsrTarget(Index).Text = Year(Date)
End Select
Call UsrTarget_lostfocus(Index)
UsrTarget(Index).SetFocus
End Sub

'Radio button toggle
Private Sub UsrTarget_GotFocus(Index As Integer)
    Call Option1_Click(0)
End Sub

'Radio button toggle
Private Sub UsrCountDown_GotFocus(Index As Integer)
    Call Option1_Click(1)
End Sub

'Value filtering/verification
Private Sub UsrCountDown_LostFocus(Index As Integer)
Select Case Index
    Case 0:
        If (Val(UsrCountDown(Index)) < 0) Then UsrCountDown(Index) = 0
        If (Val(UsrCountDown(Index)) > 999) Then UsrCountDown(Index) = 999
    Case 1:
        If (Val(UsrCountDown(Index)) < 0) Then UsrCountDown(Index) = 0
        If (Val(UsrCountDown(Index)) > 59) Then UsrCountDown(Index) = 59
    Case 2:
        If (Val(UsrCountDown(Index)) < 0) Then UsrCountDown(Index) = 0
        If (Val(UsrCountDown(Index)) > 59) Then UsrCountDown(Index) = 59
End Select
Call UpdateTarget
Call SetSourceDate
End Sub

'Value filtering/verification
Private Sub UsrTarget_lostfocus(Index As Integer)
Dim InstantaneousDiff As Double
Select Case Index
Case 0:
    If ((Val(UsrTarget(Index).Text) > 24)) Then UsrTarget(Index).Text = "23"
    If (Val(UsrTarget(Index).Text) < 0 Or UsrTarget(Index).Text = "") Then UsrTarget(Index).Text = "00"
    UsrTarget(Index).Text = Val(UsrTarget(Index).Text)
Case 1:
    If ((Val(UsrTarget(Index).Text) > 59)) Then UsrTarget(Index).Text = "59"
    If (Val(UsrTarget(Index).Text) < 0 Or UsrTarget(Index).Text = "") Then UsrTarget(Index).Text = "00"
    UsrTarget(Index).Text = Val(UsrTarget(Index).Text)
Case 2:
    If ((Val(UsrTarget(Index).Text) > 59)) Then UsrTarget(Index).Text = "59"
    If (Val(UsrTarget(Index).Text) < 0 Or UsrTarget(Index).Text = "") Then UsrTarget(Index).Text = "00"
    UsrTarget(Index).Text = Val(UsrTarget(Index).Text)
Case 3:
    If ((Val(UsrTarget(Index).Text) > 12)) Then UsrTarget(Index).Text = "12"
    If (Val(UsrTarget(Index).Text) < 1 Or UsrTarget(Index).Text = "") Then UsrTarget(Index).Text = "01"
    UsrTarget(Index).Text = Val(UsrTarget(Index).Text)
Case 4:
    If (IsDate(UsrTarget(3) & "/" & UsrTarget(4) & "/" & UsrTarget(5)) = False) Then UsrTarget(Index).Text = Month(Date)
    UsrTarget(Index).Text = Val(UsrTarget(Index).Text)
Case 5:
    If (Val(UsrTarget(Index).Text) < 0 Or (UsrTarget(Index).Text = "")) Then UsrTarget(Index).Text = Year(Date)
    UsrTarget(Index).Text = Val(UsrTarget(Index).Text)
End Select
Call CalcTargetDate
If ((TargetDate - Now) > MaxDays) Then
    UsrTarget(0).Text = Hour(Now + MaxDays)
    UsrTarget(1).Text = Minute(Now + MaxDays)
    UsrTarget(2).Text = Second(Now + MaxDays)
    UsrTarget(3).Text = Month(Now + MaxDays)
    UsrTarget(4).Text = Day(Now + MaxDays)
    UsrTarget(5).Text = Year(Now + MaxDays)
    Call CalcTargetDate
End If
Call UpdateCdown
Call SetSourceDate
End Sub

Sub CalcTargetDate()
    TargetDate = CDate(UsrTarget(3) & "/" & UsrTarget(4) & "/" & UsrTarget(5) & " " & UsrTarget(0) & ":" & UsrTarget(1) & ":" & UsrTarget(2))
End Sub

'Use countdown values to update the "target date"
Sub UpdateTarget()
    TargetDate = Now + Val(UsrCountDown(0).Text) / 24 + Val(UsrCountDown(1).Text) / (1440) + Val(UsrCountDown(2).Text) / (86400)
    UsrTarget(0).Text = Hour(TargetDate)
    UsrTarget(1).Text = Minute(TargetDate)
    UsrTarget(2).Text = Second(TargetDate)
    UsrTarget(3).Text = Month(TargetDate)
    UsrTarget(4).Text = Day(TargetDate)
    UsrTarget(5).Text = Year(TargetDate)
    Call FormatFields
End Sub

'Use the target values to update the countdown fields
Sub UpdateCdown()
    Call calcHMS
    If (HMS(0) > 0) Then
        UsrCountDown(0).Text = HMS(1)
        UsrCountDown(1).Text = HMS(2)
        UsrCountDown(2).Text = HMS(3)
    Else
        UsrCountDown(0).Text = "000"
        UsrCountDown(1).Text = "00"
        UsrCountDown(2).Text = "00"
    End If
    Call FormatFields
End Sub

'Calculate Hour, Min, Sec
Sub calcHMS()
    Dim seconds As Long
    Dim Hleft As Long
    Dim Mleft As Long
    Dim Sleft As Long

    seconds = DateDiff("s", Now, TargetDate)
    Hleft = Int(Abs(seconds) / 3600)
    Mleft = Int((Abs(seconds) - Hleft * 3600) / 60)
    Sleft = Abs(seconds) - (Hleft * 3600 + Mleft * 60)
    
    HMS(0) = seconds
    HMS(1) = Hleft
    HMS(2) = Mleft
    HMS(3) = Sleft
End Sub

'This is the CORE
Private Sub Timer1_Timer()
Dim countdown As String
Dim SoundResult
Dim Progress As Double

CurrentTm(0).Caption = Format(Date, "yyyy/MM/dd")
CurrentTm(1).Caption = Format(CStr(Time), "hh:mm:ss")

If (TimerYes = True) Then
    Call calcHMS
    If (HMS(0) < 0) Then PosNeg.Caption = "-" Else PosNeg.Caption = ""
    countdown = Format(HMS(1), "000") & ":" & Format(HMS(2), "00") & ":" & Format(HMS(3), "00")
    Display1.Caption = countdown
    If (DateDiff("s", CDate(Source.Tag), TargetDate) = 0) Then Progress = 1 Else Progress = (DateDiff("s", CDate(Source.Tag), TargetDate) - DateDiff("s", Now, TargetDate)) / DateDiff("s", CDate(Source.Tag), TargetDate)
    If (Progress < 1 And Progress >= 0) Then ProgressBar1.Value = Int(ProgressBar1.Max * Progress) Else ProgressBar1.Value = ProgressBar1.Max
    
    If (HMS(0) < 10 And HMS(0) > 3 And Dir(SoundPath(0)) <> "") Then SoundResult = sndPlaySound(SoundPath(0), 1)
    If (HMS(0) <= 3 And HMS(0) > 0 And Dir(SoundPath(1)) <> "") Then SoundResult = sndPlaySound(SoundPath(1), 1)
    If (HMS(0) = 0 And Dir(SoundPath(2)) <> "") Then SoundResult = sndPlaySound(SoundPath(2), 1)
    If (HMS(0) = 0 And LoopTmr.Visible = True And LoopTmr.Value = 1) Then
        Call UpdateTarget
        Call SetSourceDate
        Call Timer1_Timer
    End If
End If
If (TimerMainWindow.WindowState = vbNormal) Then
    TimerMainWindow.Caption = StrCaption
Else
    TimerMainWindow.Caption = PosNeg.Caption & Display1.Caption
End If
End Sub

Private Sub CenterFormOnScreen()
    TimerMainWindow.Left = (Screen.Width - TimerMainWindow.Width) / 2
    TimerMainWindow.Top = (Screen.Height - TimerMainWindow.Height) / 2
End Sub

'Colors for all relevent elements
Private Sub SetDefaultColors()
    Dim count As Integer
    For count = 0 To 5: TxtLabel(count).ForeColor = &HC0C0C0: Next
    If (TimerYes = True) Then
        Display1.ForeColor = RGB(0, 0, 0)
        PosNeg.ForeColor = RGB(0, 0, 0)
    Else
        Display1.ForeColor = RGB(192, 192, 192)
        PosNeg.ForeColor = RGB(192, 192, 192)
    End If
End Sub

'Reset display
Private Sub cover_DblClick()
    If (TimerYes = False) Then Display1.Caption = StrDefDisp: PosNeg.Caption = "": ProgressBar1.Value = ProgressBar1.Min
End Sub

'Set the colors back to default
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetDefaultColors
End Sub

'Rollover color change for the display
Private Sub UsrTarget_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (TimerYes = True) Then Display1.ForeColor = RGB(255, 0, 0): PosNeg.ForeColor = RGB(255, 0, 0)
End Sub

'Rollover color change for the display
Private Sub UsrCountDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (TimerYes = True) Then Display1.ForeColor = RGB(255, 0, 0): PosNeg.ForeColor = RGB(255, 0, 0)
End Sub

'Clicking anywhere other than the text fields should cause validation (they lose focus)
Private Sub cover_Click()
    Call Form_Click
End Sub

'Induce a "lost focus" on the user fields (in order to force verification of values) by focusing on the button
Private Sub Form_Click()
    StartStop.SetFocus
End Sub

'Rollover color change for the text labels
Private Sub TxtLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtLabel(Index).ForeColor = RGB(255, 0, 0) 'red
End Sub

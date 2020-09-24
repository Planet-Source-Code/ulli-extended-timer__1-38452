VERSION 5.00
Begin VB.UserControl ExtendedTimer 
   CanGetFocus     =   0   'False
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   840
   ScaleWidth      =   450
   ToolboxBitmap   =   "ExtendedTimer.ctx":0000
   Begin VB.Timer myTimer 
      Left            =   0
      Top             =   420
   End
   Begin VB.Image imgWatch 
      Height          =   420
      Left            =   0
      Picture         =   "ExtendedTimer.ctx":00FA
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "ExtendedTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Properties
Private myInterval         As Long
Private myMinutes          As Long
Private myHours            As Long
Private myDays             As Long
Private myWeeks            As Long
Private myMode             As Mode
'Note:  Enabled property is hidden in the internal standard timer control

'Mode-Enum
Public Enum Mode
    'Due to a quirk in VB the case (upper / lower) of members in an Enum is modified if you type
    'the member name differently in code. There's a little trick how to restore the correct case:
    'type a comment (best inside the Enum) with the correct spelling like this:

    'Continuous , Oneshot

    'and if VB has mixed up things then simply un-comment and re-comment this. Be sure to move
    'the cursor out of the line after un-commenting to give VB a chance to notice the change.

    Continuous = 1
    Oneshot = 2
End Enum

'Variables
Private TotalInterval      As Currency 'the computed total interval
Private StillToGo          As Currency 'the time still to go
Private Tmp                As Long     'temporary property storage

'Property names
Private Const pnEnabled    As String = "Enabled"
Private Const pnInterval   As String = "Interval"
Private Const pnMinutes    As String = "Minutes"
Private Const pnHours      As String = "Hours"
Private Const pnDays       As String = "Days"
Private Const pnWeeks      As String = "Weeks"
Private Const pnMode       As String = "Mode"

'Durations (in seconds)
Private Const durMinutes   As Long = 60
Private Const durHours     As Long = durMinutes * 60
Private Const durDays      As Long = durHours * 24
Private Const durWeeks     As Long = durDays * 7

'Misc
Private Const MinTimerInt  As Long = 10 'millisecs (anything less is rouded up to 10 anyway)
Private Const MaxTimerInt  As Long = ((2 ^ 16 - 1) \ 10) * 10 'millisecs
Private Const MaxMilliSecs As Long = 59999 'millisecs
Private Const MaxInt       As Long = 52 'weeks
Private Const ErrInvalid   As Long = 380 'error number

'Event
Public Event Timer()

Public Property Get Copyright() As String

    Copyright = "© 2002 UMGEDV GmbH"

End Property

Public Property Let Copyright(nwCopyright As String)

  'do nothing

End Property

Public Property Get Days() As Long
Attribute Days.VB_Description = "Sets/returns the number of days between ticks."

    Days = myDays

End Property

Public Property Let Days(ByVal nwDays As Long)

    Tmp = myDays 'save original value
    myDays = nwDays
    If IsLegalInterval Then
        PropertyChanged pnDays
      Else 'ISLEGALINTERVAL = FALSE/0
        myDays = Tmp 'restore original value
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText(pnDays, "0 thru 6")
    End If

End Property

Public Property Get Enabled() As Boolean

    Enabled = myTimer.Enabled

End Property

Public Property Let Enabled(ByVal nwEnabled As Boolean)

    myTimer.Enabled() = nwEnabled
    If myTimer.Enabled = False Then
        StillToGo = TotalInterval 'reset interval so it starts anew when enabled again
    End If
    PropertyChanged pnEnabled

End Property

Private Function ErrText(Prop As String, Valid As String) As String

    ErrText = "The value assigned to property '" & Prop & "' was invalid; valid is " & Valid

End Function

Public Property Let Hours(ByVal nwHours As Long)
Attribute Hours.VB_Description = "Sets/returns the number of hours between ticks."

    Tmp = myHours 'save original value
    myHours = nwHours
    If IsLegalInterval Then
        PropertyChanged pnHours
      Else 'ISLEGALINTERVAL = FALSE/0
        myHours = Tmp 'restore original value
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText(pnHours, "0 thru 23")
    End If

End Property

Public Property Get Hours() As Long

    Hours = myHours

End Property

Public Property Let Interval(ByVal nwInterval As Double)
Attribute Interval.VB_Description = "Sets/returns the number of milliseconds between ticks."
Attribute Interval.VB_UserMemId = 0

    Tmp = myInterval 'save original value
    myInterval = nwInterval
    If myInterval Mod 10 <> 0 Then
        myInterval = (myInterval \ 10 + 1) * 10 'round up
    End If
    If IsLegalInterval Then
        PropertyChanged pnInterval
      Else 'ISLEGALINTERVAL = FALSE/0
        myInterval = Tmp 'restore original value
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText(pnInterval, MinTimerInt & " thru " & MaxTimerInt)
    End If

End Property

Public Property Get Interval() As Double

    Interval = myInterval

End Property

Private Function IsLegalInterval() As Boolean

    IsLegalInterval = (myInterval <= MaxTimerInt And myMinutes < 60 And myHours < 24 And myDays < 7 And myWeeks <= MaxInt _
                      And myInterval >= 0 And myMinutes >= 0 And myHours >= 0 And myDays >= 0 And myWeeks >= 0)
    If IsLegalInterval Then
        TotalInterval = myMinutes * durMinutes + myHours * durHours + myDays * durDays + myWeeks * durWeeks
        If TotalInterval And myInterval > MaxMilliSecs Then 'uses M, H, D, W
            myInterval = MaxMilliSecs 'so adjust the milliseconds
        End If
        TotalInterval = TotalInterval * 1000 + myInterval
        If TotalInterval < MinTimerInt Then
            If TotalInterval Then
                IsLegalInterval = False
            End If
        End If
        Enabled = False 'reset internal timer
        If TotalInterval > MaxTimerInt Then 'we'll eat the time in chunks
            myTimer.Interval = MaxTimerInt
          Else 'NOT TOTALINTERVAL...
            On Error Resume Next 'dont want the internal timer to raise an error; we'll do that ourselves
                myTimer.Interval = TotalInterval
            On Error GoTo 0
        End If
        Enabled = True 'restart internal timer
    End If

End Function

Public Property Get MilliSecs() As Long
Attribute MilliSecs.VB_Description = "Sets/returns the number of milliseconds between ticks."

    MilliSecs = myInterval

End Property

Public Property Let MilliSecs(ByVal nwMilliSecs As Long)

    If nwMilliSecs >= 0 And nwMilliSecs <= MaxMilliSecs Then
        Interval = nwMilliSecs 'use Interval property for this
      Else 'NOT NWMILLISECS...
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText("MilliSecs", "0 thru " & MaxMilliSecs)
    End If

End Property

Public Property Get Minutes() As Long
Attribute Minutes.VB_Description = "Sets/returns the number of minutes between ticks."

    Minutes = myMinutes

End Property

Public Property Let Minutes(ByVal nwMinutes As Long)

    Tmp = myMinutes 'save original value
    myMinutes = nwMinutes
    If IsLegalInterval Then
        PropertyChanged pnMinutes
      Else 'ISLEGALINTERVAL = FALSE/0
        myMinutes = Tmp 'restore original value
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText(pnMinutes, "0 thru 59")
    End If

End Property

Public Property Get Mode() As Mode
Attribute Mode.VB_Description = "Sets/returns the mode."

    Mode = myMode

End Property

Public Property Let Mode(nwMode As Mode)

    If nwMode = Continuous Or nwMode = Oneshot Then
        myMode = nwMode
        PropertyChanged pnMode
      Else 'NOT NWMODE...
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText(pnMode, "[Continuous] or [Oneshot]")
    End If

End Property

Private Sub myTimer_Timer()

  'internal timer ticks

    If Ambient.UserMode Then 'we're in run mode
        StillToGo = StillToGo - myTimer.Interval
        Select Case StillToGo
          Case Is < MinTimerInt
            RaiseEvent Timer
            If myMode = Oneshot Then
                Enabled = False
            End If
            StillToGo = TotalInterval
          Case Is < MaxTimerInt 'last chunk of time
            myTimer.Interval = StillToGo
        End Select
    End If

End Sub

Private Sub UserControl_InitProperties()

    myInterval = 0
    myMinutes = 0
    myHours = 0
    myDays = 0
    myWeeks = 0
    myMode = Continuous

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        myTimer.Enabled = .ReadProperty(pnEnabled, True)
        myMode = .ReadProperty(pnMode, Continuous)
        Interval = .ReadProperty(pnInterval, 0)
        Minutes = .ReadProperty(pnMinutes, 0)
        Hours = .ReadProperty(pnHours, 0)
        Days = .ReadProperty(pnDays, 0)
        Weeks = .ReadProperty(pnWeeks, 0)
    End With 'PROPBAG
    StillToGo = TotalInterval

End Sub

Private Sub UserControl_Resize()

    Size imgWatch.Width, imgWatch.Height 'no resizeing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty pnEnabled, myTimer.Enabled, True
        .WriteProperty pnMode, myMode, Continuous
        .WriteProperty pnInterval, myInterval, 0
        .WriteProperty pnMinutes, myMinutes, 0
        .WriteProperty pnHours, myHours, 0
        .WriteProperty pnDays, myDays, 0
        .WriteProperty pnWeeks, myWeeks, 0
    End With 'PROPBAG

End Sub

Public Property Get Weeks() As Long
Attribute Weeks.VB_Description = "Sets/returns the number of weeks between ticks."

    Weeks = myWeeks

End Property

Public Property Let Weeks(ByVal nwWeeks As Long)

    Tmp = myWeeks 'save original value
    myWeeks = nwWeeks
    If IsLegalInterval Then
        PropertyChanged pnWeeks
      Else 'ISLEGALINTERVAL = FALSE/0
        myWeeks = Tmp 'restore original value
        Err.Raise ErrInvalid, Ambient.DisplayName, ErrText(pnWeeks, "0 thru " & MaxInt)
    End If

End Property

':) Ulli's VB Code Formatter V2.14.7 (29.08.2002 16:27:41) 55 + 261 = 316 Lines

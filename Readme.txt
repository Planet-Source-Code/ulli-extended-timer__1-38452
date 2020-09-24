This is an extended timer OCX control permitting Intervals between 10 millisecs and 52 weeks. Although it has quite a few more properties it is still a direct replacement for the standard VB timer control. Check it out, download is only 5kB. Also, it's a good example for beginners who want to learn how to make OCXes.

PS
The longer intervals have not been tested because I found that a bit boring... *smile*

...and a note on intervals: this control rounds intervals up to the next multiple of 10 because that's apparenty what the standard timer control does (but without telling us)

You can test your Timer resolution with this little proggie (needs a timer and a command button)

'-------------------------------------------------------------------------------
Option Explicit

Private Cnt     As Long
Private Start   As Single

Private Sub Command1_Click()

  Dim i As Long

    Do
        i = InputBox("Interval in millisecs:")
    Loop Until i > 0 And i < 210
    Timer1.Interval = i
    Cnt = 0
    Timer1.Enabled = True
    Screen.MousePointer = vbHourglass
    Start = Timer

End Sub

Private Sub Form_Load()

    Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()

    Cnt = Cnt + 1
    If Cnt = 100 Then
        Print "Expected "; 0.1 * Timer1.Interval; "- actual "; Timer - Start
        Timer1.Enabled = False
        Screen.MousePointer = vbDefault
    End If

End Sub
'-------------------------------------------------------------------------------

Input various values and see for yourself...


'PowerPoint required Settings:
'1, File/Options/Trust center/Trust center Settings/Enable all macros.
'2, develop tools/controls/tags, in the presentation default SlideMaster.Label1.
'3, Development tools /Visual Basic, add "Module1", paste the following code.

Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long,  ByVal lpTimerFunc As LongPtr) As Long
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public mTimer, ID As Long

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim hms As String
    hms = Format(CDate(Now()), "hh:mm:ss")
    SlideMaster.Label1.Caption = hms
    SlideMaster.Label1.BackColor = RGB (198, 228, 238) 'set the background color of the label
    SlideMaster.Label1.BorderColor = RGB (198, 228, 238) 'set the border color label
End Sub

Public Sub OnSlideShowPageChange() 'When the powerpoint presentation starts, the TimerProc function is called to start displaying the timing
    If ID <= 0 Then ID = SetTimer(win_hwnd, 1, 1000, AddressOf TimerProc) 'Timing function, run every second
End Sub

Public Sub OnSlideShowTerminate() 'The timer ends when PPT terminates the presentation
    'On Error Resume Next
    mTimer = KillTimer(0, ID)
    ID = 0
    ActivePresentation.Saved = msoTrue 'prevents the PPT save prompt from appearing
End Sub
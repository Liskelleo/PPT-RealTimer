'PowerPoint��Ҫ�����ã�
'          1���ļ�/ѡ��/��������/������������/�������кꡣ
'          2����������/�ؼ�/��ǩ������ʾ�ĸ���Ԥ��SlideMaster.Label1��
'          3����������/Visual Basic�����ģ��1��ճ�����´��롣

Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public mTimer, ID As Long
 
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim hms As String
    hms = Format(CDate(Now()), "hh:mm:ss")
    SlideMaster.Label1.Caption = hms
    SlideMaster.Label1.BackColor = RGB(198, 228, 238) '���ñ�ǩ������ɫ
    SlideMaster.Label1.BorderColor = RGB(198, 228, 238) '���ñ�ǩ�߿���ɫ
End Sub

Public Sub OnSlideShowPageChange() 'PPT��ʼչʾʱ������TimerProc��������ʼ��ʾ��ʱ
    If ID <= 0 Then ID = SetTimer(win_hwnd, 1, 1000, AddressOf TimerProc) '��ʱ������ÿһ��������һ��
End Sub
 
Public Sub OnSlideShowTerminate() 'PPT��ֹ��ʾʱ��������ʱ
    'On Error Resume Next
    mTimer = KillTimer(0, ID)
    ID = 0
    ActivePresentation.Saved = msoTrue '����PPT������ʾ����
End Sub

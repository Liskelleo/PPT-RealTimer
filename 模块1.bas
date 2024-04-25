'PowerPoint需要的设置：
'          1、文件/选项/信任中心/信任中心设置/启用所有宏。
'          2、开发工具/控件/标签，在演示文稿中预设SlideMaster.Label1。
'          3、开发工具/Visual Basic，添加模块1，粘贴如下代码。

Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public mTimer, ID As Long
 
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim hms As String
    hms = Format(CDate(Now()), "hh:mm:ss")
    SlideMaster.Label1.Caption = hms
    SlideMaster.Label1.BackColor = RGB(198, 228, 238) '设置标签背景颜色
    SlideMaster.Label1.BorderColor = RGB(198, 228, 238) '设置标签边框颜色
End Sub

Public Sub OnSlideShowPageChange() 'PPT开始展示时，调用TimerProc函数，开始显示计时
    If ID <= 0 Then ID = SetTimer(win_hwnd, 1, 1000, AddressOf TimerProc) '计时函数，每一秒钟运行一次
End Sub
 
Public Sub OnSlideShowTerminate() 'PPT终止演示时，结束计时
    'On Error Resume Next
    mTimer = KillTimer(0, ID)
    ID = 0
    ActivePresentation.Saved = msoTrue '避免PPT保存提示出现
End Sub

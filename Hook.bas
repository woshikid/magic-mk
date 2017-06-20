Attribute VB_Name = "Hook"
Option Explicit
Public status  As Long '0 = standby, 1 = recording, 2 = playing, 3 = inputing
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Type tagKBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private Type tagPOINT
    x As Long
    y As Long
End Type
Private Type tagMSLLHOOKSTRUCT
    pt As tagPOINT
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private Const WH_KEYBOARD_LL = 13
Private Const WH_MOUSE_LL = 14
Private Const KEYEVENTF_KEYDOWN = 0
Private Const KEYEVENTF_KEYUP = 2
Private Const F10 = 121
Private Const F11 = 122
Private Const F12 = 123
Private Const WM_MOUSEMOVE = 512
Private Const WM_LBUTTONDOWN = 513
Private Const WM_LBUTTONUP = 514
Private Const WM_LBUTTONDBLCLK = 515
Private Const WM_RBUTTONDOWN = 516
Private Const WM_RBUTTONUP = 517
Private Const WM_RBUTTONDBLCLK = 518
Private Const WM_MBUTTONDOWN = 519
Private Const WM_MBUTTONUP = 520
Private Const WM_MBUTTONDBLCLK = 521
Private Const WM_MOUSEWHEEL = 522
Private Const WM_MOUSEHWHEEL = 526
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_WHEEL = &H800
Private Const MOUSEEVENTF_HWHEEL = &H1000
Private Const WHEEL_DELTA = 120
Private kHook As Long
Private mHook As Long
Private keyHooked As Boolean
Private mouseHooked As Boolean
Private kRecords() As Long
Private mRecords() As Long
Private maxK As Long
Private maxM As Long
Private posK As Long
Private posM As Long
Private repeat As Long
Private repeats As String
Public timeLine As Long
Private maxTime As Long

Public Sub EnableKeyHook()
    On Error Resume Next
    If keyHooked = True Then Exit Sub
    kHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyHook, App.hInstance, 0)
    keyHooked = True
End Sub

Public Sub UnHookKey()
    On Error Resume Next
    If keyHooked = False Then Exit Sub
    UnhookWindowsHookEx kHook
    keyHooked = False
End Sub

Public Function KeyHook(ByVal vkCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    KeyHook = 0
    Dim ks As tagKBDLLHOOKSTRUCT
    CopyMemory ks, ByVal lParam, LenB(ks)
    KeyEvent ks.vkCode, ks.flags
    If ks.vkCode = F10 Or ks.vkCode = F11 Or ks.vkCode = F12 Then KeyHook = 1
    If status = 2 And (ks.flags And 16) = 0 Then KeyHook = 1
    CallNextHookEx kHook, vkCode, wParam, ByVal lParam
End Function

Public Sub KeyEvent(ByVal vkCode As Long, ByVal flags As Long)
    On Error Resume Next
    Dim i, keyf As Long
    If (flags And 128) = 0 Then
        keyf = KEYEVENTF_KEYDOWN
    Else
        keyf = KEYEVENTF_KEYUP
    End If
    
    If status = 0 Then
        If keyf = KEYEVENTF_KEYUP Then
            If vkCode = F12 Then 'start to record
                EnableMouseHook
                ReDim kRecords(2, 1024)
                ReDim mRecords(3, 10240)
                maxK = 0
                maxM = 0
                repeat = 1
                timeLine = 0
                maxTime = 0
                status = 1
                Magic.mkTimer.Enabled = True
                StartBeep
            ElseIf vkCode = F11 Then 'start to play
                EnableMouseHook
                status = 2
                Magic.mkTimer.Enabled = True
                StartBeep
            ElseIf vkCode = F10 Then 'start to input
                repeats = ""
                status = 3
                StartBeep
            End If
        End If
    ElseIf status = 1 Then
        If vkCode = F12 And keyf = KEYEVENTF_KEYUP Then 'stop recording
            UnHookMouse
            status = 0
            Magic.mkTimer.Enabled = False
            maxTime = timeLine
            timeLine = 0
            posK = 0
            posM = 0
            EndBeep
        ElseIf vkCode <> F10 And vkCode <> F11 And vkCode <> F12 Then 'don't forget the F12 keydown
            i = UBound(kRecords, 2)
            If maxK > i Then ReDim Preserve kRecords(2, i + 1024)
            kRecords(0, maxK) = timeLine
            kRecords(1, maxK) = vkCode
            kRecords(2, maxK) = keyf
            maxK = maxK + 1
        End If
    ElseIf status = 2 Then
        If vkCode = F11 And keyf = KEYEVENTF_KEYUP Then 'stop playing
            UnHookMouse
            status = 0
            Magic.mkTimer.Enabled = False
            EndBeep
        End If
    ElseIf status = 3 Then
        If keyf = KEYEVENTF_KEYUP Then
            If vkCode = F10 Then 'stop inputing
                repeat = Val(repeats)
                status = 0
                timeLine = 0
                posK = 0
                posM = 0
                EndBeep
            Else
                repeats = repeats & Chr(vkCode)
            End If
        End If
    End If
End Sub

Public Sub EnableMouseHook()
    On Error Resume Next
    If mouseHooked = True Then Exit Sub
    mHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)
    mouseHooked = True
End Sub

Public Sub UnHookMouse()
    On Error Resume Next
    If mouseHooked = False Then Exit Sub
    UnhookWindowsHookEx mHook
    mouseHooked = False
End Sub

Public Function MouseHook(ByVal vkCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    MouseHook = 0
    Dim ms As tagMSLLHOOKSTRUCT
    CopyMemory ms, ByVal lParam, LenB(ms)
    MouseEvent wParam, ms.pt.x, ms.pt.y, ms.mouseData
    If status = 2 And (ms.flags And 1) = 0 Then MouseHook = 1
    CallNextHookEx mHook, vkCode, wParam, ByVal lParam
End Function

Public Sub MouseEvent(ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal data As Long)
    On Error Resume Next
    If status <> 1 Then Exit Sub
    Dim i As Long
    i = UBound(mRecords, 2)
    If maxM > i Then ReDim Preserve mRecords(3, i + 10240)
    mRecords(0, maxM) = timeLine
    mRecords(1, maxM) = x
    mRecords(2, maxM) = y
    mRecords(3, maxM) = wParam
    If (wParam = WM_MOUSEWHEEL Or wParam = WM_MOUSEHWHEEL) And data < 0 Then mRecords(3, maxM) = -wParam
    maxM = maxM + 1
End Sub

Public Sub PlayFrame()
    On Error Resume Next
    If timeLine > maxTime Then
        timeLine = 0
        posK = 0
        posM = 0
        repeat = repeat - 1
        If repeat = 0 Then
            UnHookMouse
            status = 0
            Magic.mkTimer.Enabled = False
            repeat = 1
            EndBeep
        ElseIf repeat < 0 Then
            repeat = 0
        End If
    Else
        Do While posK < maxK
            If kRecords(0, posK) > timeLine Then Exit Do
            keybd_event kRecords(1, posK), MapVirtualKey(kRecords(1, posK), 0), kRecords(2, posK), 0
            posK = posK + 1
        Loop
        
        Do While posM < maxM
            If mRecords(0, posM) > timeLine Then Exit Do
            SetCursorPos mRecords(1, posM), mRecords(2, posM)
            If mRecords(3, posM) <> WM_MOUSEMOVE Then
                Select Case mRecords(3, posM)
                    Case WM_LBUTTONDOWN
                        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
                    Case WM_LBUTTONUP
                        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
                    Case WM_RBUTTONDOWN
                        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
                    Case WM_RBUTTONUP
                        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
                    Case WM_MBUTTONDOWN
                        mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
                    Case WM_MBUTTONUP
                        mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
                    Case WM_MOUSEWHEEL
                        mouse_event MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA, 0
                    Case -WM_MOUSEWHEEL
                        mouse_event MOUSEEVENTF_WHEEL, 0, 0, -WHEEL_DELTA, 0
                    Case WM_MOUSEHWHEEL
                        mouse_event MOUSEEVENTF_HWHEEL, 0, 0, WHEEL_DELTA, 0
                    Case -WM_MOUSEHWHEEL
                        mouse_event MOUSEEVENTF_HWHEEL, 0, 0, -WHEEL_DELTA, 0
                End Select
            End If
            posM = posM + 1
        Loop
    End If
End Sub

Private Sub StartBeep()
    On Error Resume Next
    Beep 660, 120
End Sub

Private Sub EndBeep()
    On Error Resume Next
    Beep 490, 120
End Sub

Attribute VB_Name = "modNoEnd"
'==============================================================================
' NoEnd - PowerPoint add-in that stops a slide show from running past its end.
'
' Once the show is on its last slide AND every click animation on that slide
' has played, further "next" input (clicker, keyboard, mouse click, wheel) is
' swallowed, so the last slide simply stays up.
'
' Safety net: if anything still gets through (e.g. the Next arrow in Presenter
' View, or a slide timing), a watchdog bounces the show from the black
' "End of slide show" screen straight back to the last slide, fully built.
'
' Esc, right-click > End Show and Presenter View's "End Slide Show" still work,
' so the operator can always get out.
'==============================================================================
Option Explicit

'---- Settings ----------------------------------------------------------------
' Swallow "next" input once the end is reached (no flicker at all).
Private Const BLOCK_INPUT As Boolean = True
' Also swallow left clicks / taps on the slide show window once at the end.
' Set to False if your LAST slide has hyperlinks, action buttons or trigger
' animations - the watchdog below still keeps the show from ending.
Private Const BLOCK_MOUSE_CLICKS As Boolean = True
' Bounce back from the black end screen if something slipped through.
' Needs File > Options > Advanced > "End with black slide" turned on (default).
Private Const USE_WATCHDOG As Boolean = True
Private Const WATCHDOG_MS As Long = 40

'---- Win32 -------------------------------------------------------------------
Private Declare PtrSafe Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hWnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function GetClassNameW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Sub CopyFromPtr Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyToPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByRef Source As Any, ByVal Length As LongPtr)

Private Const WH_GETMESSAGE As Long = 3
Private Const PM_REMOVE As Long = 1
Private Const GA_ROOT As Long = 2

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_POINTERDOWN As Long = &H246
Private Const WM_POINTERUP As Long = &H247
Private Const WM_POINTERWHEEL As Long = &H24E

Private Const VK_RETURN As Long = &HD
Private Const VK_SPACE As Long = &H20
Private Const VK_NEXT As Long = &H22     ' Page Down (what most clickers send)
Private Const VK_RIGHT As Long = &H27
Private Const VK_DOWN As Long = &H28
Private Const VK_N As Long = &H4E

' Window classes
Private Const CLS_SHOW As String = "screenClass"      ' full-screen slide show
Private Const CLS_EDITOR As String = "PPTFrameClass"  ' normal editing window
Private Const CLS_DIALOG As String = "#32770"

' Offsets into the Win32 MSG structure (hwnd, message, wParam, ...)
#If Win64 Then
Private Const MSG_MESSAGE As Long = 8
Private Const MSG_WPARAM As Long = 16
#Else
Private Const MSG_MESSAGE As Long = 4
Private Const MSG_WPARAM As Long = 8
#End If

'---- State -------------------------------------------------------------------
Private gEvents As clsNoEndEvents
Private gHook As LongPtr
Private gTimer As LongPtr
Private gSwallowUp As Boolean      ' swallow the button-up matching a swallowed button-down
Private gTypingNumber As Boolean   ' "type slide number + Enter" in progress
Private gInHook As Boolean
Private gInTimer As Boolean

'==============================================================================
' Add-in lifetime
'==============================================================================
Public Sub Auto_Open()
    Set gEvents = New clsNoEndEvents
    Set gEvents.App = Application
    If Application.SlideShowWindows.Count > 0 Then NoEnd_StartGuard
End Sub

Public Sub Auto_Close()
    NoEnd_StopGuard
    Set gEvents = Nothing
End Sub

' Called on SlideShowBegin. Hook and timer live only while a show is running.
Public Sub NoEnd_StartGuard()
    If BLOCK_INPUT And gHook = 0 Then
        gHook = SetWindowsHookExW(WH_GETMESSAGE, AddressOf NoEnd_GetMsgProc, 0, GetCurrentThreadId())
    End If
    If gTimer = 0 Then
        gTimer = SetTimer(0, 0, WATCHDOG_MS, AddressOf NoEnd_TimerProc)
    End If
End Sub

Public Sub NoEnd_StopGuard()
    If gHook <> 0 Then UnhookWindowsHookEx gHook: gHook = 0
    If gTimer <> 0 Then KillTimer 0, gTimer: gTimer = 0
    gSwallowUp = False
    gTypingNumber = False
End Sub

'==============================================================================
' Input filter: sees every message PowerPoint's UI thread pulls from its queue.
' A message we want to swallow is rewritten to WM_NULL.
'==============================================================================
Public Function NoEnd_GetMsgProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If nCode >= 0 And wParam = PM_REMOVE And Not gInHook Then
        gInHook = True
        If ShouldSwallow(lParam) Then
            Dim nullMsg As Long
            CopyToPtr lParam + MSG_MESSAGE, nullMsg, 4
        End If
        gInHook = False
    End If
    NoEnd_GetMsgProc = CallNextHookEx(gHook, nCode, wParam, lParam)
End Function

Private Function ShouldSwallow(ByVal pMsg As LongPtr) As Boolean
    On Error GoTo Bail
    Dim msg As Long, wp As Long, hWnd As LongPtr
    CopyFromPtr msg, pMsg + MSG_MESSAGE, 4

    Select Case msg
    Case WM_KEYDOWN
        CopyFromPtr wp, pMsg + MSG_WPARAM, 4
        If IsNextKey(wp) Then
            CopyFromPtr hWnd, pMsg, LenB(hWnd)
            If IsShowKeyTarget(hWnd) Then ShouldSwallow = AnyShowAtEnd(False)
        End If

    Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK, WM_POINTERDOWN
        If BLOCK_MOUSE_CLICKS Then
            CopyFromPtr hWnd, pMsg, LenB(hWnd)
            If RootClassName(hWnd) = CLS_SHOW Then ShouldSwallow = AnyShowAtEnd(True)
        End If
        gSwallowUp = ShouldSwallow

    Case WM_LBUTTONUP, WM_POINTERUP
        ShouldSwallow = gSwallowUp
        gSwallowUp = False

    Case WM_MOUSEWHEEL, WM_POINTERWHEEL
        CopyFromPtr wp, pMsg + MSG_WPARAM, 4
        If wp < 0 Then  ' negative wheel delta (towards the user) = next
            CopyFromPtr hWnd, pMsg, LenB(hWnd)
            If RootClassName(hWnd) = CLS_SHOW Then ShouldSwallow = AnyShowAtEnd(False)
        End If
    End Select
Bail:
End Function

Private Function IsNextKey(ByVal vk As Long) As Boolean
    Select Case vk
    Case &H30 To &H39, &H60 To &H69
        ' Digits: "type a slide number + Enter" must keep working.
        gTypingNumber = True
        Exit Function
    Case VK_RETURN
        IsNextKey = Not gTypingNumber
    Case VK_SPACE, VK_NEXT, VK_RIGHT, VK_DOWN, VK_N
        IsNextKey = True
    End Select
    gTypingNumber = False
End Function

' Keys go to whichever PowerPoint window has focus: the slide show itself or
' Presenter View (where a clicker's keys land on a two-screen setup). Only the
' normal editing window and dialogs are left alone.
Private Function IsShowKeyTarget(ByVal hWnd As LongPtr) As Boolean
    Dim cls As String: cls = RootClassName(hWnd)
    IsShowKeyTarget = (cls <> "" And cls <> CLS_EDITOR And cls <> CLS_DIALOG And Left$(cls, 9) <> "NUIDialog")
End Function

Private Function RootClassName(ByVal hWnd As LongPtr) As String
    If hWnd = 0 Then Exit Function
    Dim root As LongPtr: root = GetAncestor(hWnd, GA_ROOT)
    If root = 0 Then root = hWnd
    Dim buf As String: buf = String$(256, vbNullChar)
    Dim n As Long: n = GetClassNameW(root, StrPtr(buf), 256)
    RootClassName = Left$(buf, n)
End Function

'==============================================================================
' Watchdog: also detects the end of the show so the hook can be removed.
'==============================================================================
Public Sub NoEnd_TimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    If gInTimer Then Exit Sub
    gInTimer = True
    On Error Resume Next

    Dim n As Long: n = Application.SlideShowWindows.Count
    If Err.Number <> 0 Then n = -1: Err.Clear

    If n = 0 Then
        NoEnd_StopGuard
    ElseIf n > 0 And USE_WATCHDOG Then
        Dim i As Long
        For i = 1 To n
            BounceIfPastEnd Application.SlideShowWindows(i)
        Next
    End If

    gInTimer = False
End Sub

Private Sub BounceIfPastEnd(ByVal w As SlideShowWindow)
    On Error GoTo Bail
    If w.View.State = ppSlideShowDone Then
        ' msoFalse = come back with the slide's animations already played.
        w.View.GotoSlide LastSlideIndex(w), msoFalse
    End If
Bail:
End Sub

'==============================================================================
' Where is the show?
'==============================================================================
Private Function AnyShowAtEnd(ByVal forMouse As Boolean) As Boolean
    On Error GoTo Bail
    Dim i As Long
    For i = 1 To Application.SlideShowWindows.Count
        If IsAtEnd(Application.SlideShowWindows(i), forMouse) Then
            AnyShowAtEnd = True
            Exit Function
        End If
    Next
Bail:
End Function

' True when the show is on its last slide and no click animations are left,
' i.e. the next "next" would end the show.
Private Function IsAtEnd(ByVal w As SlideShowWindow, ByVal forMouse As Boolean) As Boolean
    On Error GoTo Bail
    Dim v As SlideShowView: Set v = w.View

    Select Case v.State
    Case ppSlideShowRunning, ppSlideShowPaused
    Case Else: Exit Function  ' black/white screen, or already past the end
    End Select

    If forMouse Then
        ' Leave clicks alone while the presenter is drawing with the pen etc.
        Select Case v.PointerType
        Case ppSlideShowPointerArrow, ppSlideShowPointerAutoArrow, ppSlideShowPointerAlwaysHidden, ppSlideShowPointerNone
        Case Else: Exit Function
        End Select
    End If

    If v.Slide.SlideIndex <> LastSlideIndex(w) Then Exit Function
    IsAtEnd = (v.GetClickIndex >= v.GetClickCount)
Bail:
End Function

' Last slide this show will display: honours custom shows, "From slide X to Y"
' ranges and hidden slides.
Private Function LastSlideIndex(ByVal w As SlideShowWindow) As Long
    Dim pres As Presentation: Set pres = w.Presentation

    If w.View.IsNamedShow Then
        Dim ids As Variant
        ids = pres.SlideShowSettings.NamedSlideShows(w.View.SlideShowName).SlideIDs
        LastSlideIndex = pres.Slides.FindBySlideID(ids(UBound(ids))).SlideIndex
        Exit Function
    End If

    Dim firstIdx As Long, lastIdx As Long, i As Long
    firstIdx = 1: lastIdx = pres.Slides.Count
    With pres.SlideShowSettings
        If .RangeType = ppShowSlideRange Then firstIdx = .StartingSlide: lastIdx = .EndingSlide
    End With

    For i = lastIdx To firstIdx Step -1
        If pres.Slides(i).SlideShowTransition.Hidden = msoFalse Then
            LastSlideIndex = i
            Exit Function
        End If
    Next
    LastSlideIndex = lastIdx
End Function

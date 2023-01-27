Option Explicit

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long

Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Declare PtrSafe Function SetWindowsHookEx Lib _
"user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Type POINTAPI
  X As Long
  Y As Long
End Type

Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Const HC_ACTION = 0
Const WH_MOUSE_LL = 14
Const WM_MOUSEWHEEL = &H20A

Dim hhkLowLevelMouse, lngInitialColor As Long
Dim udtlParamStuct As MSLLHOOKSTRUCT
Public intTopIndex As Integer

Function GetHookStruct(ByVal lParam As Long) As MSLLHOOKSTRUCT

   CopyMemory VarPtr(udtlParamStuct), lParam, LenB(udtlParamStuct)

   GetHookStruct = udtlParamStuct

End Function

Function LowLevelMouseProc _
(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error Resume Next

    If (nCode = HC_ACTION) Then

        If wParam = WM_MOUSEWHEEL Then

            LowLevelMouseProc = True

            'ATENÇÃO: Troque o nome do seu Userform
            With Cadastro

                'ROLAR PARA CIMA
                If GetHookStruct(lParam).mouseData > 0 Then
                    .ScrollTop = intTopIndex - 10
                    intTopIndex = .ScrollTop
                Else
                'ROLAR PARA BAIXO
                    .ScrollTop = intTopIndex + 10
                    intTopIndex = .ScrollTop
                End If

            End With

        End If

        Exit Function

    End If

    UnhookWindowsHookEx hhkLowLevelMouse
    LowLevelMouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
End Function

Sub Hook_Mouse()
    If hhkLowLevelMouse <> 0 Then
        UnhookWindowsHookEx hhkLowLevelMouse
    End If

    hhkLowLevelMouse = SetWindowsHookEx _
    (WH_MOUSE_LL, AddressOf LowLevelMouseProc, Application.Hinstance, 0)

End Sub

Sub UnHook_Mouse()

    If hhkLowLevelMouse <> 0 Then UnhookWindowsHookEx hhkLowLevelMouse

End Sub






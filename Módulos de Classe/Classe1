Option Explicit

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_POPUP As Long = &H80000000
Private Const WS_VISIBLE As Long = &H10000000

Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_TOOLWINDOW As Long = &H80

Private Const SC_CLOSE As Long = &HF060

Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5
Private Const SW_MAXIMIZE As Long = 3


Private Const WM_SETICON = &H80

Dim hWndForm As Long, mbSizeable As Boolean, mbCaption As Boolean, mbIcon As Boolean, miModal As Integer
'Dim mbMaximize As Boolean
Dim mbMinimize As Boolean, mbSysMenu As Boolean, mbCloseBtn As Boolean
Dim mbAppWindow As Boolean, mbToolWindow As Boolean, msIconPath As String
Dim moForm As Object
Public Property Let Modal(bModal As Boolean)
    miModal = Abs(CInt(Not bModal))

    'Make the form modal or modeless by enabling/disabling Excel itself
    EnableWindow FindWindow("XLMAIN", Application.Caption), miModal
End Property

Public Property Get Modal() As Boolean
    Modal = (miModal <> 1)
End Property

Public Property Set Form(oForm As Object)

    If Val(Application.Version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", oForm.Caption)  'XL97
    Else
        hWndForm = FindWindow("ThunderDFrame", oForm.Caption)  'XL2000
    End If

    Set moForm = oForm

    AtualizarEstiloForm
    
    Dim strIconPath As String
    Dim lngIcon As Long
    Dim lnghWnd As Long
    strIconPath = "\\192.168.0.252\ti_share\2022\CTI\Logo.ico" 'Insira aqui o caminho completo do ícone - no formato .ICO resolução 32x32'
    lngIcon = ExtractIcon(0, strIconPath, 0)
    lnghWnd = FindWindow("ThunderDFrame", oForm.Caption)
    SendMessage lnghWnd, WM_SETICON, True, lngIcon
    SendMessage lnghWnd, WM_SETICON, False, lngIcon
    
End Property

Private Sub AtualizarEstiloForm()

    Dim iStyle As Long, hMenu As Long, hID As Long, iItems As Integer

    If hWndForm = 0 Then Exit Sub

    iStyle = GetWindowLong(hWndForm, GWL_STYLE)

    iStyle = iStyle Or WS_CAPTION
    iStyle = iStyle Or WS_SYSMENU
    'iStyle = iStyle Or WS_THICKFRAME
    iStyle = iStyle Or WS_MINIMIZEBOX
    iStyle = iStyle Or WS_MAXIMIZEBOX
    iStyle = iStyle And Not WS_VISIBLE And Not WS_POPUP

    SetWindowLong hWndForm, GWL_STYLE, iStyle

    iStyle = GetWindowLong(hWndForm, GWL_EXSTYLE)

    iStyle = iStyle And Not WS_EX_DLGMODALFRAME
    iStyle = iStyle Or WS_EX_APPWINDOW

    SetWindowLong hWndForm, GWL_EXSTYLE, iStyle

    hMenu = GetSystemMenu(hWndForm, 0)

    ShowWindow hWndForm, SW_MAXIMIZE 'Substitua SW_SHOW por SW_MAXIMIZE, para ter um tela maximizada no início
    DrawMenuBar hWndForm
    SetFocus hWndForm

End Sub


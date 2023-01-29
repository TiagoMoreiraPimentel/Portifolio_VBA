# Tela (forms) tela-cheia responsiva

# 1 - Crie um módulo de classe
## Crie um módulo de classe no seu projeto VBA e cole o código abaixo.

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
        strIconPath = ThisWorkbook.Path & "\exemplo.ico" 'Insira aqui o caminho completo do ícone - no formato .ICO resolução 32x32'
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

        ShowWindow hWndForm, SW_SHOW 'Substitua SW_SHOW por SW_MAXIMIZE, para ter um tela maximizada no início
        DrawMenuBar hWndForm
        SetFocus hWndForm

    End Sub

# 2 - Crie um módulo (normal)
## Crie um módulo normal no seu projeto e insira o código abaixo.

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
    x As Long
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
                With UserForm1

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

# 3 - Crie um formulário
## Pronto, agora em todos os formulários que você criar no seu projeto, insira o código abaixo. Cada um no seu respectivo procedimento.

    'Essa parte vai no topo do seu código, fora dos subs e procedimentos'
    Dim nAtualizaForm As New Classe1
    Dim T
    Dim frmUserWidth As Double
    Dim frmUserWidthRatio As Double
    Dim frmUserHeight As Double
    Dim frmUserHeightRatio As Double
    Dim r As Integer
    Dim c As Integer
    Dim ctl As Control

    'No procedimento Activate do Userform'
    Private Sub UserForm_Activate()
    Set nAtualizaForm.Form = Me
    End Sub

    'No evento INITIALIZE do Userform'
    Private Sub UserForm_Initialize()
    Dim hWnd As Long

        'Vai para o topo do formulário
        ScrollTop = 0

        'Define os botões minimizar e maximizar do form
        hWnd = FindWindow(vbNullString, Me.Caption)
        SetWindowLong hWnd, -16, &H20000 Or &H10000 Or &H84C80080
        
        frmUserWidth = Me.InsideWidth
        frmUserHeight = Me.InsideHeight
    end sub


    'Essa parte no evento Resize'
    Private Sub UserForm_Resize()

        If Me.InsideHeight < 1 Then Exit Sub
        
        frmUserWidthRatio = Me.InsideWidth / frmUserWidth
        frmUserHeightRatio = Me.InsideHeight / frmUserHeight
        
    ' Eliminate this section to prevent resizing of controls on form.
        ' Stick any control on the form at any location.
        For Each ctl In Me.Controls
            ctl.Width = frmUserWidthRatio * ctl.Width
            ctl.Left = frmUserWidthRatio * ctl.Left
            ctl.Height = frmUserHeightRatio * ctl.Height
            ctl.Top = frmUserHeightRatio * ctl.Top
        Next
        
        frmUserWidth = Me.InsideWidth
        frmUserHeight = Me.InsideHeight

    End Sub



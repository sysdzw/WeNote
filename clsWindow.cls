VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================
'名    称：windows窗体控制类v2.3
'描    述：一个操作windows窗口的类，可对窗口进行很多常用的操作(类名为clsWindow)
'使用范例：Dim window As New clsWindow
'         window.GetWindowByTitle("计算器").closeWindow' ***!!!win10如果异常请用管理员权限执行***!!!
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份，共同维护
'发布日期：2013/06/01
'博    客：https://blog.csdn.net/sysdzw
'用户手册：https://www.kancloud.cn/sysdzw/clswindow/
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                                             2012/12/03
'          V1.1 修正了几个正则相关的函数，调整了部分类结构                       2013/05/28
'          V1.2 增加属性Caption，可以获取或设置当前标题栏                        2013/05/29
'          V1.3 增加了方法Focus，可以激活当前窗口                                2013/06/01
'               增加了方法Left,Top,Width,Height,Move，处理窗口位置等
'          V1.4 增加了窗口位置调整的几个函数                                     2013/06/04
'               增加了得到应用程序路径的函数AppName
'               增加了得到应用程序启动参数的函数AppCommandLine
'          V1.5 增加了窗口最大最小化，隐藏显示正常的几个函数                     2013/06/06
'               增加了获取控件相关函数是否使用正则的参数UseRegExp默认F
'          V1.6 将Left，Top函数改为属性，可获得可设置                            2013/06/10
'          V1.7 增加函数：CloseApp 结束进程                                      2013/06/13
'               修正了部分跟正则匹配相关的函数
'               增加函数：GetElementTextByText
'               增加函数：GetElementHwndByText
'          V1.8 增加函数：GetWindowByClassName                                   2013/06/26
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByAppName
'               增加私有变量hWnd_
'               增加属性hWnd，可设置，单设置时候会检查，非法则设置为0
'               更新GetWindowByTitleEx函数，使之可以选择性支持正则
'               删除GetWindowByTitleRegExp函数，合并到上面函数
'               增加SetFocus函数，调用Focus实现，为了是兼容VB习惯
'               扩了ProcessID、AppPath、AppName、AppCommandLine三个函数，可带参数
'               网友wwb(wwbing@gmail.com)提供了一些函数和方法属性：
'                 CheckWindow, Load, WindowState, Visible, hDC, ZOrder
'                 AlphaBlend, Enabled, Refresh, TransparentColor
'               采纳wwb网友的部分意见，将句柄变量改为hWnd_，但是hWnd作为公共属性
'          V1.9 修正函数：GetMatchHwndFromWindow 正则表达式的错误                2013/08/07
'               修正函数：GetMatchHwndFromWindow 函数中的一些错误                2014/09/23
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByPID 根据PID取窗口句柄
'               增加函数：GetCaptionByHwnd 根据句柄取得标题
'               增加函数：SetTop设置窗体置顶，传入参数false则取消                2014/09/24
'               增加函数：Shake、FadeIn、FadeOut 抖动、淡入、淡出特效
'          V2.0 修正函数：GetWindowByPID 遍历窗体Win7下有一些问题                2015/09/29
'               修正函数：GetWindowByAppName 遍历窗体Win7下有一些问题
'               修正函数：GetWindowByAppNameEx 遍历窗体Win7下有一些问题
'          V2.1 修正函数：ClickPoint 增加位置模式参数相对和绝对，默认相对        2018/06/05
'               增加函数：SelectComboBoxIndex 根据指定的index选择下拉框中的项
'                         上述方法得到网友Chen8013的不少帮助，特此感谢
'               增加函数：GetWindowByHwnd 根据指定的句柄确定窗口                 2018/07/22
'               增加函数：GetWindowByCursorPos 根据当前光标获取窗口（控件）
'               增加函数：GetWindowByPoint 根据指定的位置获取窗口（控件）
'               升级ClickPoint函数，支持点击前后分别延时，默认延时为0            2018/07/23
'          V2.2 修正正则：网友小凡反应了句柄和id存在负数的情况                   2020/01/08
'               优化属性：Caption(Get)，根据网友小凡的建议改成可获得文本框内容
'               增加方法：Wait 此方法原为clsWaitableTimer模块中，现集成进来      2020/01/09
'               增加方法：ClickCurrentPoint 点击当前点                           2020/01/10
'               增加方法：SetCursor(别名:SetPoint MoveCursor MoveCursorTo)
'               更新函数：将所有默认等待超时60秒的函数中默认等待时间都改为10秒
'               增加属性：Text、Value、Title（均为Caption别名）                  2020/01/12
'               优化代码：GetCaptionByHwnd采用原Caption(Get)代码，后者也做了调整
'               增加函数：GetCursorPosCurrent(别名：GetCursorPoint)得到当前坐标
'               优化函数：所有窗口获取的函数增加了是否过滤可见的参数             2020/01/16
'               增加函数：GetTextByHwnd（同GetCaptionByHwnd）
'               优化代码结构。将模块中能移过来的都移到类模块中了                 2020/01/19
'               增加函数：myIsWindowVisibled 判断窗体可见，长宽为0也认为不可见   2020/01/31
'               优化函数：GetTextByHwnd 网友小凡提供                             2020/02/03
'               增加函数：CommandLine（同AppCommandLine）                        2020/02/05
'               增加函数：MakeTransparent 设置窗口透明度                         2020/02/18
'               增加函数：MoveToCenter 移动窗口到屏幕中心
'               增加函数：IsTopmost 判断窗口是否为置顶                           2020/02/20
'               增加函数：GetWindowTextByHwnd 获得窗口标题，给窗口句柄专用       2020/02/28
'               修正函数：Focus 旧方法使用后会改变置顶窗口属性                   2020/03/02
'               增加函数：IsWin64 网友小凡提供                                   2020/03/12
'               修正函数：AppPath 网友小凡提供兼容64位系统的方法
'               修正函数：AppCommandLine 网友小凡做了兼容64位处理及其他代码优化  2020/03/15
'               增加函数：IsForegroundWindow 判断窗口是否为活动窗口              2020/03/17
'               增加函数：GetClassNameByHwnd 根据句柄得到类名
'               增加属性：ClassName(Get) 返回窗口的类名
'               更新函数：CheckWindow 返回值由Long改成Boolean了，并且设为Public
'               增加函数：Click 点击当前句柄或者指定句柄                         2020/06/29
'               更新函数：为兼容win10将设置窗口最大最小化改用SendMessage实现     2021/12/10
'               增加函数：Restore 还原窗口，比如最大最小化后需要还原
'          V2.3 增加函数：SendKeys 替代VB自带的，解决win10下提示拒绝的错误       2022/05/10
'               优化函数：GetMatchHwndFromWindow 设置搜集窗口信息最多尝试10次
'               增加函数：FileToClipboard 设置指定文件到剪切板                   2022/06/26
'               增加函数：Paste 粘贴内容，可以是字符串或文件
'               修正函数：GetWindowByClassNameEx 修复了do里面取类名的错误        2022/08/26
'               增加函数：GetWindowClassNameByHwnd  同GetClassNameByHwnd
'==============================================================================================
Option Explicit

'常量定义
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Const SM_CXSCREEN As Long = 0&
Private Const SM_CYSCREEN As Long = 1&

Private Const BM_CLICK = &HF5
Private Const GWL_ID = (-12)
Private Const GWL_STYLE = (-16)
Private Const WM_SYSCOMMAND = &H112
Private Const WM_SETFOCUS = &H7
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_CLOSE = &H10
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Const SC_MINIMIZE = &HF020&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_RESTORE = &HF120&
Private Const SC_MOVE = &HF010&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&

Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_TOPMOST = &H8

Private Const WS_VISIBLE = &H10000000

Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const WS_EX_LAYERED = &H80000

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Const CB_SETCURSEL = &H14E
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&
Private Const INFINITE = &HFFFF
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Const UNITS = 4294967296#
Private Const MAX_LONG = -2147483648#

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const CF_HDROP = &HF

'结构体定义
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type DROPFILES
    pFiles As Long
    pT As POINTAPI
    fNC As Long
    fWide As Long
End Type

Public Enum enumShift
    Horizontal
    Vertical
    Both
End Enum

Public Enum enumPositionMode
    ABSOLUTE
    RELATIVE
End Enum
'过滤窗口是否可见
Public Enum enumWindowVisible
    HiddenWindow
    DisplayedWindow
    AllWindow
End Enum

Public Enum enumWindowState
    vbNormal = 0
    vbMinimized = 1
    vbMaximized = 2
End Enum

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "PSAPI.DLL" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As Any, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As rect) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Declare Function QueryFullProcessImageName Lib "Kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As Any, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileName Lib "PSAPI.DLL" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageName As String, nSize As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const DelayOneTime = 300
Private hWnd_ As Long '窗口句柄
Private wReturn As New clsWindow '作为返回对象以便连用
Private Const strVersion = "V2.2"
Public DebugMe As Boolean '表示是否启用调试模式

'自定义变量声明
Private strWindowInfo$ '保存所有窗口的信息，格式为 句柄 文本内容
Private strWindowKeyWord$ '要参与的过滤的窗口的关键字，如果不需要过滤就留空

'移动和设置窗口到指定位置，以及设置窗口的宽高。
Public Function Move(Optional ByVal X, Optional ByVal Y, Optional ByVal nWidth, Optional ByVal nHeight) As Long
    If IsMissing(X) Then X = Left
    If IsMissing(Y) Then Y = Top
    If IsMissing(nWidth) Then nWidth = Width
    If IsMissing(nHeight) Then nHeight = Height
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Move 被调用。参数为 x=" & X & vbTab & " y=" & Y & vbTab & " nWidth=" & nWidth & vbTab & " nHeight=" & nHeight, False
    End If
    
    Move = MoveWindow(hWnd_, X, Y, nWidth, nHeight, True)
End Function
'移动窗口到屏幕中心
Public Function MoveToCenter() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 MoveToCenter 被调用。", False
    End If
    
    Dim X&, Y&, nWidth&, nHeight&
    nWidth = Width
    nHeight = Height
    X = (GetSystemMetrics(SM_CXSCREEN) - nWidth) \ 2
    Y = (GetSystemMetrics(SM_CYSCREEN) - nHeight) \ 2
    MoveToCenter = MoveWindow(hWnd_, X, Y, nWidth, nHeight, True)
End Function
'得到当前窗口的句柄
Public Property Get hWnd() As Long
    hWnd = hWnd_
End Property
'设置类的句柄
Public Property Let hWnd(ByVal hWnd As Long)
    If IsWindow(hWnd) Then
        hWnd_ = hWnd
    Else
        hWnd_ = 0
    End If
End Property
'得到窗口的Left位置
Public Property Get Left() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Left = rect.Left
    Else
        Left = -1
    End If
End Property
'设置窗口的Left位置
Public Property Let Left(ByVal lngLeft As Long)
    Move lngLeft
End Property
'得到窗口的Top位置
Public Property Get Top() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Top = rect.Top
    Else
        Top = -1
    End If
End Property
'设置窗口的Top位置
Public Property Let Top(ByVal lngTop As Long)
    Move , lngTop
End Property
'得到窗口的Width大小
Public Property Get Width() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Width = rect.Right - rect.Left
    Else
        Width = -1
    End If
End Property
'设置窗口的Width
Public Property Let Width(ByVal lngWidth As Long)
    Move , , lngWidth
End Property
'得到窗口的Height大小
Public Property Get Height() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Height = rect.Bottom - rect.Top
    Else
        Height = -1
    End If
End Property
'设置窗口的Width
Public Property Let Height(ByVal lngHeight As Long)
    Move , , , lngHeight
End Property

'设置当前窗口的标题栏文字
Public Property Let Caption(ByVal strNewText As String)
    If hWnd_ <> 0 Then
        SendMessage hWnd_, WM_SETTEXT, 0&, ByVal strNewText
    End If
End Property
'得到当前窗口的标题栏文字 使用GetWindowTextByHwnd函数
Public Property Get Caption() As String
    Caption = GetText(hWnd_)
End Property
'根据句柄获得窗口内容
Public Function GetCaptionByHwnd(ByVal hWnd As Long) As String
    GetCaptionByHwnd = GetTextByHwnd(hWnd)
End Function
'根据句柄获得窗口内容
Public Function GetTextByHwnd(ByVal hWnd As Long) As String
    GetTextByHwnd = GetText(hWnd)
End Function
'根据句柄获得窗口类名
Public Function GetWindowClassNameByHwnd(ByVal hWnd As Long) As String
    GetWindowClassNameByHwnd = GetClassNameByHwnd(hWnd)
End Function
'根据句柄获得窗口类名
Public Function GetClassNameByHwnd(ByVal hWnd As Long) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetClassNameByHwnd 被调用。参数为 hWnd=" & hWnd, False
    End If
    
    Dim lngHwnd&, strClassNameCurrent As String * 255
    GetClassName hWnd, strClassNameCurrent, 255  '获得窗口类
    GetClassNameByHwnd = Replace(strClassNameCurrent, Chr$(0), "")
End Function
'根据句柄获得窗口内容 窗口专用
Public Function GetWindowTextByHwnd(ByVal hWnd As Long) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowTextByHwnd 被调用。参数为 hWnd=" & hWnd, False
    End If
    
    Dim l&, strCaption$
    l = GetWindowTextLength(hWnd)
    strCaption = Space(l)
    GetWindowText hWnd, strCaption, l + 1
    GetWindowTextByHwnd = strCaption
End Function
'得到当前窗口的类名
Public Property Get ClassName() As String
    ClassName = GetClassNameByHwnd(hWnd_)
End Property
'设置当前窗口的标题栏文字
Public Property Let Text(ByVal strNewText As String)
    Caption = strNewText
End Property
'得到当前窗口的标题栏文字
Public Property Get Text() As String
    Text = Caption
End Property
'设置当前窗口的标题栏文字
Public Property Let Value(ByVal strNewText As String)
    Caption = strNewText
End Property
'得到当前窗口的标题栏文字
Public Property Get Value() As String
    Value = Caption
End Property
'设置当前窗口的标题栏文字
Public Property Let Title(ByVal strNewText As String)
    Caption = strNewText
End Property
'得到当前窗口的标题栏文字
Public Property Get Title() As String
    Title = Caption
End Property
'窗口置前
Public Function SetPosFront() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetPosFront 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
        SetPosFront = SetWindowPos(hWnd_, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    End If
End Function
'窗口置后
Public Function SetPosBottom() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetPosBottom 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
        SetPosBottom = SetWindowPos(hWnd_, SW_SHOWNORMAL, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    End If
End Function
'设置窗口位置正常
Public Function SetPosNormal() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetPosNormal 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
        SetPosNormal = SetWindowPos(hWnd_, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    End If
End Function
'设置窗口最小化
Public Function Minimized() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Minimized 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
'        Minimized = ShowWindow(hWnd_, SW_SHOWMINIMIZED)
        Minimized = SendMessage(hWnd_, WM_SYSCOMMAND, SC_MINIMIZE, ByVal 0&)
    End If
End Function
'设置窗口最大化
Public Function Maximized() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Maximized 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
'        Maximized = ShowWindow(hWnd_, SW_SHOWMAXIMIZED)
        Maximized = SendMessage(hWnd_, WM_SYSCOMMAND, SC_MAXIMIZE, ByVal 0&)
    End If
End Function
'设置窗口还原
Public Function Restore() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Restore 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
        Restore = SendMessage(hWnd_, WM_SYSCOMMAND, SC_RESTORE, ByVal 0&)
    End If
End Function

'设置窗口位置隐藏
Public Function Hide() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Hide 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
        Hide = ShowWindow(hWnd_, SW_HIDE)
    End If
End Function
'设置窗口显示
Public Function Show(Optional ByVal ShowStatus As VbAppWinStyle = vbNormalFocus) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Show 被调用。", False
    End If
    If hWnd_ <> 0 Then
        Show = ShowWindow(hWnd_, ShowStatus)
    End If
End Function
'设置窗口正常显示
Public Function Normal() As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Normal 被调用。", False
    End If
    
    If hWnd_ <> 0 Then
        Normal = ShowWindow(hWnd_, SW_NORMAL)
    End If
End Function
'设置当前窗口为活动窗口
Public Sub Focus()
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Focus 被调用。", False
    End If
    
    Dim hForeWnd&, dwForeID&, dwCurID&, isTopmostCurrentWindow As Boolean
    isTopmostCurrentWindow = IsTopMost()
    hForeWnd = GetForegroundWindow()
    dwForeID = GetWindowThreadProcessId(hForeWnd, 0)
    dwCurID = GetCurrentThreadId()
    AttachThreadInput dwCurID, dwForeID, True
'    ShowWindow hWnd_, SW_SHOWNORMAL

    If Not isTopmostCurrentWindow Then '非置顶窗口可以设置置顶再取消置顶
        SetWindowPos hWnd_, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
        SetWindowPos hWnd_, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    End If
    SetForegroundWindow hWnd_
    AttachThreadInput dwCurID, dwForeID, False
End Sub
'设置当前窗口为活动窗口，有时无效，暂且保留
Public Sub Focus2()
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Focus2 被调用。", False
    End If
    
    BringWindowToTop hWnd_
    SetFocusAPI hWnd_
End Sub
'支持和VB中通常名称的方法
Public Sub SetFocus()
    Call Focus
End Sub
'得到进程标识
Public Function ProcessID(Optional ByVal hWnd1&) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ProcessID 被调用。参数为 hWnd1=" & hWnd1, False
    End If
    
    If hWnd1 = 0 Then hWnd1 = hWnd_ '如果没写则表示处理当前存储的句柄
    GetWindowThreadProcessId hWnd1, ProcessID
End Function
'得到进程标识
Public Function pid(Optional ByVal hWnd1&) As Long
    pid = ProcessID(hWnd1)
End Function
'判断64位系统
Public Function IsWin64() As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 IsWin64 被调用。", False
    End If
    
    Dim lngReturn As Long
    Call IsWow64Process(GetCurrentProcess, lngReturn)
    If lngReturn = 0 Then
        IsWin64 = False '非64位系统
    Else
        IsWin64 = True '64位系统
    End If
End Function
'根据窗口句柄得到应用程序的路径
Public Function AppPath(Optional ByVal hWn1&) As String
    If hWn1 = 0 Then hWn1 = hWnd_ '如果没写则表示处理当前存储的句柄

    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ProcessID 被调用。参数为 hWn1=" & hWn1, False
    End If
    
    Dim dwProcessId As Long, hProcess As Long, i&, sTmp2 As String * 255, sTmp$
    
    dwProcessId = ProcessID(hWn1)
    If dwProcessId = 0 Then Exit Function

    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, dwProcessId)
    AppPath = String$(255, vbNullChar)
    
    If IsWin64 Then '64位系统
        Call GetProcessImageFileName(hProcess, AppPath, 255)
        For i = 65 To 90 'Asc("A") To Asc("Z")
            Call QueryDosDevice(Chr$(i) & ":", sTmp2, 255)
            sTmp = Split(sTmp2, vbNullChar, 2)(0) & "\"
            If Len(sTmp) > 1 And InStr(1, AppPath, sTmp, vbTextCompare) = 1 Then
                AppPath = Chr$(i) & ":" & Mid$(AppPath, Len(sTmp))
                Exit For
            End If
        Next
    Else
        Call GetModuleFileNameEx(hProcess, 0, StrPtr(AppPath), 255)
    End If
    AppPath = Split(AppPath, vbNullChar, 2)(0)
    Call CloseHandle(hProcess)
End Function
'得到exe名字,如果缺省就获取自己的exe名称
Public Function AppName(Optional ByVal hWnd1&) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 AppName 被调用。参数为 hWnd1=" & hWnd1, False
    End If
    
    Dim v
    If hWnd1 = 0 Then hWnd1 = hWnd_ '如果没写则表示处理当前存储的句柄
    v = Split(AppPath(hWnd1) & "\", "\")
    AppName = v(UBound(v) - 1)
End Function
'得到exe文件的启动参数
Public Function AppCommandLine(Optional ByVal lPid&) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 AppCommandLine 被调用。参数为 lPid=" & lPid, False
    End If
    
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    Dim objProType As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    If lPid = 0 Then lPid = Me.pid '如果没写则表示处理当前存储的PID
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where ProcessId='" & lPid & "'")
    For Each objProcess In colProcessList
        For Each objProType In objProcess.Properties_
            If objProType.Name = "CommandLine" Then
                AppCommandLine = objProType.Value & ""  'Value为Null时须连接空字符串才不会出错
                Exit For
            End If
        Next
    Next
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 AppCommandLine 执行完毕。", False
    End If
End Function
'得到exe文件的启动参数
Public Function CommandLine(Optional ByVal lPid&) As String
    If lPid = 0 Then lPid = Me.pid '如果没写则表示处理当前存储的PID
    CommandLine = AppCommandLine(lPid&)
End Function
'检查窗口可见是否符合指定要求
Private Function CheckWindowVisible(ByVal lngHwnd&, ByVal windowVisible As enumWindowVisible) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 CheckWindowVisible 被调用。参数为 lngHwnd=" & lngHwnd & vbTab & " windowVisible=" & windowVisible, False
    End If
    
    If lngHwnd <> 0 Then
        If windowVisible = AllWindow Then
            CheckWindowVisible = True
        ElseIf windowVisible = DisplayedWindow Then
            If myIsWindowVisibled(lngHwnd) Then CheckWindowVisible = True
        ElseIf windowVisible = HiddenWindow Then
            If Not myIsWindowVisibled(lngHwnd) Then CheckWindowVisible = True
        End If
    End If
End Function
'设置根据窗体标题得到窗口句柄，可以指定等待几秒检测指定标题的窗体是否出现，默认是等待10秒
Public Function GetWindowByTitle(ByVal strTitle$, Optional ByVal intWaitSeconds& = 10, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByTitle 被调用。参数为 strTitle=" & strTitle & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " windowVisible=" & windowVisible, False
    End If

    Dim lngDelayCount&
    hWnd_ = 0
    Do
        hWnd_ = FindWindow(vbNullString, strTitle)
        If CheckWindowVisible(hWnd_, windowVisible) Then Exit Do '窗口可见度检查，符合条件表示找到了即退出

        lngDelayCount = lngDelayCount + DelayOneTime
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '未发现关键字为 strTitle
        Wait DelayOneTime

        hWnd_ = 0
    Loop

    wReturn.hWnd = hWnd_
    Set GetWindowByTitle = wReturn
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByTitle 执行完毕。", False
    End If
End Function
'和GetWindowByTitle函数功能类似，只是这个是模糊匹配
Public Function GetWindowByTitleEx(ByVal strTitle$, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional isUseRegExp As Boolean = False, Optional ByVal checkPid = "", Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByTitleEx 被调用。参数为 strTitle=" & strTitle & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " isUseRegExp=" & isUseRegExp & vbTab & " checkPid=" & checkPid & vbTab & " windowVisible=" & windowVisible, False
    End If
    
    Dim lngHwnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strHwndAllTmp$
    Dim isMatch As Boolean

    Do
        hWnd_ = 0
        Do
            lngHwnd = FindWindowEx(0, lngHwnd, vbNullString, vbNullString)
            If CheckWindowVisible(lngHwnd, windowVisible) Then  '窗口可见度检查，符合条件则继续下面的检查
                strCaption = GetWindowTextByHwnd(lngHwnd) '使用窗口专用的函数
                If strCaption <> "" Then
                    If Not isUseRegExp Then
                        isMatch = InStr(strCaption, strTitle) > 0
                    Else
                        isMatch = regTest(strCaption, strTitle)
                    End If
                Else
                    isMatch = (strTitle = "") '用户可能就需要获取窗口标题为空的呢
                End If
                If isMatch Then
                    strHwndAllTmp = strHwndAllTmp & CStr(lngHwnd) & " "
                    If checkPid <> "" Then '如果pid相等，优先放到第一个
                        If checkPid = ProcessID() Then strHwndAllTmp = lngHwnd & " " & strHwndAllTmp
                    End If
                End If
            End If
        Loop Until lngHwnd = 0 '这层loop是遍历所有窗口的
        If strHwndAllTmp <> "" Then Exit Do '如果一遍循环得到结果了就退出

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回

    If strHwndAllTmp <> "" Then
        wReturn.hWnd = Split(strHwndAllTmp, " ")(0)
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByTitleEx = wReturn
    hWnd_ = wReturn.hWnd
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByTitleEx 执行完毕。", False
    End If
End Function
'设置根据窗体类名返回窗体句柄，可以指定等待几秒检测指定标题的窗体是否出现，默认是等待10秒
Public Function GetWindowByClassName(ByVal strClassName$, Optional ByVal intWaitSeconds& = 10, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByClassName 被调用。参数为 strClassName=" & strClassName & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " windowVisible=" & windowVisible, False
    End If
    
    Dim lngDelayCount&
    hWnd_ = 0
    Do
        hWnd_ = FindWindow(strClassName, vbNullString)
        If CheckWindowVisible(hWnd_, windowVisible) Then Exit Do '窗口可见度检查，符合条件表示找到了即退出

        lngDelayCount = lngDelayCount + DelayOneTime
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strClassName
        Wait DelayOneTime

        hWnd_ = 0
    Loop

    wReturn.hWnd = hWnd_
    Set GetWindowByClassName = wReturn
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByClassName 执行完毕。", False
    End If
End Function
'设置根据窗体类名得到窗口句柄，可以指定等待几秒检测指定类名的窗体是否出现，默认是等待10秒，允许使用正则表达式
Public Function GetWindowByClassNameEx(ByVal strClassName$, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional isUseRegExp = False, Optional ByVal checkPid = "", Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByClassNameEx 被调用。参数为 strClassName=" & strClassName & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " isUseRegExp=" & isUseRegExp & vbTab & " checkPid=" & checkPid & vbTab & " windowVisible=" & windowVisible, False
    End If
        
    Dim lngHwnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strClassNameCurrent As String
    Dim strHwndAllTmp$
    Dim isMatch As Boolean

    hWnd_ = 0
    Do
        Do
            lngHwnd = FindWindowEx(0, lngHwnd, vbNullString, vbNullString)
            If CheckWindowVisible(lngHwnd, windowVisible) Then
                strClassNameCurrent = GetWindowClassNameByHwnd(lngHwnd)
                If strClassNameCurrent <> "" Then
                    If Not isUseRegExp Then
                        isMatch = InStr(strClassNameCurrent, strClassName) > 0
                    Else
                        isMatch = regTest(strClassNameCurrent, strClassName)
                    End If
                Else
                    isMatch = (strClassName = "") '用户可能就需要获取窗口标题为空的呢
                End If
                If isMatch Then
                    strHwndAllTmp = strHwndAllTmp & CStr(lngHwnd) & " "
                    If checkPid <> "" Then '如果pid相等，优先放到第一个
                        If checkPid = ProcessID() Then strHwndAllTmp = lngHwnd & " " & strHwndAllTmp
                    End If
                End If
            End If
        Loop Until lngHwnd = 0
        If strHwndAllTmp <> "" Then Exit Do

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回

    If strHwndAllTmp <> "" Then
        wReturn.hWnd = Split(strHwndAllTmp, " ")(0)
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByClassNameEx = wReturn
    hWnd_ = wReturn.hWnd
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByClassNameEx 执行完毕。", False
    End If
End Function
'设置根据进程名称返回窗体句柄，可以指定等待几秒检测指定进程名称的窗体是否出现，默认是等待10秒
Public Function GetWindowByAppName(ByVal strAppName$, Optional ByVal intWaitSeconds& = 10, Optional ByVal isWholeMatch As Boolean = False, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByAppName 被调用。参数为 strAppName=" & strAppName & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " isWholeMatch=" & isWholeMatch & vbTab & " windowVisible=" & windowVisible, False
    End If
    
    Dim lngHwnd&, currWnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim isMatch As Boolean
    Dim strGetHwndAppName$
    Dim strHwnd

    strAppName = LCase$(strAppName)
    hWnd_ = 0
    Do
        currWnd = GetDesktopWindow()
        currWnd = GetWindow(currWnd, GW_CHILD)
        Do While currWnd <> 0
            If CheckWindowVisible(currWnd, windowVisible) Then
                strGetHwndAppName = LCase$(AppName(currWnd)) '根据句柄得到进程名字，都转换成小写
                If isWholeMatch Then '判断是否完全相等
                    isMatch = (LCase$(strGetHwndAppName) = strAppName)
                Else
                    isMatch = (InStr(LCase$(strGetHwndAppName), strAppName) > 0)
                End If
                strCaption = GetWindowTextByHwnd(currWnd)
                If InStr("|MSCTFIME UI|Program Manager|M|Default IME|", strCaption) = 0 And isMatch Then '过滤掉输入法相关的窗口
                    hWnd_ = currWnd
                    Exit Do
                End If
            End If
            currWnd = GetWindow(currWnd, GW_HWNDNEXT)

            DoEvents
        Loop
        If hWnd_ <> 0 Then Exit Do

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    wReturn.hWnd = hWnd_
    Set GetWindowByAppName = wReturn
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByAppName 执行完毕。", False
    End If
End Function
'设置根据进程名称返回窗体句柄，可以指定等待几秒检测指定进程名称的窗体是否出现，默认是等待10秒，原理还是枚举所有窗体检查它们的所属进程，只不过是根据目标的进程名跟strAppName比对
Public Function GetWindowByAppNameEx(ByVal strAppName$, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional isUseRegExp = False, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByAppNameEx 被调用。参数为 strAppName=" & strAppName & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " isUseRegExp=" & isUseRegExp & vbTab & " windowVisible=" & windowVisible, False
    End If
        
    Dim lngHwnd&, currWnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim strHwndAllTmp$
    Dim isMatch As Boolean
    Dim strAppNameCurrent$
    Dim strHwnd

    strAppName = LCase$(strAppName)

    hWnd_ = 0
    strHwndAllTmp = ""
    Do
        currWnd = GetDesktopWindow()
        currWnd = GetWindow(currWnd, GW_CHILD)
        Do While currWnd <> 0
            If CheckWindowVisible(currWnd, windowVisible) Then   '窗口可见度检查，符合条件则继续下面的检查
                strAppNameCurrent = LCase$(AppName(currWnd))  '根据句柄得到进程名字，都转换成小写
                If strAppNameCurrent <> "" Then
                    If Not isUseRegExp Then '如果不用正则就包含关系匹配
                        isMatch = InStr(strAppNameCurrent, strAppName) > 0
                    Else
                        isMatch = regTest(strAppNameCurrent, strAppName)
                    End If
                Else
                    isMatch = (strAppName = "") '用户可能就需要获取窗口标题为空的呢
                End If
                strCaption = GetWindowTextByHwnd(currWnd)
                If InStr("|MSCTFIME UI|Program Manager|M|Default IME|", strCaption) = 0 And isMatch Then '过滤掉输入法相关的窗口
                    strHwndAllTmp = strHwndAllTmp & CStr(currWnd) & " "
                End If
            End If

            currWnd = GetWindow(currWnd, GW_HWNDNEXT)
            DoEvents
        Loop
        If strHwndAllTmp <> "" Then Exit Do '表示匹配到了一部分数据

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回

    If strHwndAllTmp <> "" Then
        wReturn.hWnd = Split(strHwndAllTmp, " ")(0)
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByAppNameEx = wReturn
    hWnd_ = wReturn.hWnd
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByAppNameEx 执行完毕。", False
    End If
End Function
'根据pid获得窗体
Public Function GetWindowByPID(ByVal pid As Long, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByPID 被调用。参数为 pid=" & pid & vbTab & " intWaitSeconds=" & intWaitSeconds & vbTab & " windowVisible=" & windowVisible, False
    End If
    
    Dim lngHwnd&, currWnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim strHwndAllTmp$
    Dim isMatch As Boolean
    Dim currentPID$
    Dim strHwnd

    hWnd_ = 0
    strHwndAllTmp = ""
    Do '循环找几遍
        currWnd = GetDesktopWindow()
        currWnd = GetWindow(currWnd, GW_CHILD)
        Do While currWnd <> 0
            If CheckWindowVisible(currWnd, windowVisible) Then
                currentPID = ProcessID(currWnd) '得到pid
                isMatch = (currentPID = pid)
                If isMatch Then
                    strCaption = GetWindowTextByHwnd(currWnd)

                    If InStr("|MSCTFIME UI|Program Manager|M|Default IME|", strCaption) = 0 Then '过滤掉输入法相关的窗口
                        strHwndAllTmp = strHwndAllTmp & CStr(currWnd) & " "
                    End If
                End If
            End If
            currWnd = GetWindow(currWnd, GW_HWNDNEXT)
            DoEvents
        Loop
        If strHwndAllTmp <> "" Then Exit Do '表示匹配到了一部分数据

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回

    If strHwndAllTmp <> "" Then
        Dim vHwnd
        vHwnd = Split(strHwndAllTmp, " ")
        wReturn.hWnd = vHwnd(UBound(vHwnd))
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByPID = wReturn
    hWnd_ = wReturn.hWnd
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByPID 执行完毕。", False
    End If
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'以下几个方法为网友wwb（邮箱：wwbing@gmail.com）所增加
'日期：2013-6-26
'具体有：CheckWindow, Load, WindowState, Visible, hDC,
'       AlphaBlend, Enabled, Refresh, TransparentColor, ZOrder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'检查当前窗口句柄是否还有效
Public Function CheckWindow(Optional ByVal lngHwnd&) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 CheckWindow 被调用。参数为 lngHwnd=" & lngHwnd, False
    End If
    
    If lngHwnd = 0 Then lngHwnd = hWnd_ '如果没写则表示处理当前存储的句柄
    
    CheckWindow = (IsWindow(lngHwnd) = 1)
End Function
'Load 窗口载入的入口
Public Function Load(WindowHwnd As Variant) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Load 被调用。参数为 WindowHwnd=" & WindowHwnd, False
    End If
    
    If IsNumeric(WindowHwnd) Then
        hWnd_ = CLng(WindowHwnd)
        If CheckWindow Then
            wReturn.hWnd = hWnd_
            Set Load = wReturn
        Else
            MsgBox "类初始化错误！句柄“" & WindowHwnd & "”对应的窗口不存在！", vbExclamation
        End If
    Else
        MsgBox "类初始化错误！设置的句柄“" & WindowHwnd & "”应当为一个数字。", vbExclamation
    End If
End Function
'直接由句柄指定窗口，与直接设置w.hwnd或者load方法同等效果
Public Function GetWindowByHwnd(WindowHwnd As Variant) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByHwnd 被调用。参数为 WindowHwnd=" & WindowHwnd, False
    End If
    
    If IsNumeric(WindowHwnd) Then
        hWnd_ = CLng(WindowHwnd)
        If CheckWindow Then
            wReturn.hWnd = hWnd_
            Set GetWindowByHwnd = wReturn
        Else
            MsgBox "类初始化错误！句柄“" & WindowHwnd & "”对应的窗口不存在！", vbExclamation
        End If
    Else
        MsgBox "类初始化错误！设置的句柄“" & WindowHwnd & "”应当为一个数字。", vbExclamation
    End If
End Function
'直接获得当前鼠标下的窗口或控件的句柄
Public Function GetWindowByCursorPos() As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByCursorPos 被调用。", False
    End If
    
    Dim tPoint As POINTAPI
    GetCursorPos tPoint '获得当前鼠标位置
    hWnd_ = WindowFromPoint(tPoint.X, tPoint.Y)
    wReturn.hWnd = hWnd_
    Set GetWindowByCursorPos = wReturn
End Function
'根据指定的点获取窗口或控件的句柄
Public Function GetWindowByPoint(ByVal X As Long, ByVal Y As Long) As clsWindow
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetWindowByPoint 被调用。参数为 X=" & X & vbTab & " Y=" & Y, False
    End If
    
    hWnd_ = WindowFromPoint(X, Y)
    wReturn.hWnd = hWnd_
    Set GetWindowByPoint = wReturn
End Function
'根据控件的类名设置控件的显示文字
Public Function SetElementTextByClassName(ByVal strClassName$, ByVal strNewText$, Optional ByVal intIndex% = 1, Optional ByVal UseRegExp As Boolean = False) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetElementTextByClassName 被调用。参数为 strClassName=" & strClassName & vbTab & " strNewText=" & strNewText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long

    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If

    If lngControlsHwnd <> 0 Then
        SendMessage lngControlsHwnd, WM_SETTEXT, 0&, ByVal strNewText
        SetElementTextByClassName = True
    Else
        SetElementTextByClassName = False
    End If
End Function
'根据控件的类名追加显示控件的文字
Public Function AppendElementTextByClassName(ByVal strClassName$, ByVal strAppendText$, Optional ByVal intIndex% = 1, Optional ByVal UseRegExp As Boolean = False) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 AppendElementTextByClassName 被调用。参数为 strClassName=" & strClassName & vbTab & " strAppendText=" & strAppendText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If

    If lngControlsHwnd <> 0 Then
        strAppendText = RTrim$(GetElementTextByClassName(strClassName, intIndex, UseRegExp)) & strAppendText
        AppendElementTextByClassName = SendMessage(lngControlsHwnd, WM_SETTEXT, 0&, ByVal strAppendText)
    Else
        AppendElementTextByClassName = 0
    End If
End Function
'根据控件的类名得到控件的显示文字
Public Function GetElementTextByClassName(ByVal strClassName$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetElementTextByClassName 被调用。参数为 strClassName=" & strClassName & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long

    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If

    If lngControlsHwnd <> 0 Then
        GetElementTextByClassName = GetTextByHwnd(lngControlsHwnd)
    Else
        GetElementTextByClassName = ""
    End If
End Function
'根据控件的内容得到控件的显示文字
Public Function GetElementTextByText(ByVal strText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetElementTextByText 被调用。参数为 strText=" & strText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If

    If lngControlsHwnd <> 0 Then
        GetElementTextByText = GetTextByHwnd(lngControlsHwnd)
    Else
        GetElementTextByText = ""
    End If
End Function
'根据控件的类名得到句柄
Public Function GetElementHwndByClassName(ByVal strClassName$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetElementHwndByClassName 被调用。参数为 strClassName=" & strClassName & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        GetElementHwndByClassName = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        GetElementHwndByClassName = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
End Function
'根据控件的文本内容得到控件的句柄
Public Function GetElementHwndByText(ByVal strText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetElementHwndByText 被调用。参数为 strText=" & strText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long

    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        GetElementHwndByText = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        GetElementHwndByText = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
End Function
'根据控件的显示文字设置新的显示文字
Public Function SetElementTextByText(ByVal strText$, ByVal strNewText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetElementTextByText 被调用。参数为 strText=" & strText & vbTab & " strNewText=" & strNewText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
        
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
    If lngControlsHwnd <> 0 Then
        SendMessage lngControlsHwnd, WM_SETTEXT, 0&, ByVal strNewText
        SetElementTextByText = True
    Else
        SetElementTextByText = False
    End If
End Function
'点击指定句柄
Public Function Click(Optional ByVal hWnd1&) As Long
    If hWnd1 = 0 Then hWnd1 = hWnd_ '如果没写则表示处理当前存储的句柄
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Click 被调用" & vbCrLf & "关闭窗口，句柄为=" & hWnd1, False
    End If
    
'    Click = PostMessage(hWnd1, BM_CLICK, 0&, ByVal 0)
    PostMessage hWnd1, WM_LBUTTONDOWN, 0, ByVal (1 + 1 * 65536)
    PostMessage hWnd1, WM_LBUTTONUP, 0, ByVal (1 + 1 * 65536)

End Function
'根据控件的类名点击控件
Public Function ClickElementByClassName(ByVal strClassName$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp = False) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ClickElementByClassName 被调用。参数为 strClassName=" & strClassName & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
    If lngControlsHwnd <> 0 Then
        Focus
        ClickElementByClassName = PostMessage(lngControlsHwnd, BM_CLICK, 0&, ByVal 0)
    Else
        ClickElementByClassName = 0
    End If
End Function
'根据控件的显示文字点击控件
Public Function ClickElementByText(ByVal strText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp = False) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ClickElementByText 被调用。参数为 strText=" & strText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
    End If
    
    Dim lngControlsHwnd As Long
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
    If lngControlsHwnd <> 0 Then
        Focus
        ClickElementByText = PostMessage(lngControlsHwnd, BM_CLICK, 0&, ByVal 0)
    Else
        ClickElementByText = 0
    End If
End Function
'从匹配结果中得到句柄，一般是从一个窗口中的所有子窗口中进行检索
Private Function GetMatchHwndFromWindow(ByVal Pattern As String, Optional ByVal intIndex = 1, Optional ByVal UseRegExp = False, Optional ByVal isUseClass = True) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetMatchHwndFromWindow 被调用。参数为 Pattern=" & Pattern & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp & vbTab & " isUseClass=" & isUseClass, False
    End If
    
    Dim reg As Object
    Dim matchs As Object, match As Object
    Dim strControlsInfo As String
    Dim intGet As Integer
    Dim intRetryCount As Integer
    
err_retry:
    strControlsInfo = ControlsInfo(hWnd_, DebugMe)
    Set reg = CreateObject("vbscript.regexp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True

    If Not UseRegExp Then
        reg.Pattern = Pattern
        Set matchs = reg.Execute(strControlsInfo)
        If matchs.Count >= intIndex Then
            GetMatchHwndFromWindow = CLng(matchs(intIndex - 1).SubMatches(0))
        End If
    Else '每行结构为="43452836    0   ToolbarWindow32 缩放级别”
        reg.Pattern = "^([-\d]+)\t[-\d]+\t(.*?)\t(.*?)$" '默认检索获得“ToolbarWindow32”
        If Not isUseClass Then reg.Pattern = "^([-\d]+)\t[-\d]+\t.*?\t(.*?)$" '如果不是类名那么就用文本内容进行正则检索“缩放级别”
        Set matchs = reg.Execute(strControlsInfo)
        For Each match In matchs
            If regTest(match.SubMatches(1), Pattern) Then '匹配的
                intGet = intGet + 1
                GetMatchHwndFromWindow = CLng(match.SubMatches(0))
                If intGet = intIndex Then Exit For '如果已经确定当前索引次序就是要处理的，那么退出检索循环
            End If
        Next
    End If
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " >> 函数 getMatchHwndFromWindow 执行完毕。匹配数量为=" & matchs.Count, False
    End If

    '如果获取结果为0，那么可能是执行太快，目标窗口还没渲染好，所以我们延时50ms再检测下，最多重复执行10次
    If matchs.Count = 0 And intRetryCount < 9 Then
        Wait 50
        intRetryCount = intRetryCount + 1
        GoTo err_retry
    End If
End Function
Public Function CloseWindow(Optional ByVal hWnd1&) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 closeWindow 被调用。参数为 hWnd1=" & hWnd1, False
    End If
    
    If hWnd1 = 0 Then hWnd1 = hWnd_ '如果没写则表示处理当前存储的句柄

    CloseWindow = PostMessage(hWnd1, WM_CLOSE, 0, 0)
End Function
Public Function CloseApp(Optional ByVal hWnd1&) As Long
    Dim hProcess&
    If hWnd1 = 0 Then hWnd1 = hWnd_ '如果没写则表示处理当前存储的句柄
    hProcess = OpenProcess(&H1F0FFF, False, ProcessID(hWnd1))
    CloseApp = TerminateProcess(hProcess, 0)
    CloseHandle hProcess
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数CloseApp执行完毕" & vbCrLf & "关闭窗口，句柄为=" & hWnd1, False
    End If
End Function
'点击具体的某个点
Public Sub ClickPoint(x1 As Long, y1 As Long, Optional ByVal PositionMod As enumPositionMode = RELATIVE, Optional ByVal clickTimes As Integer = 1, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ClickPoint 被调用。参数为 x1=" & x1 & vbTab & " y1=" & y1 & vbTab & " PositionMod=" & PositionMod & vbTab & " clickTimes=" & clickTimes & vbTab & " delayMilliSecondsBefore=" & delayMilliSecondsBefore & vbTab & " delayMilliSecondsAfter=" & delayMilliSecondsAfter, False
    End If
    
    If PositionMod = ABSOLUTE Then
        SetCursorPos x1, y1
    ElseIf PositionMod = RELATIVE Then
        SetCursorPos Left + x1, Top + y1
    Else
        MsgBox "函数ClickPoint执行一次。未能识别点击模式" & PositionMod & "，要么是relative(相对位置)要么是absolute(绝对位置)，默认是相对位置。", vbExclamation
        Exit Sub
    End If
    Do While clickTimes > 0
        Wait delayMilliSecondsBefore
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        Wait delayMilliSecondsAfter
        clickTimes = clickTimes - 1
    Loop
End Sub
'点击当前坐标
Public Sub ClickCurrentPoint()
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ClickCurrentPoint 被调用。", False
    End If
    
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
'移动鼠标到指定点
Public Sub SetCursor(ByVal x1 As Long, ByVal y1 As Long, Optional ByVal PositionMod As enumPositionMode = ABSOLUTE, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetCursor 被调用。参数为 x1=" & x1 & vbTab & " y1=" & y1 & vbTab & " PositionMod=" & PositionMod & vbTab & " delayMilliSecondsBefore=" & delayMilliSecondsBefore & vbTab & " delayMilliSecondsAfter=" & delayMilliSecondsAfter, False
    End If
    
    If PositionMod = ABSOLUTE Then
        Wait delayMilliSecondsBefore
        SetCursorPos x1, y1
        Wait delayMilliSecondsAfter
    ElseIf PositionMod = RELATIVE Then
        Wait delayMilliSecondsBefore
        SetCursorPos Left + x1, Top + y1
        Wait delayMilliSecondsAfter
    Else
        MsgBox "函数SetCursor执行一次。未能识别点击模式" & PositionMod & "，要么是relative(相对位置)要么是absolute(绝对位置)，默认是相对位置。", vbExclamation
        Exit Sub
    End If
End Sub
'同SetCursor
Public Sub SetPoint(ByVal x1 As Long, ByVal y1 As Long, Optional ByVal PositionMod As enumPositionMode = ABSOLUTE, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    SetCursor x1, y1, PositionMod, delayMilliSecondsBefore, delayMilliSecondsAfter
End Sub
'同SetCursor
Public Sub MoveCursor(ByVal x1 As Long, ByVal y1 As Long, Optional ByVal PositionMod As enumPositionMode = ABSOLUTE, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    SetCursor x1, y1, PositionMod, delayMilliSecondsBefore, delayMilliSecondsAfter
End Sub
'同SetCursor
Public Sub MoveCursorTo(ByVal x1 As Long, ByVal y1 As Long, Optional ByVal PositionMod As enumPositionMode = ABSOLUTE, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    SetCursor x1, y1, PositionMod, delayMilliSecondsBefore, delayMilliSecondsAfter
End Sub
'直接获得当前鼠标下的窗口或控件的句柄
Public Function GetCursorPosCurrent() As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 GetCursorPosCurrent 被调用。", False
    End If
    
    Dim tPoint As POINTAPI
    GetCursorPos tPoint  '获得当前鼠标位置
    GetCursorPosCurrent = tPoint.X & "," & tPoint.Y
End Function
'直接获得当前鼠标下的窗口或控件的句柄
Public Function GetCursorPoint() As String
    GetCursorPoint = GetCursorPosCurrent()
End Function
'以下是几个特效的方法
'特效1：窗体抖动
Public Sub Shake(Optional ByVal shakeShift As enumShift = Both, Optional ByVal shakeRepeats = 30, Optional ByVal shakePads = 10, Optional ByVal shakeInterval = 25)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Shake 被调用。参数为 shakeShift=" & shakeShift & vbTab & " shakeRepeats=" & shakeRepeats & vbTab & " shakePads=" & shakePads & vbTab & " shakeInterval=" & shakeInterval, False
    End If
    
    Dim i%

    If Not CheckWindow() Then Exit Sub
    Select Case shakeShift
        Case Horizontal
            For i = 1 To shakeRepeats
                If i Mod 2 = 0 Then
                    Left = Left + shakePads
                Else
                    Left = Left - shakePads
                End If
                Wait shakeInterval
            Next
        Case Vertical
            For i = 1 To shakeRepeats
                If i Mod 2 = 0 Then
                    Top = Top + shakePads
                Else
                    Top = Top - shakePads
                End If
                Wait shakeInterval
            Next
        Case Both
            For i = 1 To shakeRepeats
                If i Mod 2 = 0 Then
                    Left = Left + shakePads
                    Wait 10
                    Top = Top + shakePads
                Else
                    Left = Left - shakePads
                    Wait 10
                    Top = Top - shakePads
                End If
                Wait shakeInterval
            Next
    End Select
End Sub
'特效2：淡入 第二个参数只能在1到100之间，表示透明度截止百分比
Public Sub FadeIn(Optional ByVal speed = 10, Optional ByVal intPercentTansparentRemains)
    Dim sty As Long, i As Long
    Dim intMaxTrans As Integer

    If Not CheckWindow() Then Exit Sub
    sty = GetWindowLong(hWnd_, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    SetWindowLong hWnd_, GWL_EXSTYLE, sty

    If IsMissing(intPercentTansparentRemains) Then
        intMaxTrans = 255
    ElseIf intPercentTansparentRemains < 0 Or intPercentTansparentRemains > 255 Then
        intMaxTrans = 255
    Else
        intMaxTrans = (intPercentTansparentRemains / 100) * 255
    End If
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 FadeIn 被调用。参数为 speed=" & speed & vbTab & " intPercentTansparentRemains~=" & intMaxTrans, False
    End If

'    SetLayeredWindowAttributes hWnd_, 0, 100, LWA_ALPHA
    If speed > 255 Then speed = 255
    If speed < 1 Then speed = 1
    For i = 1 To 255 Step speed
        SetLayeredWindowAttributes hWnd_, 0, i, LWA_ALPHA   '设置窗体透明度
        Wait 50
        If i > intMaxTrans Then Exit Sub
    Next
    SetLayeredWindowAttributes hWnd_, 0, 255, LWA_ALPHA
End Sub
'特效3：淡出
'可以选择是否关闭窗口或者经常
Public Sub FadeOut(Optional ByVal speed = 10, Optional ByVal isCloseWindow = False, Optional ByVal isCloseApp = False)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 FadeOut 被调用。参数为 speed=" & speed & vbTab & " isCloseWindow=" & isCloseWindow & vbTab & " isCloseApp=" & isCloseApp, False
    End If
    
    Dim sty As Long, i As Long

    If Not CheckWindow() Then Exit Sub
    sty = GetWindowLong(hWnd_, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    SetWindowLong hWnd_, GWL_EXSTYLE, sty

'    SetLayeredWindowAttributes hWnd_, 0, 100, LWA_ALPHA
    If speed > 255 Then speed = 255
    If speed < 1 Then speed = 1
    For i = 255 To 1 Step -speed
        SetLayeredWindowAttributes hWnd_, 0, i, LWA_ALPHA   '设置窗体透明度
        Wait 50
    Next
    SetLayeredWindowAttributes hWnd_, 0, 0, LWA_ALPHA
    If isCloseWindow Then CloseWindow
    If isCloseApp Then CloseApp
End Sub
'窗体设置透明，Percent 只接受：0到100的数、0%到100%、以及介于0到1的小数
Public Function Transparent(Optional ByVal Percent = "50", Optional ByVal lngHwnd&) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Transparent 被调用。参数为 Percent=" & Percent & vbTab & " lngHwnd=" & lngHwnd, False
    End If
    
    If lngHwnd = 0 Then lngHwnd = hWnd_ '如果没写则表示处理当前存储的句柄

    Dim sty As Long
    
    If Not CheckWindow(lngHwnd) Then Exit Function

    Percent = Trim$(CStr(Percent))
    If regTest(Percent, "^0?\.\d+$") Then '例如：0.5
        Percent = Val(Percent) * 255
    ElseIf regTest(Percent, "^\d+(\.\d+)?$") Then '例如：50或者50.5
        Percent = Percent / 100 * 255 '将百分比转换成255范围内显示
    ElseIf regTest(Percent, "^\d+(\.\d+)?%$") Then '例如：50%或者50.5%
        Percent = Replace(Percent, "%", "") / 100 * 255
    Else
'        MsgBox "参数值：Percent=" & Percent & " 不合法，只接受：0到100的整数、0%到100%、以及介于0到1的小数。", vbExclamation
        Exit Function
    End If

    If Percent > 255 Then Percent = 255
    If Percent < 0 Then Percent = 0

    sty = GetWindowLong(lngHwnd, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    SetWindowLong lngHwnd, GWL_EXSTYLE, sty
    Transparent = SetLayeredWindowAttributes(lngHwnd, 0, Percent, LWA_ALPHA)  '表示执行结果，1表示成功，0表示失败
End Function
'同Transparent
Public Function MakeTransparent(Optional ByVal Percent = "50", Optional ByVal lngHwnd) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 MakeTransparent 被调用。参数为 Percent=" & Percent & vbTab & " lngHwnd=" & lngHwnd, False
    End If
    
    If IsMissing(Percent) And IsMissing(lngHwnd) Then
        MakeTransparent = Transparent
    ElseIf IsMissing(lngHwnd) Then
        MakeTransparent = Transparent(Percent)
    ElseIf IsMissing(Percent) Then
        MakeTransparent = Transparent(, lngHwnd)
    End If
End Function
'同Transparent
Public Function SetTransparent(Optional ByVal Percent = "50", Optional ByVal lngHwnd) As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetTransparent 被调用。参数为 Percent=" & Percent & vbTab & " lngHwnd=" & lngHwnd, False
    End If
    
    If IsMissing(Percent) And IsMissing(lngHwnd) Then
        SetTransparent = Transparent
    ElseIf IsMissing(lngHwnd) Then
        SetTransparent = Transparent(Percent)
    ElseIf IsMissing(Percent) Then
        SetTransparent = Transparent(, lngHwnd)
    End If
End Function
'得到窗口状态
Public Property Get WindowState() As enumWindowState
    If Not CheckWindow() Then Exit Property
    If IsZoomed(hWnd_) Then
        WindowState = vbMaximized
    ElseIf IsIconic(hWnd_) Then
        WindowState = vbMinimized
    Else
        WindowState = vbNormal
    End If
End Property
'设置窗口状态
Public Property Let WindowState(ByVal vNewValue As enumWindowState)
ShowWindow hWnd_, SW_HIDE
ShowWindow hWnd_, SW_SHOWNOACTIVATE
Exit Property
    If Not CheckWindow() Then Exit Property
    Select Case vNewValue
        Case 0
            ShowWindow hWnd_, SW_NORMAL
        Case 1
            ShowWindow hWnd_, SW_SHOWMINIMIZED
        Case 2
            ShowWindow hWnd_, SW_SHOWMAXIMIZED
    End Select
End Property

Public Property Get Visible() As Boolean
    If Not CheckWindow() Then Exit Property '检查是否是窗口
    Visible = myIsWindowVisibled(hWnd_)
End Property
'检查窗口是否可见，如果长宽有一项为0也认为不可见
Private Function myIsWindowVisibled(ByVal hWnd_ As Long) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 myIsWindowVisibled 被调用。参数为 hWnd_=" & hWnd_, False
    End If
    
    myIsWindowVisibled = CBool(IsWindowVisible(hWnd_) = 1) '用api函数判断窗口是否可见
    '以下为检测宽高是否为0，如果有一项为0则认为不可见
    Dim rect As rect, chkWidth&, chkHeight&
    GetWindowRect hWnd_, rect
    chkWidth = rect.Right - rect.Left
    chkHeight = rect.Bottom - rect.Top
    If chkWidth = 0 Or chkHeight = 0 Then myIsWindowVisibled = False
End Function

Public Property Let Visible(ByVal vNewValue As Boolean)
    If vNewValue = True Then
        Call Show
    Else
        Call Hide
    End If
End Property

Public Property Get hDC() As Long
    If Not CheckWindow() Then Exit Property
    hDC = GetDC(hWnd_)
End Property

Public Property Let hDC(ByVal vNewValue As Long)
'
End Property

Public Property Get AlphaBlend() As Byte
    If Not CheckWindow() Then Exit Property
    GetLayeredWindowAttributes hWnd_, vbNull, AlphaBlend, vbNull
End Property

Public Property Let AlphaBlend(ByVal vNewValue As Byte)
    Dim r As Long, v As Long
    If Not CheckWindow() Then Exit Property
    r = GetWindowLong(hWnd_, GWL_EXSTYLE)
    If (r And WS_EX_LAYERED) = 0 Then
        r = r Or WS_EX_LAYERED
        SetWindowLong hWnd_, GWL_EXSTYLE, r
    End If
    r = 0
    GetLayeredWindowAttributes hWnd_, v, vbNull, r
    SetLayeredWindowAttributes hWnd_, v, vNewValue, r Or LWA_ALPHA
End Property

Public Property Get Enabled() As Boolean
    If Not CheckWindow() Then Exit Property
    Enabled = CBool(IsWindowEnabled(hWnd_) = 1)
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    If Not CheckWindow() Then Exit Property
    EnableWindow hWnd_, vNewValue
End Property

Public Sub Refresh()
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 Refresh 被调用。", False
    End If
    
    If Not CheckWindow() Then Exit Sub
    UpdateWindow hWnd_
End Sub

Public Property Get TransparentColor() As Long
    If Not CheckWindow() Then Exit Property
    GetLayeredWindowAttributes hWnd_, TransparentColor, vbNull, vbNull
End Property

Public Property Let TransparentColor(ByVal vNewValue As Long)
    Dim r As Long, v As Byte
    If Not CheckWindow() Then Exit Property
    r = GetWindowLong(hWnd_, GWL_EXSTYLE)
    If (r And WS_EX_LAYERED) = 0 Then
        r = r Or WS_EX_LAYERED
        SetWindowLong hWnd_, GWL_EXSTYLE, r
    End If
    r = 0
    GetLayeredWindowAttributes hWnd_, vbNull, v, r
    SetLayeredWindowAttributes hWnd_, vNewValue, v, r Or LWA_COLORKEY
End Property
'设置窗体置顶
Public Sub SetTop(Optional ByVal isSetTop As Boolean = True)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 SetTop 被调用。参数为 isSetTop=" & isSetTop, False
    End If
    
    SetWindowPos hWnd, IIf(isSetTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub
'判断窗口置顶函数
Public Function IsTopMost(Optional ByVal hWnd&) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 IsTopMost 被调用。参数为 hWnd=" & hWnd, False
    End If
    
    If hWnd = 0 Then hWnd = hWnd_ '如果没写则表示处理当前存储的句柄
    Dim Ret As Long, t1 As Long, t2 As Long
    Ret = GetWindowLong(hWnd, GWL_EXSTYLE)
    t1 = Ret Or WS_EX_TOPMOST
    IsTopMost = (Ret = t1)
End Function
'判断当前窗口是否为活动窗口
Public Function IsForegroundWindow(Optional ByVal hWnd&) As Boolean
    If hWnd = 0 Then hWnd = hWnd_ '如果没写则表示处理当前存储的句柄
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 IsForegroundWindow 被调用。参数为 hWnd=" & hWnd, False
    End If
    
    IsForegroundWindow = (GetForegroundWindow() = hWnd_)
End Function
'此方法未经测试
Public Sub ZOrder()
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数 ZOrder 被调用。", False
    End If
    
    If Not CheckWindow() Then Exit Sub
    
    BringWindowToTop hWnd_
End Sub
'得到版本信息
Public Property Get Version() As String
    Version = strVersion
End Property
'根据指定的index选择下拉框中内容
Public Sub SelectComboBoxIndex(ByVal hWnd As Long, ByVal intIndex As Integer)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 过程 SelectComboBoxIndex 被调用。参数为 hWnd=" & hWnd & vbTab & " intIndex=" & intIndex, False
    End If
    
    SendMessage hWnd, CB_SETCURSEL, intIndex, 0
End Sub
'延时，单位为毫秒
Public Function Wait(ByVal MilliSeconds As Long)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 过程 Wait 被调用。参数为 MilliSeconds=" & MilliSeconds, False
    End If
    
    Dim dSavetime As Double
    dSavetime = timeGetTime + MilliSeconds   '记下开始时的时间
    While timeGetTime < dSavetime '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Wend
End Function
Private Function replaceReg(ByVal str1$, ByVal strPattern$, ByVal strNew$) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 过程 replaceReg 被调用。参数为 str1=" & str1 & vbTab & " strPattern=" & strPattern & vbTab & " strNew=" & strNew, False
    End If
    
    Dim strData As String
    Dim reg As Object

    strData = str1 ' "1)2{}?*3.7(7"

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern ' "([\\+-\.()\[\]{}?*])"
    replaceReg = reg.Replace(strData, strNew) '"\$1"
End Function
Private Function regTest(ByVal str1$, ByVal strPattern$) As Boolean
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 过程 regTest 被调用。参数为 str1=" & str1 & vbTab & " strPattern=" & strPattern, False
    End If
    
    Dim strData As String
    Dim reg As Object

    strData = str1 ' "1)2{}?*3.7(7"

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern ' "([\\+-\.()\[\]{}?*])"
    regTest = reg.test(strData)  '"\$1"
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据所给文件名和内容直接写文件
'函数名：writeToFile
'入口参数(如下)：
'  strFileName 所给的文件名；
'  strContent 要输入到上述文件的字符串
'  isCover 是否覆盖该文件，默认为覆盖
'返回值：True或False，成功则返回前者，否则返回后者
'备注：sysdzw 于 2007-5-2 提供
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo Err1
    Dim fileHandl%
    fileHandl = FreeFile
    If isCover Then
        Open strFileName For Output As #fileHandl
    Else
        Open strFileName For Append As #fileHandl
    End If
    Print #fileHandl, strContent
    Close #fileHandl
    writeToFile = True
    Exit Function
Err1:
    writeToFile = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：得到所有控件的信息，是按次序获得的，可用于编写脚本的参考和程序设置值时使用。此函数需要和EnumChildProc一起使用
'函数名：ControlsInfo
'入口参数：hWnd   long型  容器句柄，一般指窗体句柄。可缺省，缺省为最新获取到的窗体的句柄，也可以指定一个句柄
'返回值：string   保存了容器内所有控件的信息，包含“句柄、ID、类名、显示文字”
'备注：sysdzw 于 2010-11-13 提供
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ControlsInfo(ByVal lngMainHwnd As Long, Optional isDebug As Boolean = False) As String
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 过程 ControlsInfo 被调用。参数为 lngMainHwnd=" & lngMainHwnd & vbTab & " isDebug=" & isDebug, False
    End If
    
    Dim lRet&
    strControlInfo = ""
'    Wait 1000
    lRet = EnumChildProc(lngMainHwnd, 0&)
    lRet = EnumChildWindows(lngMainHwnd, AddressOf EnumChildProc, 0&)
    ControlsInfo = strControlInfo
    If isDebug Then writeToFile "controls.txt", strControlInfo
End Function
'替代vb自带的sendkeys，可以解决win10下提示拒绝的问题
Public Sub SendKeys(ByVal strKeys As String)
    Dim ws As Object
    Set ws = CreateObject("wscript.shell")
    ws.SendKeys strKeys
    Set ws = Nothing
End Sub
'粘贴内容，可以是一个字符串也可以是一个文件，默认为字符串形式，如果是文件请传递第二个参数为true
Public Sub Paste(ByVal strKeys As String, Optional isFile As Boolean = False)
    If isFile Then
        FileToClipboard strKeys
    Else
        Clipboard.Clear
        Clipboard.SetText strKeys
    End If
    
    SendKeys "^v"
End Sub
'拷贝指定文件或者文件夹到剪切板，等同于右击文件或者文件夹选择复制
Public Function FileToClipboard(ByVal FileName As String)
    Dim dFiles As DROPFILES
    Dim hMem As Long, MemPtr As Long, Len1 As Long, Len2 As Long
    
    If Dir(FileName) = "" And Dir(FileName, vbDirectory) = "" Then
        FileToClipboard = "error:file not exits"
        Clipboard.Clear
        Clipboard.SetText FileToClipboard
        Exit Function
    End If
    
    FileName = FileName & vbNullChar & vbNullChar
      
    If OpenClipboard(Me.hWnd) = 0 Then Exit Function
    Call EmptyClipboard
    Len1 = Len(dFiles)
    Len2 = LenB(StrConv(FileName, vbFromUnicode))
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len1 + Len2)
    MemPtr = GlobalLock(hMem)
    dFiles.pFiles = Len1
    Call CopyMemory(ByVal MemPtr, dFiles, Len1)
    Call CopyMemory(ByVal (MemPtr + Len1), ByVal FileName, Len2)
    Call GlobalUnlock(hMem)
    Call SetClipboardData(CF_HDROP, hMem)
    Call CloseClipboard
    FileToClipboard = "ok"
End Function
'测试专用
Public Function test() As String
    test = SendMessage(hWnd_, WM_LBUTTONDOWN, 100, 100)
End Function

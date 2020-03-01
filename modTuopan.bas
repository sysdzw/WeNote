Attribute VB_Name = "modTuopan"
Option Explicit

Public Const DefaultIconIndex = 1    '图标缺省索引
Public Const WM_LBUTTONDOWN = &H201    '按鼠标左键
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204    '按鼠标右键
Public Const WM_RBUTTONUP = &H205
Public Const NIM_ADD = 0    '添加图标
Public Const NIM_MODIFY = 1    '修改图标
Public Const NIM_DELETE = 2    '删除图标
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = 1    'message 有效
Public Const NIF_ICON = 2    '图标操作（添加、修改、删除）有效
Public Const NIF_TIP = 4    'ToolTip(提示）有效

'图标操作
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'判断窗口是否最小化
Declare Function IsIconic Lib "user32" (ByVal Hwnd As Long) As Long
'设置窗口位置和状态（position）的功能
Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'定义类型
'通知栏图标状态
Public Type NOTIFYICONDATA
    cbSize As Long
    Hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'添加图标至通知栏
Public Function Icon_Add(iHwnd As Long, sTips As String, hIcon As Long, IconID As Long) As Long
    '参数说明：iHwnd：窗口句柄，sTips：当鼠标移到通知栏图标上时显示的提示内容
    'hIcon：图标句柄，IconID：图标Id号
Dim IconVa As NOTIFYICONDATA
    With IconVa
        .Hwnd = iHwnd
        .szTip = sTips + Chr$(0)
        .hIcon = hIcon
        .uID = IconID
        .uCallbackMessage = WM_MOUSEMOVE
        .cbSize = Len(IconVa)
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        Icon_Add = Shell_NotifyIcon(NIM_ADD, IconVa)
    End With
End Function
'删除通知栏图标(参数说明同Icon_Add)
Function Icon_Del(iHwnd As Long, lIndex As Long) As Long
Dim IconVa As NOTIFYICONDATA
Dim l As Long
    With IconVa
        .Hwnd = iHwnd
        .uID = lIndex
        .cbSize = Len(IconVa)
    End With
    Icon_Del = Shell_NotifyIcon(NIM_DELETE, IconVa)
End Function


Attribute VB_Name = "modWindow"
'=====================================================================================
'描    述：是clsWindow.cls类的依赖模块，一些无法放到类模块中的代码放在这里 (modWindow)
'编    程：sysdzw 原创开发，如果有需要对模块进行更新请发我一份，共同维护
'发布日期：2013/05/28
'博    客：http://blog.csdn.net/sysdzw
'用户手册：https://www.kancloud.cn/sysdzw/clswindow/
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                        2012/12/3
'          V1.1 将类中的api函数以及部分变量挪到此模块         2013/05/28
'          V1.2 将EnumChildProc中获取控件文字函数修改了      2013/06/13
'          V1.3 将本模块中能移到类模块中的都移过去了          2020/01/19
'               将GetText函数放到这儿来了                   2020/03/12
'=====================================================================================
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const lMaxLength& = 500
Public strControlInfo$ '保存容器内所有控件的信息
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：和api函数EnumChildWindows结合使用得到一个窗体容器内的所有child控件
'函数名：EnumChildProc
'入口参数：hWnd   long型  容器句柄，一般指窗体句柄
'返回值：long   这里直接返回的true，如果是true则继续调用
'备注：sysdzw 于 2010-11-13 提供
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim strClassName As String * 256
    Dim strText As String
    Dim lngCtlId As Long
    Dim strHwnd$, strCtlId$, strClass$, lRet&

    EnumChildProc = True

    lngCtlId = GetWindowLong(hWnd, (-12))
    lRet = GetClassName(hWnd, strClassName, 255)

    strText = GetText(hWnd)
    strText = Replace(strText, vbCrLf, " ") '强制将文本框中内容回车替换成空格，以防止影响正则获取

    strHwnd$ = CStr(hWnd) & vbTab
    strCtlId$ = CStr(lngCtlId) & vbTab
    strClass$ = Left$(strClassName, lRet) & vbTab
    strControlInfo = strControlInfo & strHwnd$ & _
                    strCtlId$ & _
                    strClass$ & _
                    strText & vbCrLf
End Function
'根据句柄获得窗口内容
Public Function GetText(ByVal hWnd As Long) As String
    '方案1 性能一般
'    Dim Txt2() As Byte, i&
'    i = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
'    If i = 0 Then Exit Function '没有内容
'    ReDim Txt2(i)
'    SendMessage hWnd, WM_GETTEXT, i + 1, Txt2(0)
'    ReDim Preserve Txt2(i - 1)
'    GetTextByHwnd = StrConv(Txt2, vbUnicode)

    '方案2 混合方案，尽量减少api调用（本代码由网友小凡提供）
    Dim Txt2() As Byte, i&
    ReDim Txt2(lMaxLength&) '须比实际内容多设一个字节来装结束符0
    SendMessage hWnd, &HD, lMaxLength&, Txt2(0)
    If Txt2(0) = 0 Then Exit Function  '没有内容
    For i = 1 To lMaxLength&
        If Txt2(i) = 0 Then Exit For '结束
    Next
    If i >= lMaxLength - 2& Then '如果接近就视为取内容不完整，直接用api计算长度取
        i = SendMessage(hWnd, &HE, 0&, 0&)
        If i = 0 Then Exit Function '没有内容
        ReDim Txt2(i) '须比实际内容多设一个字节来装结束符0
        SendMessage hWnd, &HD, i + 1, Txt2(0)
    End If
    ReDim Preserve Txt2(i - 1) '去掉多的字节
    GetText = StrConv(Txt2, vbUnicode) '转ASI字串为宽字串
End Function

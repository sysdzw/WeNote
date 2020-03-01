Attribute VB_Name = "modWindow"
'=====================================================================================
'描    述：是clsWindow.cls类的依赖模块，一些无法放到类模块中的代码放在这里 (modWindow)
'编    程：sysdzw 原创开发，如果有需要对模块进行更新请发我一份，共同维护
'发布日期：2013/05/28
'博    客：http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                        2012/12/3
'          V1.1 将类中的api函数以及部分变量挪到此模块         2013/05/28
'          V1.2 将EnumChildProc中获取控件文字函数修改了      2013/06/13
'          V1.3 将本模块中能移到类模块中的都移过去了          2020/01/19
'=====================================================================================
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public strControlInfo$ '保存容器内所有控件的信息
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：和api函数EnumChildWindows结合使用得到一个窗体容器内的所有child控件
'函数名：EnumChildProc
'入口参数：hWnd   long型  容器句柄，一般指窗体句柄
'返回值：long   这里直接返回的true，如果是true则继续调用
'备注：sysdzw 于 2010-11-13 提供
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim Txt2(64000) As Byte
    Dim strClassName As String * 255
    Dim strText As String
    Dim lngCtlId As Long
    Dim strHwnd$, strCtlId$, strClass$, lRet&
    
    EnumChildProc = True
    
    lngCtlId = GetWindowLong(hwnd, (-12))
    lRet = GetClassName(hwnd, strClassName, 255)
    
    SendMessage hwnd, &HD, 64000, Txt2(0)
    strText = Split(StrConv(Split(Txt2, Chr$(0), 2)(0), vbUnicode) & Chr$(0), Chr$(0), 2)(0)
    strText = Replace(strText, vbCrLf, " ") '强制将文本框中内容回车替换成空格，以防止影响正则获取
    
    strHwnd$ = CStr(hwnd) & vbTab
    strCtlId$ = CStr(lngCtlId) & vbTab
    strClass$ = Left$(strClassName, lRet) & vbTab
    
    strControlInfo = strControlInfo & strHwnd$ & _
                    strCtlId$ & _
                    strClass$ & _
                    strText & vbCrLf
End Function

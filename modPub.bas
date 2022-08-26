Attribute VB_Name = "modPub"
Option Explicit
Public strAppPath As String '应用程序目录
Public strSetFile As String
Public strSet As String
Public strDataFile As String
Public strData As String
Public isHasCreateIcon As Boolean
Public strInfo As String
Public strInitData As String
Public lngCurrentIndex As Long  '当前id

Public lngLeftLatest&, lngTopLatest& '最新的坐标。

Public isFormAllLoadCompleted As Boolean
Public isFirstNote As Boolean '是否是第一张

Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

's,n,h,d,m,yyyy
Public Type ShijianDanwei
    strTag As String 'h
    strName As String '时
    strNameShow As String '展示的名称，更符合日常的叫法
    lngSeconds As Long
End Type

Type ChooseColor
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type
Public tDanwei(6) As ShijianDanwei

Public Const NOTE_DEFAULT_WIDTH = 4395 '便签默认宽度
Public Const NOTE_DEFAULT_HEIGHT = 3615 '便签默认高度
Public Const NEW_NOTE_MOVE_RIGHT = 320 '新便签
Public Const NEW_NOTE_MOVE_DOWN = 570

Public lngHwndDesktop As Long '桌面的句柄
Public isNeedSetToDesktop As Boolean '是否需要设置嵌入到桌面

Sub Main()
    Dim w As New clsWindow
    If App.PrevInstance Then '防止重复运行
        w.GetWindowByTitle("WeNote", 1).Focus  '调出当前已经打开任意的窗口激活显示
        End
    End If

    strAppPath = App.Path
    If Right(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    
    strInfo = "WeNote | 微便签 V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
        "  作者:sysdzw" & vbCrLf & _
        "  主页:https://blog.csdn.net/sysdzw" & vbCrLf & _
        "  Q  Q:171977759" & vbCrLf & _
        "  邮箱:sysdzw@163.com" & vbCrLf & vbCrLf & _
        "2020-01-20"
    Call initDanwei
    
    lngHwndDesktop = w.GetWindowByClassName("Progman", 1).hwnd  '得到桌面句柄
    isNeedSetToDesktop = isSetToDesktop()
    
    Load frmStartup
    strDataFile = strAppPath & "数据.txt"
    strData = fileStr(strDataFile)
    If strData <> "" Then
        Dim vLine, i%, j%
        vLine = Split(strData, vbCrLf)
        For i = 0 To UBound(vLine)
            If vLine(i) <> "" Then
                lngCurrentIndex = Split(vLine(i), vbTab)(0) '最新id
                strInitData = vLine(i)
                Call NewNote
            End If
        Next
    Else
        isFirstNote = True
        Call NewNote
    End If
    isFormAllLoadCompleted = True
End Sub
'添加一个便签
Private Sub NewNote()
    Dim frmNewNote As New frmNote
    Load frmNewNote
End Sub
'初始化时间单位
Private Sub initDanwei()
    Dim vTag, vName, vNameShow, vJinweiBefore, vJinweiAfter, vSeconds, i%
    vTag = Split("s,n,h,d,m,yyyy", ",")
    vName = Split("秒,分,时,日,月,年", ",")
    vNameShow = Split("秒,分钟,小时,天,月,年", ",")
    vSeconds = Split("1,60,3600,86400,2592000,31104000,31536000", ",")
    For i = 0 To UBound(vTag)
        tDanwei(i).strTag = vTag(i)
        tDanwei(i).strName = vName(i)
        tDanwei(i).strNameShow = vNameShow(i)
        tDanwei(i).lngSeconds = vSeconds(i)
    Next
End Sub
'得到单位的索引
Public Function getDanweiIndex(ByVal strDanweiName$) As Integer
    Dim v, i%, intDanwei%
    For i = 0 To UBound(tDanwei)
        If tDanwei(i).strName = strDanweiName Or tDanwei(i).strNameShow = strDanweiName Then
            getDanweiIndex = i
            Exit Function
        End If
    Next
End Function
'设置combobox高度
Public Sub setComboHeight(oComboBox As ComboBox, lNewHeight As Long)
    Dim oldscalemode As Integer
    Dim lngLeft&, lngTop&, lngWidth&
    lngLeft = oComboBox.Left
    lngTop = oComboBox.Top
    lngWidth = oComboBox.Width
    If TypeOf oComboBox.Parent Is Frame Then Exit Sub
    oldscalemode = oComboBox.Parent.ScaleMode
    oComboBox.Parent.ScaleMode = vbPixels
    MoveWindow oComboBox.hwnd, lngLeft \ 15, lngTop \ 15, lngWidth \ 15, lNewHeight, 1
    oComboBox.Parent.ScaleMode = oldscalemode
End Sub
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
    On Error GoTo err1
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
err1:
    writeToFile = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据所给的文件名返回文件的内容
'函数名：fileStr
'入口参数(如下)：
'  strFileName 所给的文件名；
'返回值：文件的内容
'备注：sysdzw 于 2007-5-3 提供
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function fileStr(ByVal strFileName As String) As String
    On Error GoTo err1
    Dim lFile&
    lFile = FreeFile
    Open strFileName For Input As #lFile
    fileStr = StrConv(InputB$(LOF(lFile), #lFile), vbUnicode)
    Close #lFile
    Do While InStr(fileStr, vbCrLf & vbCrLf) > 0
        fileStr = Replace$(fileStr, vbCrLf & vbCrLf, vbCrLf)
    Loop
    If Left(fileStr, 2) = vbCrLf Then fileStr = Mid(fileStr, 3)
    If Right(fileStr, 2) = vbCrLf Then fileStr = Left(fileStr, Len(fileStr) - 2)
    Exit Function
err1:
'    MsgBox "不存在该文件或该文件不能访问！" & vbCrLf & strFileName, vbExclamation
End Function
Public Function regGetStrSub1(ByVal strData$, strPattern$) As String
     Dim reg As Object
     Dim matchs As Object, match As Object
    
     Set reg = CreateObject("vbscript.regexp")
     reg.Global = True
     reg.IgnoreCase = True
     reg.Pattern = strPattern
    
     Set matchs = reg.Execute(strData$)
     If matchs.Count >= 1 Then
         regGetStrSub1 = matchs(0).SubMatches(0)
     End If
End Function
'5分钟 、1秒、1天 等等
'得到括号匹配的所有结果，列用制表符，行用回车隔开
Public Function regGetStrSubAll(ByVal strData$, strPattern$) As String
     Dim reg As Object
     Dim matchs As Object, match As Object, i As Integer, j As Integer
    
     Set reg = CreateObject("vbscript.regexp")
     reg.Global = True
     reg.IgnoreCase = True
     reg.Pattern = strPattern

    Set matchs = reg.Execute(strData)
    For i = 0 To matchs.Count - 1
        For j = 0 To matchs(i).SubMatches.Count - 1
           regGetStrSubAll = regGetStrSubAll & matchs(i).SubMatches(j) & vbTab
        Next
        If Right(regGetStrSubAll, 1) = vbTab Then regGetStrSubAll = Left(regGetStrSubAll, Len(regGetStrSubAll) - 1)
        regGetStrSubAll = regGetStrSubAll & vbCrLf
    Next
    If Right(regGetStrSubAll, 2) = vbCrLf Then regGetStrSubAll = Left(regGetStrSubAll, Len(regGetStrSubAll) - 2)
End Function
'用正则对字符串进行替换，用法参考：regReplace("fas7f897fsa9fsd0f8", "\d+", "")
Public Function regReplace(ByVal strData$, strPattern$, strNewString$) As String
    Dim reg As Object
    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern
    regReplace = reg.Replace(strData, strNewString)
    Set reg = Nothing
End Function
'用正则对字符串进行测试是否匹配，用法参考：regTest("13895554788", "^\d{11}$")
Public Function regTest(ByVal strData$, strPattern$) As Boolean
    Dim reg As Object
    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern
    regTest = reg.Test(strData)
    Set reg = Nothing
End Function
'延时，单位为毫秒
Public Function Wait(ByVal MilliSeconds As Long)
    Dim dSavetime As Double
    dSavetime = timeGetTime + MilliSeconds   '记下开始时的时间
    While timeGetTime < dSavetime '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Wend
End Function
'检查是否设置到桌面
Public Function isSetToDesktop() As Boolean
    If GetSetting("WeNote", "Set", "SetToDesktop") = "" Then
        isSetToDesktop = False
    Else
        isSetToDesktop = (GetSetting("WeNote", "Set", "SetToDesktop") = "1")
        
    End If
End Function

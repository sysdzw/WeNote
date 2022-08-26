VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   0  'None
   Caption         =   "系统托盘管理"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4680
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Menu mnuSys 
      Caption         =   "系统菜单"
      WindowList      =   -1  'True
      Begin VB.Menu mnuNewNote 
         Caption         =   "新建一个便签(&N)"
      End
      Begin VB.Menu mnuShowAllNote 
         Caption         =   "显示所有便签(&V)"
      End
      Begin VB.Menu mnuHideAllNote 
         Caption         =   "隐藏所有便签(&H)"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunStartup 
         Caption         =   "开机启动(&S)"
      End
      Begin VB.Menu mnuToDesktop 
         Caption         =   "嵌入桌面(&Q)"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "使用帮助(&H)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于 微便签(&A)..."
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'名    称：微便签，WeNote
'描    述：微便签是一款windows操作系统下的便签软件，完全可以替代系统自带的便笺，每个
'          便签可单独设置提醒，可设置半透明。便签平常都是在系统右下角托盘区，不占用
'          任务栏。使用起来非常方便。具体使用方法可右击系统托盘菜单查看帮助。
'编    程：sysdzw 原创开发，如您对本软件进行改进或拓展请发我一份
'发布日期：2020-03-02
'博    客：https://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                                           2020-02-20
'          V1.1 解决了倒计时更新时会闪动的小问题                                2020-03-15
'==============================================================================================
Option Explicit
Dim isDealing As Boolean

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
 
Private Sub Form_Load()
    Icon_Add Me.hWnd, "微便签", Me.Icon, 0
    mnuRunStartup.Checked = isHasSetAutoRun()
    mnuToDesktop.Checked = isNeedSetToDesktop
End Sub
Private Sub mnuAbout_Click()
    MsgBox strInfo, vbInformation
End Sub

Public Sub mnuExit_Click()
    If isFormAllLoadCompleted Then
        Call Icon_Del(Me.hWnd, 0)
        
        Dim frm As Form
        Dim w As New clsWindow
        For Each frm In Forms
            If frm.Caption = "WeNote" Then
                Unload frm
            End If
        Next
        
        Unload Me
    Else
        MsgBox "请等待便签全部加载完毕再操作退出！", vbInformation
    End If
End Sub

Private Sub mnuHelp_Click()
    Dim strHelp$
    strHelp = "WeNote | 微便签 V" & App.Major & "." & App.Minor & "." & App.Revision & " 使用说明：" & vbCrLf & vbCrLf & _
        "本软件参考win7系统自带便签开发，不过比系统自带的好用方便且更强大。下面为具体使用方法：" & vbCrLf & vbCrLf & _
        "【1】新建便签。方法有3种：a.直接双击exe，如果当前没有便签会自动新建一个。b.点击已有便签的左上角+新建。c.右击系统托盘选择菜单“新建一个便签”" & vbCrLf & vbCrLf & _
        "【2】设置透明度。双击便签顶部打开设置，拉动滚动条调整成您希望的透明度。此项仅对当前便签有效。" & vbCrLf & vbCrLf & _
        "【3】设置便签颜色。双击便签顶部打开设置，点击备选中的颜色，如果没有您喜欢的可以点击“更多”到调色板进行选择。" & vbCrLf & vbCrLf & _
        "【4】设置窗口置顶。双击便签顶部打开设置，勾选“保持当前便签最前”。此项仅对当前便签有效。" & vbCrLf & vbCrLf & _
        "【5】设置闹钟提醒。双击便签顶部打开设置，勾选最后一项，并设置数量和时间单位，例如“5”、“分钟”。此项仅对当前便签有效。" & vbCrLf & vbCrLf & _
        "【6】设置开机启动。右击系统托盘图标，点击“设为开机自启动”，如果已经设置过会变成“手动启动”，点击会来回切换。此项全局生效。" & vbCrLf & vbCrLf & _
        "  如有问题可联系QQ171977759反馈" & vbCrLf & vbCrLf & _
        "2020-02-20"
    MsgBox strHelp, vbInformation
End Sub


Private Sub mnuNewNote_Click()
    Dim frmNote As New frmNote
    Load frmNote
End Sub
'开机启动
Private Sub mnuRunStartup_Click()
    mnuRunStartup.Checked = Not mnuRunStartup.Checked
    If mnuRunStartup.Checked Then
        Call setAutoRun
    Else
        Call cancelAutoRun
    End If
End Sub
'嵌入桌面
Private Sub mnuToDesktop_Click()
    mnuToDesktop.Checked = Not mnuToDesktop.Checked
    isNeedSetToDesktop = mnuToDesktop.Checked
    
'    Dim frm As Form
'    Dim w As New clsWindow
'    For Each frm In Forms
'        If frm.Caption = "WeNote" Then
'            SetParent frm.hWnd, IIf(isNeedSetToDesktop, lngHwndDesktop, 0)
'        End If
'    Next

    SaveSetting "WeNote", "Set", "SetToDesktop", IIf(isNeedSetToDesktop, "1", "0")
    If isNeedSetToDesktop Then
        MsgBox "已设置嵌入桌面，请重新打开软件以便生效。", vbInformation
        Call mnuExit_Click
    Else
        Dim frm As Form
        Dim w As New clsWindow
        For Each frm In Forms
            If frm.Caption = "WeNote" Then
                SetParent frm.hWnd, IIf(isNeedSetToDesktop, lngHwndDesktop, 0)
            End If
        Next
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX

    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
    Case WM_RBUTTONUP
        SetForegroundWindow Me.hWnd
        PopupMenu mnuSys
    Case WM_LBUTTONDOWN
        mnuShowAllNote_Click
    End Select
End Sub

Private Sub mnuShowAllNote_Click()
    Dim frm As Form
    Dim w As New clsWindow
    For Each frm In Forms
        If frm.Caption = "WeNote" Then
            w.hWnd = frm.hWnd
            frm.Visible = True
            w.Focus
        End If
    Next
End Sub

Private Sub mnuHideAllNote_Click()
    Dim frm As Form
    Dim w As New clsWindow
    For Each frm In Forms
        If frm.Caption = "WeNote" Then
            w.hWnd = frm.hWnd
            w.Hide
        End If
    Next
End Sub


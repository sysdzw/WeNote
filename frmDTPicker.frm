VERSION 5.00
Begin VB.Form frmDTPicker 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2160
   ClientLeft      =   5055
   ClientTop       =   1920
   ClientWidth     =   3330
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00CD895C&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00CD895C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   2190
      ScaleHeight     =   285
      ScaleWidth      =   525
      TabIndex        =   6
      Top             =   45
      Width           =   525
      Begin VB.Label lblMonth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "12月"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD895C&
         Height          =   255
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H00CD895C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   600
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   45
      Width           =   735
      Begin VB.Label lblYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "2010年"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CD895C&
         Height          =   255
         Left            =   15
         TabIndex        =   5
         Top             =   15
         Width           =   705
      End
   End
   Begin VB.Timer tmrMouseDown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1980
      Top             =   810
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   2715
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   30
      Width           =   315
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1868
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   30
      Width           =   315
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   1328
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   30
      Width           =   315
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   278
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   30
      Width           =   315
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00800000&
      BorderStyle     =   3  'Dot
      Height          =   315
      Left            =   720
      Top             =   900
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Menu popYear 
      Caption         =   "年"
      Visible         =   0   'False
      Begin VB.Menu mnuYear 
         Caption         =   "2005年"
         Index           =   0
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2006年"
         Index           =   1
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2007年"
         Index           =   2
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2008年"
         Index           =   3
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2009年"
         Index           =   4
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2010年"
         Index           =   5
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2011年"
         Index           =   6
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2012年"
         Index           =   7
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2013年"
         Index           =   8
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2014年"
         Index           =   9
      End
      Begin VB.Menu mnuYear 
         Caption         =   "2015年"
         Index           =   10
      End
   End
   Begin VB.Menu popMonth 
      Caption         =   "月"
      Visible         =   0   'False
      Begin VB.Menu mnuMonth 
         Caption         =   "1月"
         Index           =   0
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "2月"
         Index           =   1
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "3月"
         Index           =   2
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "4月"
         Index           =   3
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "5月"
         Index           =   4
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "6月"
         Index           =   5
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "7月"
         Index           =   6
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "8月"
         Index           =   7
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "9月"
         Index           =   8
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "10月"
         Index           =   9
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "11月"
         Index           =   10
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "12月"
         Index           =   11
      End
   End
End
Attribute VB_Name = "frmDTPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" _
                    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_ShowMDIWindow = &H40
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As rect) As Long
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'************************************************************************************
Private Const mc_GridRows& = 6
Private Const mc_Rows& = 8
Private Const mc_Cols& = 7


Private m_RowHeight As Single
Private m_ColWidth As Single
Private m_FirstRowY As Single
Private m_FirstColX As Single
'************************************************************************************
Public Event SelectDate(ByVal newDate As Date)
Public Event OnLoaded()
Public Event OnUnload()
'************************************************************************************
Private m_CurrentYear As Long
Private m_CurrentMonth As Long
Private m_FirstDate As Date

Private m_DefaultDate As Long
Private m_Inited As Boolean
Private m_Canceled As Boolean

Private m_blnLoaded As Boolean

Private m_MousePos As Integer
Private m_ButtonIndex As Integer
Private m_MouseDownButton As Integer

Private m_PopMenu As Integer
'**************************************************************************
'***************************************************************************

Public Sub ShowList(ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single, ByVal sngHeight As Single, ByVal defValue As Date)
    Dim iClientLeft As Long, iClientTop As Long, iClientRight As Long, iClientBottom As Long
    Dim sngCliLeft As Single, sngCliTop As Single, sngCliRight As Single, sngCliBottom As Single
    
    
    Dim sngWinWidth As Single, sngWinHeight As Single
    Dim sngWinLeft As Single, sngWinTop As Single
    
    m_blnLoaded = False
    
    '****************************************************************
    Call GetClientSize(iClientLeft, iClientTop, iClientRight, iClientBottom) '取得Windows桌面尺寸及位置
    sngCliLeft = iClientLeft * 15#
    sngCliTop = iClientTop * 15#
    sngCliRight = iClientRight * 15#
    sngCliBottom = iClientBottom * 15#
    '****************************************************************
    
    Call Load(Me)
    Me.CurrentDate = defValue
    
    sngWinWidth = Me.Width
    sngWinHeight = Me.Height
    If sngLeft + sngWinWidth > sngCliRight Then
        sngWinLeft = sngLeft + sngWidth - sngWinWidth
    Else
        sngWinLeft = sngLeft
    End If
    If sngTop + sngHeight + sngWinHeight > sngCliBottom Then
        sngWinTop = sngTop - sngWinHeight + 15
    Else
        sngWinTop = sngTop + sngHeight - 15
    End If
    Me.Move sngWinLeft, sngWinTop
    
    
    Call SetWindowPos(Me.hWnd, -1, sngWinLeft / 15, sngWinTop / 15, sngWinWidth / 15, sngWinHeight / 15, &H40)
    
    RaiseEvent OnLoaded
    If m_blnLoaded Then Exit Sub
    m_blnLoaded = True
    Call SetCapture(Me.hWnd)
End Sub

Public Property Get CurrentDate() As Date
    CurrentDate = CDate(m_DefaultDate)
End Property

Public Property Let CurrentDate(ByVal New_Value As Date)
    m_Inited = True
    m_DefaultDate = CLng(Int(New_Value))
End Property


Private Sub SelectDate()
    Dim dateValue As Date
    Dim iRow As Long, iCol As Long
    If m_MousePos > 14 Then
        iRow = (m_MousePos - 1) \ 7
        iCol = m_MousePos - iRow * 7 - 1
        
        If m_MousePos >= 55 Then
            dateValue = CDate(Format(Now, "YYYY-MM-DD"))
        Else
            dateValue = m_FirstDate + (m_MousePos - 14 - 1)
        End If
        m_Canceled = False
        Unload Me
        RaiseEvent SelectDate(dateValue)
    End If
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub




'***************************************************************************************
'Mouse Event
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iButtonIndex As Integer
    m_MousePos = GetCellIndex(X, Y)
    If m_MousePos < 0 Then
        Unload Me
    Else
        m_MouseDownButton = Button
        If Button = 1 Then
            If m_MousePos > 0 And m_MousePos < 5 Then
                iButtonIndex = m_MousePos - 1
                Call picButton_MouseDown(iButtonIndex, 1, Shift, X, Y)
            End If
        End If
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iButtonIndex As Integer
    Dim iCellIndex As Integer
    Dim iRow As Long, iCol As Long
    Dim sngLeft As Single, sngTop As Single, sngWidth As Single, sngHeight As Single
    Dim blnShapeVisible As Boolean
    
    iCellIndex = GetCellIndex(X, Y)
    
    If iCellIndex > 14 Then
        iRow = (iCellIndex - 1) \ 7
        iCol = iCellIndex - iRow * 7 - 1
        
        sngLeft = m_FirstColX + iCol * m_ColWidth
        sngTop = m_FirstRowY + iRow * m_RowHeight
        sngHeight = m_RowHeight
        If iRow = 7 And iCol = 5 Then
            sngWidth = m_ColWidth + m_ColWidth
        Else
            sngWidth = m_ColWidth
        End If
        shpBorder.Move sngLeft, sngTop, sngWidth, sngHeight
        blnShapeVisible = True
    End If
    shpBorder.Visible = blnShapeVisible
    
    If m_MouseDownButton = 1 Then
        If iCellIndex > 0 And iCellIndex < 5 Then
            If m_ButtonIndex < 0 Then
                iButtonIndex = iCellIndex - 1
                Call picButton_MouseDown(iButtonIndex, 1, Shift, X, Y)
            End If
        Else
            If m_ButtonIndex >= 0 Then
                Call picButton_MouseUp(m_ButtonIndex, 1, Shift, X, Y)
            End If
        End If
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCellIndex As Long
    Dim bCapture As Boolean
    
    iCellIndex = GetCellIndex(X, Y)
    bCapture = True
    m_MouseDownButton = 0
    If m_ButtonIndex >= 0 Then Call picButton_MouseUp(m_ButtonIndex, Button, Shift, X, Y)
    
    If Button = 1 And iCellIndex = m_MousePos Then
        If m_MousePos > 14 Then
            Unload Me
            Call SelectDate
            bCapture = False
        ElseIf m_MousePos = 5 Then
            Call lblYear_MouseUp(Button, Shift, X, Y)
        ElseIf m_MousePos = 6 Then
            Call lblMonth_MouseUp(Button, Shift, X, Y)
        End If
    End If
    If bCapture Then
        Call SetCapture(Me.hWnd)
    Else
        Call ReleaseCapture
    End If
    m_PopMenu = 0
    m_MousePos = 0
End Sub



Private Sub PaintDTPicker(iYear As Long, iMonth As Long)
    Const c_clrCurrentBack& = &HF8E9D3
    Const c_clrDarkSplit& = &H6A240A
    Const c_clrCurrentDay& = &HFF0000
    Const c_clrToday& = &HFF
    Const c_clrCurrentMonth& = &HCD895C
    Const c_clrOtherMonth& = &H808080
    Const c_clrGridLine& = &HEEEEEE
    Const c_BorderColor As Long = &HF8E9D3
    Dim clrText As Long
    Dim sOutText As String
    
    Dim sngWidth As Single, sngHeight As Single
    
    Dim d_firstDate As Date
    Dim iDays As Long
    Dim iWeekday As Long
    Dim YY As Long, MM As Long, DD As Long
    Dim d_Temp As Date, d_Today As Date, d_Current As Date
    
    Dim I As Long, j As Long
    Dim iRow As Long, iCol As Long
    Dim sngLeft As Single, sngTop As Single, sngRight As Single, sngBottom As Single
    Dim sngOffsetX As Single, sngOffsetY As Single
    Dim X As Single, Y As Single
    
    
    d_firstDate = GetFirstDate(iYear, iMonth)
    iWeekday = Weekday(d_firstDate, vbSunday) - 1
    d_firstDate = d_firstDate - iWeekday
    m_FirstDate = d_firstDate
    iDays = mc_GridRows& * mc_Cols - 2
    With Me
        sngWidth = .Width
        sngHeight = .Height
        
        .DrawMode = 13
        .Cls
        
        Me.Line (0, 0)-(sngWidth - 15, sngHeight - 15), c_BorderColor, B '画边框
        
        
        iRow = mc_Rows&
        sngLeft = m_FirstColX
        sngRight = m_FirstColX + m_ColWidth * mc_Cols&
        sngTop = m_FirstRowY + m_RowHeight + m_RowHeight
        sngBottom = m_FirstRowY + m_RowHeight * iRow
        
        Y = m_FirstRowY + m_RowHeight
        Me.Line (sngLeft, Y)-(sngRight, Y), c_clrDarkSplit&
        
        Y = sngTop
        For I = 2 To iRow
            Me.Line (sngLeft, Y)-(sngRight, Y), c_clrGridLine&
            Y = Y + m_RowHeight
        Next
        X = sngLeft
        For j = 0 To mc_Cols& - 2
            Me.Line (X, sngTop)-(X, sngBottom), c_clrGridLine&
            X = X + m_ColWidth
        Next
        Me.Line (X, sngTop)-(X, sngBottom - m_RowHeight), c_clrGridLine&
        X = X + m_ColWidth
        Me.Line (X, sngTop)-(X, sngBottom), c_clrGridLine&
        
        
        iRow = 1
        d_Today = Int(Now)
        d_Current = CDate(m_DefaultDate)
        d_Temp = d_firstDate
        Y = m_FirstRowY + m_RowHeight + (m_RowHeight - Me.TextHeight("1")) / 2
        X = m_FirstColX + (m_ColWidth - Me.TextHeight("日")) / 2
        .ForeColor = &H6A240A
        .CurrentY = Y
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H1010FF: Me.Print "日";
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H6A240A: Me.Print "一";
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H6A240A: Me.Print "二";
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H6A240A: Me.Print "三";
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H6A240A: Me.Print "四";
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H6A240A: Me.Print "五";
        .CurrentX = X:  X = X + m_ColWidth:  Me.ForeColor = &H1010FF: Me.Print "六";
        
        
        For I = 1 To iDays
            YY = Year(d_Temp)
            MM = Month(d_Temp) - 1
            DD = Day(d_Temp)
            sOutText = CStr(DD)
            
            iCol = I Mod mc_Cols
            If iCol = 0 Then
                iCol = mc_Cols
            ElseIf iCol = 1 Then
                Y = Y + m_RowHeight
                iRow = iRow + 1
            End If
            iCol = iCol - 1
            
            
            If d_Temp = d_Current Then
                sngTop = m_FirstRowY + m_RowHeight * iRow + 15
                sngBottom = sngTop + m_RowHeight - 30
                sngLeft = m_FirstColX + m_ColWidth * iCol + 15
                sngRight = sngLeft + m_ColWidth - 30
                Me.Line (sngLeft, sngTop)-(sngRight, sngBottom), c_clrCurrentBack&, BF
            End If
            If d_Temp = d_Today Then
                clrText = c_clrToday&
            Else
                If YY = iYear And MM = iMonth Then
                    clrText = c_clrCurrentMonth&
                Else
                    clrText = c_clrOtherMonth&
                End If
            End If
            .CurrentX = m_FirstColX + iCol * m_ColWidth + (m_ColWidth - Me.TextWidth(sOutText)) / 2
            .CurrentY = Y
            .ForeColor = clrText
            Me.Print sOutText
            
            d_Temp = d_Temp + 1
        Next
        
        Call PrintTodayButton
        
    End With
End Sub

Private Function GetCellIndex(ByVal X As Single, ByVal Y As Single) As Long
    Dim iCellIndex As Long
    
    Dim iRow  As Long, iCol As Long
    Dim YY As Long, XX As Long, w As Long, H As Long
    Dim bMouseOnCell As Boolean
    Dim I As Long
    
    If X <= 0 Or X >= Me.Width Or Y <= 0 Or Y >= Me.Height Then
        iCellIndex = -1
    Else
        For I = 0 To 3
            With picButton(I)
                If (X > .Left) And (X < (.Left + .Width)) Then
                    If (Y > .Top) And (Y < (.Top + .Height)) Then
                        iCellIndex = I + 1
                        Exit For
                    End If
                End If
            End With
        Next
        If iCellIndex = 0 Then
            With picBG(0)
                If (X > .Left) And (X < (.Left + .Width)) Then
                    If (Y > .Top) And (Y < (.Top + .Height)) Then iCellIndex = 5
                End If
            End With
        End If
        If iCellIndex = 0 Then
            With picBG(1)
                If (X > .Left) And (X < (.Left + .Width)) Then
                    If (Y > .Top) And (Y < (.Top + .Height)) Then iCellIndex = 6
                End If
            End With
        End If
        If iCellIndex = 0 Then
            YY = Y - m_FirstRowY
            XX = X - m_FirstColX
            w = m_ColWidth
            H = m_RowHeight
            If YY > 0 And YY < H * 8 Then
                iRow = YY \ H
                If iRow > 1 Then
                    If Abs(iRow * H - YY) > 15 And Abs(iRow * H + H - YY) > 15 Then
                        If XX > 0 And XX < w * 7 Then
                            iCol = XX \ w
                            If Abs(iCol * w - XX) > 15 And Abs(iCol * w + w - XX) > 15 Then
                                bMouseOnCell = True
                            Else
                                If iRow = 7 And Abs(XX - w * 6) <= 15 Then bMouseOnCell = True
                            End If
                        End If
                    End If
                End If
            End If
            If bMouseOnCell Then
                If iRow = 7 And iCol = 6 Then iCol = 5
                iCellIndex = iRow * 7 + iCol + 1
            End If
        End If
    End If
    GetCellIndex = iCellIndex
End Function
Private Sub PrintTodayButton(Optional ByVal bMouseDown As Boolean)
    Dim sOutText As String
    Dim sngLeft As Single, sngTop As Single, sngRight As Single, sngBottom As Single
    sOutText = "今天"
    With Me
        .DrawMode = 13
        
        sngLeft = m_FirstColX + m_ColWidth * 5 + 15
        sngRight = m_FirstColX + m_ColWidth * 7 - 15
        sngTop = m_FirstRowY + m_RowHeight * 7 + 15
        sngBottom = m_FirstRowY + m_RowHeight * 8 - 15
        Me.Line (sngLeft, sngTop)-(sngRight, sngBottom), &HF8E9D3, BF
        
        If bMouseDown Then
            .CurrentY = sngTop + (m_RowHeight - Me.TextHeight(sOutText)) / 2 + 15
        Else
            .CurrentY = sngTop + (m_RowHeight - Me.TextHeight(sOutText)) / 2
        End If
        .CurrentX = sngLeft + (m_ColWidth + m_ColWidth - Me.TextWidth(sOutText)) / 2
        .ForeColor = &H6A240A
        Me.Print sOutText
        
        .Refresh
    End With
End Sub

Private Function GetFirstDate(iYear As Long, iMonth As Long) As Date
    Dim iYearAdd As Long
    Dim iMonth2 As Long
    If iMonth <> 0 Then
        If iMonth < 0 Then
            iMonth2 = (iMonth Mod 12)
            If iMonth2 = 0 Then
                iYearAdd = iMonth \ 12
            Else
                iMonth2 = 12 + iMonth2
                iYearAdd = (iMonth - iMonth2) \ 12
            End If
        Else
            iMonth2 = iMonth Mod 12
            iYearAdd = (iMonth - iMonth2) \ 12
        End If
        iMonth = iMonth2
        iYear = iYear + iYearAdd
    End If
    GetFirstDate = CDate(CStr(iYear) & "-" & CStr(iMonth + 1) & "-1")
End Function

Private Sub Form_Activate()
    Dim dCurrentDate As Date
    Dim I As Long
    
    dCurrentDate = CurrentDate
    m_CurrentYear = Year(dCurrentDate)
    m_CurrentMonth = Month(dCurrentDate) - 1
    Call RepaintDTPicker
    
    m_ButtonIndex = -1
    
    For I = 0 To 3
        Call PaintButton(I)
    Next
End Sub

Private Sub PaintButton(ByVal Index As Long)
    Dim blnButtonDown As Boolean
    Dim iDir As Long
    If Index < 0 Then Exit Sub
    
    blnButtonDown = (m_ButtonIndex = Index)
    
    If Index Mod 2 = 0 Then iDir = 2 Else iDir = -2
    With picButton(Index)
        If blnButtonDown Then
            Call PaintRect(picButton(Index), 0, 0, .Width, .Height, iDir, blnButtonDown, &H808080, &H808080, .BackColor, .ForeColor)
        Else
            Call PaintRect(picButton(Index), 0, 0, .Width, .Height, iDir, blnButtonDown, &HFFFFFF, &HFFFFFF, .BackColor, .ForeColor)
        End If
    End With
End Sub

Private Sub RepaintDTPicker()
    lblYear.Caption = CStr(m_CurrentYear) & "年"
    lblMonth.Caption = CStr(m_CurrentMonth + 1) & "月"
    Call PaintDTPicker(m_CurrentYear, m_CurrentMonth)
End Sub
Private Sub Form_Load()
    m_Canceled = True
    If Not m_Inited Then CurrentDate = Now
    
    Call GetGridSize
End Sub

Private Sub GetGridSize()
    Dim sngWidth As Single, sngHeight As Single
    Dim iWidth As Long, iHeight As Long
    Dim iRowHeight As Long, iColWidth As Long
    
    sngWidth = Me.Width - 30
    sngHeight = Me.Height - 120
    
    iWidth = sngWidth / 15
    iHeight = sngHeight / 15
    m_ColWidth = Int(iWidth / mc_Cols&) * 15
    m_RowHeight = Int(iHeight / mc_Rows&) * 15
    
    m_FirstColX = CLng((sngWidth - m_ColWidth * mc_Cols&) / 30) * 15
    m_FirstRowY = CLng((sngHeight - m_RowHeight * mc_Rows&) / 30) * 15 + 90
End Sub





Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseCapture
End Sub

Private Sub lblMonth_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.PopupMenu popMonth, , picBG(1).Left, picBG(1).Top + picBG(1).Height - 30
    m_PopMenu = 2
End Sub

Private Sub lblYear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    
    mnuYear(5).Caption = m_CurrentYear & "年"
    For I = 4 To 0 Step -1
        mnuYear(I).Caption = CStr(m_CurrentYear - (5 - I)) & "年"
    Next
    For I = 6 To mnuYear.UBound Step 1
        mnuYear(I).Caption = CStr(m_CurrentYear + (I - 5)) & "年"
    Next
    Me.PopupMenu popYear, , picBG(0).Left, picBG(0).Top + picBG(0).Height - 30
    m_PopMenu = 1
End Sub

Public Sub PaintRect(oDC As Object, ByVal sngLeft As Single, ByVal sngTop As Single, _
                    ByVal sngRight As Single, ByVal sngBottom As Single, _
                    Optional ByVal iArrowDir_0None_1Up__1Down_2Left__2Right As Long, Optional ByVal bMouseDown As Boolean, _
                    Optional ByVal clrBorderDark As OLE_COLOR = &H404040, Optional ByVal clrBorderLight As OLE_COLOR = &HFFFFFF, _
                    Optional ByVal clrButtonBack As OLE_COLOR = &HC8D0D4, Optional ByVal clrButtonArrow As OLE_COLOR = &H404040)
    Const c_LineWidth# = 15
    Dim X1 As Single, Y1 As Single
    Dim X2 As Single, Y2 As Single

    Dim clrColorUp As Long, clrColorDown As Long
    Dim clrDCBack As OLE_COLOR

    Dim sngWidth As Single, sngHeight As Single

    Dim I As Long
    Dim iScaleWidth As Long, iScaleHeight As Long
    Dim iTrigonStep As Long, iTrigonSize As Long
    Dim fTrigonXPos As Single, fTrigonYPos As Single
    Dim iTrigonDir As Long

    On Error Resume Next
    '***********************************************
    '颜色处理
    If oDC.DrawMode = 7 Then
        clrDCBack = oDC.BackColor
        If Err.Number = 0 Then
            clrBorderDark = clrBorderDark Xor clrDCBack
            clrBorderLight = clrBorderLight Xor clrDCBack
            clrButtonBack = clrButtonBack Xor clrDCBack
            clrButtonArrow = clrButtonArrow Xor clrDCBack
        End If
    End If
    '***********************************************
    sngWidth = sngRight - sngLeft + c_LineWidth#
    sngHeight = sngBottom - sngTop + c_LineWidth#
    '******************************************
    If bMouseDown Then
        clrColorUp = clrBorderDark
        clrColorDown = clrBorderLight
    Else
        clrColorUp = clrBorderLight
        clrColorDown = clrBorderDark
    End If
    X1 = sngLeft
    Y1 = sngTop
    X2 = sngRight - c_LineWidth#
    Y2 = sngBottom - c_LineWidth#
    oDC.Line (X1, Y1)-(X2, Y2), clrButtonBack, BF
    '******************************************
    '绘制左边框
    X1 = sngLeft
    Y1 = sngTop
    X2 = X1
    Y2 = sngBottom - c_LineWidth#
    oDC.Line (X1, Y1)-(X2, Y2), clrColorUp
    '绘制上边框
    X1 = sngLeft + c_LineWidth#
    Y1 = sngTop
    X2 = sngRight
    Y2 = Y1
    oDC.Line (X1, Y1)-(X2, Y2), clrColorUp
    '绘制右边框
    X1 = sngRight - c_LineWidth#
    Y1 = sngTop + c_LineWidth#
    X2 = X1
    Y2 = sngBottom
    oDC.Line (X1, Y1)-(X2, Y2), clrColorDown
    '绘制下边框
    X1 = sngLeft
    Y1 = sngBottom - c_LineWidth#
    X2 = sngRight - c_LineWidth#
    Y2 = Y1
    oDC.Line (X1, Y1)-(X2, Y2), clrColorDown

    If iArrowDir_0None_1Up__1Down_2Left__2Right <> 0 Then '画三角形
        iScaleWidth = CLng(sngWidth / c_LineWidth#)
        iScaleHeight = CLng(sngHeight / c_LineWidth#)

        If iArrowDir_0None_1Up__1Down_2Left__2Right > 0 Then iTrigonDir = 1 Else iTrigonDir = -1

        If iArrowDir_0None_1Up__1Down_2Left__2Right = 1 Or iArrowDir_0None_1Up__1Down_2Left__2Right = -1 Then
            iTrigonSize = iScaleWidth \ 2
            fTrigonXPos = iTrigonSize * c_LineWidth# + sngLeft

            If iTrigonSize Mod 2 = 0 Then iTrigonSize = iTrigonSize - 1
            iTrigonSize = iTrigonSize - 2
            If iTrigonSize < 0 Then iTrigonSize = 1
            iTrigonStep = (iTrigonSize + 1) \ 2



            If iTrigonDir < 0 Then
                fTrigonYPos = sngBottom - ((iScaleHeight - iTrigonSize - 1) \ 4) * 3 * c_LineWidth#
            Else
                fTrigonYPos = sngTop + ((iScaleHeight - iTrigonSize - 1) \ 4) * 3 * c_LineWidth#
            End If

            For I = 0 To iTrigonStep - 1
                X1 = fTrigonXPos - I * c_LineWidth#
                X2 = fTrigonXPos + I * c_LineWidth#
                Y1 = fTrigonYPos + (I * (c_LineWidth#)) * iTrigonDir
                'Y2 = Y1 + c_LineWidth# * iTrigonDir
                oDC.Line (X1, Y1)-(X2, Y1), clrButtonArrow, BF
            Next
        Else
            iTrigonSize = iScaleHeight \ 2
            fTrigonYPos = iTrigonSize * c_LineWidth# + sngTop

            If iTrigonSize Mod 2 = 0 Then iTrigonSize = iTrigonSize - 1
            iTrigonSize = iTrigonSize - 2
            If iTrigonSize < 0 Then iTrigonSize = 1
            iTrigonStep = (iTrigonSize + 1) \ 2

            If iTrigonDir < 0 Then
                fTrigonXPos = sngRight - ((iScaleWidth - iTrigonSize - 1) \ 4) * 3 * c_LineWidth#
            Else
                fTrigonXPos = sngLeft + ((iScaleWidth - iTrigonSize - 1) \ 4) * 3 * c_LineWidth#
            End If

            For I = 0 To iTrigonStep - 1
                Y1 = fTrigonYPos - I * c_LineWidth#
                Y2 = fTrigonYPos + I * c_LineWidth#
                X1 = fTrigonXPos + (I * (c_LineWidth#)) * iTrigonDir
                'X2 = X1 + c_LineWidth# * iTrigonDir
                oDC.Line (X1, Y1)-(X1, Y2), clrButtonArrow, BF
            Next
        End If
    End If
End Sub


Private Sub mnuMonth_Click(Index As Integer)
    m_CurrentMonth = Index
    Call RepaintDTPicker
End Sub

Private Sub mnuYear_Click(Index As Integer)
    m_CurrentYear = Val(Replace(mnuYear(Index).Caption, "年", ""))
    Call RepaintDTPicker
End Sub



Private Sub picButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        m_ButtonIndex = Index
        Call PaintButton(Index)
        
        Call ButtonClick(Index)
        tmrMouseDown.Interval = 1000
        tmrMouseDown.Enabled = True
    End If
End Sub
Private Sub ButtonClick(ByVal Index As Integer)
    If Index = 0 Then
        m_CurrentYear = m_CurrentYear - 1
    ElseIf Index = 1 Then
        m_CurrentYear = m_CurrentYear + 1
    ElseIf Index = 2 Then
        If m_CurrentMonth = 0 Then
            m_CurrentYear = m_CurrentYear - 1
            m_CurrentMonth = 11
        Else
            m_CurrentMonth = m_CurrentMonth - 1
        End If
    Else
        If m_CurrentMonth = 11 Then
            m_CurrentYear = m_CurrentYear + 1
            m_CurrentMonth = 0
        Else
            m_CurrentMonth = m_CurrentMonth + 1
        End If
    End If
    Call RepaintDTPicker
End Sub

Private Sub picButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    If m_ButtonIndex >= 0 Then
        I = Index
        m_ButtonIndex = -1
        Call PaintButton(I)
    End If
    tmrMouseDown = False
End Sub


Private Sub tmrMouseDown_Timer()
    If m_ButtonIndex >= 0 Then
        Call ButtonClick(m_ButtonIndex)
        tmrMouseDown.Interval = 200
    Else
        tmrMouseDown.Enabled = False
    End If
End Sub


Private Sub GetClientSize(iLeft As Long, iTop As Long, iRight As Long, iBottom As Long, Optional ByVal bFullScreen As Boolean)
    Dim lpRect As rect
    Dim iScreenWidth As Long
    Dim iScreenHeight As Long
    
    iScreenWidth = Screen.Width / 15
    iScreenHeight = Screen.Height / 15
    If bFullScreen Then
        iLeft = 0
        iTop = 0
        iRight = iScreenWidth
        iBottom = iScreenHeight
    Else
        Call GetWindowRect(FindWindow("Shell_TrayWnd", ""), lpRect)
        If lpRect.Left <= 0 Then
            If lpRect.Top <= 0 Then
                If lpRect.Right >= iScreenWidth Then    '任务栏在顶部
                    iLeft = 0
                    iTop = lpRect.Bottom
                    iRight = iScreenWidth
                    iBottom = iScreenHeight
                Else '任务栏在左边
                    iLeft = lpRect.Right
                    iTop = 0
                    iRight = iScreenWidth
                    iBottom = iScreenHeight
                End If
            Else '任务栏靠下
                iLeft = 0
                iTop = 0
                iRight = iScreenWidth
                iBottom = lpRect.Top
            End If
        Else '任务栏靠右
            iLeft = 0
            iTop = 0
            iRight = lpRect.Left
            iBottom = iScreenHeight
        End If
    End If
End Sub


VERSION 5.00
Begin VB.UserControl ucDTPicker 
   BackColor       =   &H00F9894D&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   315
   ScaleWidth      =   1275
   Begin VB.PictureBox picDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   990
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F9894D&
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Text            =   "1899-12-30"
      Top             =   30
      Width           =   915
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   270
      Left            =   60
      Top             =   15
      Width           =   900
   End
End
Attribute VB_Name = "ucDTPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************************************
'编码：宋华
'*********************************************************************************

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'Default Property Values:
Event Change()
Event GetFocus()
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private m_blnListPage As Boolean
Private m_blnButtonMouseDown As Boolean
Private WithEvents m_frmList As frmDTPicker
Attribute m_frmList.VB_VarHelpID = -1
Private m_Value As Date

Private m_ListLeft As Single
Private m_ListTop As Single
Private m_ScreenLeft As Single
Private m_ScreenTop As Single

Private m_blnShowList As Boolean

Public Property Get Value() As Date
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Date)
    Dim strDate As String
    strDate = Format(New_Value, "YYYY-MM-DD")
    m_Value = CDate(strDate)
    txtInput.Text = strDate
    PropertyChanged "Value"
End Property




Private Sub m_frmList_OnLoaded()
    m_blnShowList = True
End Sub

Private Sub m_frmList_OnUnload()
    m_blnShowList = False
End Sub

Private Sub m_frmList_SelectDate(ByVal newDate As Date)
    txtInput.Text = Format(newDate, "YYYY-MM-DD")
    If newDate <> m_Value Then
        m_Value = newDate
        RaiseEvent Change
    End If
End Sub

Private Sub picDown_Click()
    Dim iPos As Long, iLen As Long
    
    m_blnButtonMouseDown = False
    Call PaintButton
    Call ShowList
    
    iPos = txtInput.SelStart
    iLen = txtInput.SelLength
    txtInput.SetFocus
    txtInput.SelStart = iPos
    txtInput.SelLength = txtInput.SelLength
End Sub
Public Sub ShowList()
    Dim ControlPos As POINTAPI
    Dim sngWidth As Single, sngHeight As Single
    Dim sngLeft As Single, sngTop As Single
    
    If m_blnListPage Then Exit Sub
    
    On Error Resume Next
        
    ClientToScreen UserControl.hWnd, ControlPos
    sngLeft = (ControlPos.X) * 15#
    sngTop = (ControlPos.Y) * 15#
    
    sngHeight = UserControl.Height
    sngWidth = UserControl.Width
    
    Call m_frmList.ShowList(sngLeft, sngTop, sngWidth, sngHeight, m_Value)
End Sub



Private Sub picDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_blnButtonMouseDown = True
    Call PaintButton
End Sub

Private Sub picDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_blnButtonMouseDown = False
    Call PaintButton
End Sub



Private Sub txtInput_Click()
    RaiseEvent Click
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    If Not m_blnShowList Then
        RaiseEvent GetFocus
    End If
End Sub



Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyEscape Then
        If m_blnShowList Then Unload m_frmList
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtInput_LostFocus()
    Dim strDate As String
    Dim newDate As Date
    If m_blnShowList Then Unload m_frmList
    strDate = FormatDate(Trim(txtInput.Text))
    If strDate = "" Then
        newDate = m_Value
    Else
        newDate = CDate(strDate)
    End If
    txtInput.Text = Format(newDate, "YYYY-MM-DD")
    If newDate <> m_Value Then
        m_Value = newDate
        RaiseEvent Change
    End If
End Sub

Private Sub txtInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtInput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtInput_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub



Private Sub UserControl_Initialize()
    On Error Resume Next
    Set m_frmList = New frmDTPicker
End Sub

Private Sub UserControl_InitProperties()
    Me.Value = CDate(Format(Now, "YYYY-MM-DD"))
End Sub

'Initialize Properties for User Control
Private Sub UserControl_Resize()
    Const c_BorderWidth# = 15
    Dim sngTop As Single
    Dim sngHeight As Single
    Dim sngWidth As Single
    Dim sngButtonSize As Single
    Dim sngButtonLeft As Single
    Dim sngTextHeight As Single
    
    
    sngHeight = UserControl.Height
    sngWidth = UserControl.Width
    
    If sngHeight < 120 Then
        UserControl.Height = 120
        sngHeight = 120
    End If
    If sngWidth < 1200 Then
        UserControl.Width = 1200
        sngWidth = 1200
    End If
    
    sngHeight = sngHeight - c_BorderWidth# - c_BorderWidth#
    sngWidth = sngWidth - c_BorderWidth# - c_BorderWidth#
    
    sngButtonSize = sngHeight
    sngButtonLeft = sngWidth + c_BorderWidth# - sngButtonSize
    picDown.Move sngButtonLeft, c_BorderWidth#, sngButtonSize, sngButtonSize
    Call PaintButton
    
    Call shpBack.Move(c_BorderWidth#, c_BorderWidth#, sngButtonLeft - c_BorderWidth#, sngHeight)
    Call SetTextPos
End Sub
Private Sub SetTextPos()
    Dim sngWidth As Single, sngHeight As Single
    Dim sngLeft As Single, sngTop As Single
    Dim sngTextHeight As Single
    Dim sngTextWidth As Single
    
    With shpBack
        sngWidth = .Width
        sngHeight = .Height
        sngLeft = .Left
        sngTop = .Top
    End With
    
    sngTextHeight = UserControl.TextHeight("0") + 30
    If sngTextHeight > sngHeight Then sngTextHeight = sngHeight
    sngTop = sngTop + (sngHeight - sngTextHeight) / 2
    sngTextWidth = UserControl.TextWidth(Format(m_Value, "YYYY-MM-DD"))
    If sngTextWidth + 60 < sngWidth Then
        sngWidth = sngWidth - 60
        sngLeft = sngLeft + 30
    End If
    txtInput.Move sngLeft, sngTop, sngWidth, sngTextHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Value = PropBag.ReadProperty("Value", Now)
    Me.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txtInput.ForeColor = PropBag.ReadProperty("ForeColor", &HF9894D)
    UserControl.BackColor = PropBag.ReadProperty("BorderColor", &HF9894D)
    
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    
    Set txtInput.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = txtInput.Font
    txtInput.Alignment = PropBag.ReadProperty("Alignment", 0)
    
    Call UserControl_Resize
End Sub


Private Sub UserControl_Terminate()
    On Error Resume Next
    If Not m_frmList Is Nothing Then
        Unload m_frmList
        Set m_frmList = Nothing
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, #1/1/2010#)
    Call PropBag.WriteProperty("BackColor", txtInput.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", txtInput.ForeColor, &HF9894D)
    Call PropBag.WriteProperty("BorderColor", UserControl.BackColor, &HF9894D)
    
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    
    Call PropBag.WriteProperty("Font", txtInput.Font, Ambient.Font)
    Call PropBag.WriteProperty("Alignment", txtInput.Alignment, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtInput,txtInput,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = txtInput.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtInput.BackColor() = New_BackColor
    shpBack.BackColor = New_BackColor
    shpBack.BorderColor = New_BackColor
    PropertyChanged "BackColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtInput,txtInput,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtInput.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtInput.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtInput,txtInput,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    txtInput.Enabled() = New_Enabled
    picDown.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtInput,txtInput,-1,Font
Public Property Get Font() As Font
    Set Font = txtInput.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtInput.Font = New_Font
    Set UserControl.Font = New_Font
    Call SetTextPos
    PropertyChanged "Font"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtInput,txtInput,-1,Alignment
Public Property Get Alignment() As Integer
    Alignment = txtInput.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    txtInput.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = UserControl.BackColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    UserControl.BackColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Private Sub PaintButton()
    If m_blnButtonMouseDown Then
        Call PaintRect(picDown, 0, 0, picDown.Width, picDown.Height, -1, m_blnButtonMouseDown, &H808080, &H808080, picDown.BackColor, picDown.ForeColor)
    Else
        Call PaintRect(picDown, 0, 0, picDown.Width, picDown.Height, -1, m_blnButtonMouseDown, &HFFFFFF, &HFFFFFF, picDown.BackColor, picDown.ForeColor)
    End If
End Sub

'************************************************************
'格式化日期字符串
Private Function FormatDate(ByVal strDate) As String
    Dim ret As String
    Dim aDate As Date
    Dim strDefYear As String
    Dim YY As String, MM As String, DD As String
    Dim iLen As Long
    
    
    strDate = StrConv(strDate, vbNarrow)
    strDate = Replace(Replace(Replace(Replace(strDate, "/", "-"), ".", "-"), ",", "-"), " ", "")
    
    If IsDate(strDate) Then
        aDate = CDate(strDate)
        ret = Format(aDate, "YYYY-MM-DD")
    Else
        iLen = Len(strDate)
        If iLen >= 3 Then
            If IsNumeric(strDate) Then
                strDefYear = CStr(Year(Now))
                DD = Right(strDate, 2)
                Select Case iLen
                Case 3
                    YY = strDefYear
                    MM = Left(strDate, 1)
                Case 4
                    YY = strDefYear
                    MM = Left(strDate, 2)
                Case 6
                    YY = Left(strDate, 2)
                    MM = Mid(strDate, 3, 2)
                Case 8
                    YY = Left(strDate, 4)
                    MM = Mid(strDate, 5, 2)
                End Select
                strDate = YY & "-" & MM & "-" & DD
                If IsDate(strDate) Then
                    aDate = CDate(strDate)
                    ret = Format(aDate, "YYYY-MM-DD")
                End If
            End If
        End If
    End If
    FormatDate = ret
End Function

Private Sub PaintRect(oDC As Object, ByVal sngLeft As Single, ByVal sngTop As Single, _
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




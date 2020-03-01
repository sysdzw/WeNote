Attribute VB_Name = "modRunStartup"
Option Explicit
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long    ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Const HKEY_CURRENT_USER = &H80000001 '常量宣告
Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const REG_BINARY = 3

Public Sub setAutoRun()
    Dim fname$
    If Right(App.Path, 1) = "\" Then
        fname = App.Path
    Else
        fname = App.Path & "\"
    End If
    SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "WeNote", fname & App.EXEName & ".exe "
'    MsgBox "建立开机启动成功！"
End Sub

Public Sub cancelAutoRun()
    Dim hKey&, subKey$
    subKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    RegOpenKey HKEY_CURRENT_USER, subKey, hKey
    RegDeleteValue hKey, "WeNote"    ' 删除键值
    RegCloseKey hKey    ' 关闭句柄
End Sub
Public Function isHasSetAutoRun() As Boolean
    isHasSetAutoRun = (GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "WeNote") <> "")
End Function

Private Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim keyHand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyHand)
    r = RegSetValueEx(keyHand, strValue, 0, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)))
    r = RegCloseKey(keyHand)
End Sub

Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim ret
    'Open the key
    RegOpenKey hKey, strPath, ret
    'Get the key's content
    GetString = RegQueryStringValue(ret, strValue)
    'Close the key
    RegCloseKey ret
End Function

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function


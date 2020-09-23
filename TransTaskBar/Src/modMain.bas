Attribute VB_Name = "modMain"
Option Explicit

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const ERROR_SUCCESS = 0&
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
    KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Function prvGetLevel() As Byte
    Dim LnLevel As Byte
    Dim LsCommand As String
    Dim LnToken As Integer
    Dim LsKey As String
    Dim LsInput As String
    
    LsCommand = Command()
    LsKey = "/translevel:"
    LnToken = InStr(1, LsCommand, LsKey, vbTextCompare)
    If (LnToken = 0) Then
        LsInput = InputBox("Enter transparency level for task bar" & vbCrLf & "1 to 255", , 100)
    Else
        Dim LnEnd As Integer
        
        LnToken = (LnToken + Len(LsKey))
        LsInput = Mid$(LsCommand, LnToken, 3)
    End If
    If (Trim$(LsInput) = 0) Then
        LnLevel = 100
    Else
        LnLevel = Val(Left$(LsInput, 3))
    End If
    If (LnLevel > 255) Then LnLevel = 255
    If (LnLevel < 50) Then LnLevel = 50
    prvGetLevel = LnLevel
End Function

Private Sub prvMakeTransparent(LhWnd As Long, bLevel As Byte)
    Dim lOldStyle As Long
    lOldStyle = GetWindowLong(LhWnd, GWL_EXSTYLE)
    SetWindowLong LhWnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes LhWnd, 0, bLevel, LWA_ALPHA
End Sub

Public Sub Main()
    Dim LhWnd As Long
    Dim LnLevel As Byte

    LnLevel = prvGetLevel
    If (InStr(1, Command(), "/silent", vbTextCompare) = 0) Then
        If SetAutoStart(LnLevel) Then
            MsgBox "TransTaskBar will be loaded when OS starts.", vbOKOnly Or vbInformation
        End If
    Else
    End If
    LhWnd = FindWindow("Shell_TrayWnd", vbNullString)
    If (LhWnd <> 0) Then
        prvMakeTransparent LhWnd, LnLevel
    End If
End Sub

Public Function SetAutoStart(nLevel As Byte) As Boolean
    Dim nRet As Long
    Dim hKey As Long
    Dim nResult As Long
    Dim LsFullPath As String
   
    With App
        LsFullPath = App.Path & "\" & App.EXEName & ".exe"
    End With
    If (InStr(1, LsFullPath, " ") > 0) Then
        LsFullPath = """" & LsFullPath & """"
    End If
    LsFullPath = LsFullPath & " /silent /TransLevel:" & CStr(nLevel)
    ' Open (or create and open) key
    nRet = RegCreateKeyEx(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Run", 0&, vbNullString, _
        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
    If nRet = ERROR_SUCCESS Then
    ' Write new value to registry
        nRet = RegSetValueEx(hKey, App.EXEName, 0&, 1&, ByVal LsFullPath, Len(LsFullPath))
        Call RegCloseKey(hKey)
    End If
    SetAutoStart = (nRet = ERROR_SUCCESS)
End Function



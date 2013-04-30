Attribute VB_Name = "API"
' создание файла
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

' создание архивной копии БД
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

' запуск браузера и почтовой программы
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' звук
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_APPLICATION = &H80
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000

' перемещение окна и анимация кнопок
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long

' перетаскивание
Dim ClickBool As Boolean
Dim Xs%, Ys%

Sub MInit()
 ClickBool = False
 Xs = 0
 Ys = 0
End Sub

Sub MMove(ByVal Handle As Long, ByVal x%, ByVal y%)
    Dim R As RECT
    If ClickBool Then
        Call GetWindowRect(Handle, R)
        W% = R.Right - R.Left
        H% = R.Bottom - R.Top
        x = R.Left + (x - Xs) / Screen.TwipsPerPixelX
        y = R.Top + (y - Ys) / Screen.TwipsPerPixelY
        Call MoveWindow(Handle, x, y, W, H, True)
    End If
End Sub

Sub MDown(ByVal x%, ByVal y%)
 ClickBool = True
 Xs = x
 Ys = y
End Sub

Sub MUp()
 ClickBool = False
End Sub

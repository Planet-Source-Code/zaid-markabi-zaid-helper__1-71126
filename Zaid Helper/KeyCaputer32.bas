Attribute VB_Name = "KeyCaputer32"
Public Const DT_CENTER = &H1
Public Const DT_WORDBREAK = &H10
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Global Cnt As Long, sSave As String, sOld As String, Ret As String
Dim Tel As Long
Function GetPressedKey() As String
   For Cnt = 32 To 128
       'Get the keystate of a specified key
       If GetAsyncKeyState(Cnt) <> 0 Then
           GetPressedKey = Chr$(Cnt)
           Exit For
       End If
   Next Cnt
End Function
Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
   Ret = GetPressedKey
   If Ret <> sOld Then
       sOld = Ret
       sSave = sSave + sOld
   End If
End Sub


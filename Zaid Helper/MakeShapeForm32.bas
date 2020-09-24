Attribute VB_Name = "MakeShapeForm32"
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Function fMakeATranspArea(AreaType As String, pCordinate() As Long, hWnd_Form As Long, Width As Single, Height As Single) As Boolean
 Const RGN_DIFF = 4
 Dim lOriginalForm As Long
 Dim ltheHole As Long
 Dim lNewForm As Long
 Dim lFwidth As Single
 Dim lFHeight As Single
 Dim lborder_width As Single
 Dim ltitle_height As Single

  On Error GoTo Trap
    lFwidth = (Width)
    lFHeight = (Height)
    lOriginalForm = CreateRectRgn(0, 0, lFwidth, lFHeight)
    lborder_width = (lFHeight - ScaleWidth) / 2
    ltitle_height = lFHeight - lborder_width - ScaleHeight
  Select Case AreaType
    Case "Elliptic"
      ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
    Case "RectAngle"
      ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
    Case "RoundRect"
      ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))
    Case "Circle"
      ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))
    Case Else
      MsgBox "Unknown Shape!!"
      Exit Function
    End Select
    lNewForm = CreateRectRgn(0, 0, 0, 0)
    CombineRgn lNewForm, lOriginalForm, ltheHole, RGN_DIFF
    SetWindowRgn hWnd_Form, lNewForm, True
    fMakeATranspArea = True
    Exit Function
Trap:
End Function

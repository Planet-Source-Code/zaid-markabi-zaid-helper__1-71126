VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.UserControl ZaidHelper 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2910
   ScaleHeight     =   2490
   ScaleWidth      =   2910
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   2280
      Top             =   1920
   End
   Begin VB.PictureBox BackClr 
      Height          =   1695
      Left            =   0
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.Timer TimerAnimatedText 
         Left            =   2040
         Top             =   1200
      End
      Begin VB.PictureBox Picture1 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   1035
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         Begin MCI.MMControl MMControl1 
            Height          =   330
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   582
            _Version        =   393216
            PrevVisible     =   0   'False
            NextVisible     =   0   'False
            PauseVisible    =   0   'False
            BackVisible     =   0   'False
            StepVisible     =   0   'False
            RecordVisible   =   0   'False
            EjectVisible    =   0   'False
            DeviceType      =   ""
            FileName        =   ""
         End
         Begin VB.Image StyleCmd 
            Height          =   270
            Index           =   4
            Left            =   360
            Picture         =   "ZaidHelper.ctx":0000
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image StyleCmd 
            Height          =   270
            Index           =   3
            Left            =   240
            Picture         =   "ZaidHelper.ctx":3498
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image StyleCmd 
            Height          =   270
            Index           =   2
            Left            =   120
            Picture         =   "ZaidHelper.ctx":8638
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
         Begin VB.Image StyleCmd 
            Height          =   270
            Index           =   1
            Left            =   0
            Picture         =   "ZaidHelper.ctx":C544
            Stretch         =   -1  'True
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Label LblTxt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   1695
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Label LblCmnd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image CommandBut 
      Height          =   495
      Left            =   120
      Picture         =   "ZaidHelper.ctx":F9C8
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   1845
   End
End
Attribute VB_Name = "ZaidHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event Click() 'MappingInfo=LblCmnd,LblCmnd,-1,Click
'Event Click() 'MappingInfo=CommandBut,CommandBut,-1,Click


Dim storer() As Byte
Dim data As String
Dim Length As Byte
Dim things(3) As Byte
Dim binary(7) As Byte

Sub SplitFile(FilePath As String, OutputFolder As String, BytsSize As Long)
Dim byt() As Byte
ReDim byt(BytsSize - 1)
Dim Pos As Long
Pos = 1
Dim FilesNumber As Integer
FilesNumber = 0

Open FilePath For Binary As #1
Do While EOF(1) = False
FilesNumber = FilesNumber + 1

If Pos + BytsSize < LOF(1) Then
Get #1, Pos, byt()
Open OutputFolder + "\File_" + Format(FilesNumber, "0000000") For Binary As #2
Put #2, 1, byt()
Close #2
Pos = Pos + BytsSize
Else
ReDim byt(LOF(1) - Pos)
Get #1, Pos, byt()
Open OutputFolder + "\File_" + Format(FilesNumber, "0000000") For Binary As #2
Put #2, 1, byt()
Close #2
Exit Do
End If

Loop
Close #1

End Sub

Sub CombFile(OutputtedFolder As String, FilePath As String)
Open FilePath For Binary As #1
On Error GoTo 1
Dim byt() As Byte
Dim FilesNumber As Long
For FilesNumber = 1 To 9999999
Open OutputtedFolder + "\File_" + Format(FilesNumber, "0000000") For Binary As #2
ReDim byt(LOF(2) - 1)
Get #2, 1, byt()
Put #1, , byt()
Close #2
Next
1:
Close #1
End Sub

Sub ZipFile(FilePath As String, ZippedFilePath As String)
Dim bob As String
Dim jim1 As Byte
Dim jim2 As Byte
Dim bin() As Byte
Dim counter(255) As Long
Dim storearr(255) As Long
Dim placearr(255) As Integer
Dim lngcounter As Long
Dim lngcounter2 As Long
Dim lngcounter3 As Long
bob = FilePath
data = ""
Open bob For Binary Access Read As #1
ReDim bin(LOF(1) - 1)
Get #1, , bin
Close #1
For lngcounter = 0 To UBound(bin)
counter(bin(lngcounter)) = counter(bin(lngcounter)) + 1
Next lngcounter
storearr(0) = counter(0)
placearr(0) = 0
For lngcounter = 1 To 255
For lngcounter2 = 0 To 255
If counter(lngcounter) > storearr(lngcounter2) Or counter(lngcounter) = storearr(lngcounter2) Then
For lngcounter3 = 254 To lngcounter2 Step -1
storearr(lngcounter3 + 1) = storearr(lngcounter3)
placearr(lngcounter3 + 1) = placearr(lngcounter3)
Next lngcounter3
storearr(lngcounter2) = counter(lngcounter)
placearr(lngcounter2) = lngcounter
Exit For
End If
Next lngcounter2
Next lngcounter
For lngcounter = 0 To 3
things(lngcounter) = placearr(lngcounter)
Next lngcounter
jim1 = bin(0)
If jim1 = things(0) Then
ReDim Preserve storer(2)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 0
storer(UBound(storer)) = 0
ElseIf jim1 = things(1) Then
ReDim Preserve storer(2)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 0
storer(UBound(storer)) = 1
ElseIf jim1 = things(2) Then
ReDim Preserve storer(2)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 1
storer(UBound(storer)) = 0
ElseIf jim1 = things(3) Then
ReDim Preserve storer(2)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 1
storer(UBound(storer)) = 1
Else
ReDim Preserve storer(8)
storer(UBound(storer) - 8) = 0
Call binaryconv(jim1, 0)
For lngcounter2 = 0 To 7
storer(UBound(storer) - 7 + lngcounter2) = binary(lngcounter2)
Next lngcounter2
End If
For lngcounter = 1 To UBound(bin)
jim1 = bin(lngcounter)
If jim1 = things(0) Then
ReDim Preserve storer(UBound(storer) + 3)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 0
storer(UBound(storer)) = 0
ElseIf jim1 = things(1) Then
ReDim Preserve storer(UBound(storer) + 3)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 0
storer(UBound(storer)) = 1
ElseIf jim1 = things(2) Then
ReDim Preserve storer(UBound(storer) + 3)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 1
storer(UBound(storer)) = 0
ElseIf jim1 = things(3) Then
ReDim Preserve storer(UBound(storer) + 3)
storer(UBound(storer) - 2) = 1
storer(UBound(storer) - 1) = 1
storer(UBound(storer)) = 1
Else
ReDim Preserve storer(UBound(storer) + 9)
storer(UBound(storer) - 8) = 0
Call binaryconv(jim1, 0)
For lngcounter2 = 0 To 7
storer(UBound(storer) - 7 + lngcounter2) = binary(lngcounter2)
Next lngcounter2
End If
If UBound(storer) > 7 Then
For lngcounter2 = 0 To 7
binary(lngcounter2) = storer(lngcounter2)
Next lngcounter2
data = data & CharFind(binary)
For lngcounter2 = 8 To UBound(storer)
storer(lngcounter2 - 8) = storer(lngcounter2)
Next lngcounter2
ReDim Preserve storer(UBound(storer) - 8)
End If
Next lngcounter
Length = (UBound(storer) + 1) Mod 8
Length = 8 - Length
ReDim Preserve storer(UBound(storer) + Length)
For lngcounter = 0 To UBound(storer) Step 8
For lngcounter2 = 0 To 7
binary(lngcounter2) = storer(lngcounter + lngcounter2)
Next lngcounter2
data = data & CharFind(binary)
Next lngcounter
Call PrintToFile(ZippedFilePath)
End Sub

Public Sub binaryconv(number As Byte, counter As Long)
For lngcounter = 1 To 8

If number - (2 ^ (8 - lngcounter)) > -1 Then
number = number - (2 ^ (8 - lngcounter))
binary(counter * 8 + (lngcounter - 1)) = 1
Else
binary(counter * 8 + (lngcounter - 1)) = 0
End If
Next
End Sub

Public Function CharFind(bob() As Byte) As String
For lngcounter = 0 To 7

If bob(lngcounter) = 1 Then
number = number + (2 ^ (7 - lngcounter))
End If
Next
CharFind = Chr(number)
End Function

Public Sub PrintToFile(ZippedFilePath As String)
Open ZippedFilePath For Binary Access Write As #1
For lngcounter = 0 To 3
Put #1, , things(lngcounter)
Next lngcounter
Put #1, , Length
Put #1, , data
Close #1
End Sub

Sub UnZipFile(ZippedFilePath As String, OutputFilePath As String)
Dim origdata() As Byte
Dim num As Long
Dim newdata() As Byte
Dim bin(7) As Byte
Dim origd As Long
bob = ZippedFilePath
Open bob For Binary Access Read As #1
For lngcounter = 0 To 3
Get #1, , things(lngcounter)
Next lngcounter
Get #1, , Length
ReDim origdata(LOF(1) - 6)
Get #1, , origdata
Close #1
ReDim storer(15)
For lngcounter = 0 To 1

Call binaryconv(origdata(lngcounter), 0)
For lngcounter2 = 0 To 7
storer(lngcounter * 8 + lngcounter2) = binary(lngcounter2)
Next lngcounter2
Next lngcounter
origd = 1
If (UBound(origdata) + 1) Mod 2 = 1 Then
ReDim Preserve storer(UBound(storer) + 8)
origd = origd + 1
Call binaryconv(origdata(origd), 0)
For lngcounter2 = 0 To 7
storer(15 + lngcounter2) = binary(lngcounter2)
Next lngcounter2
End If
num = 0
ReDim newdata(0)
Do Until (num > UBound(storer))

If storer(num) = 1 Then

If storer(num + 1) = 0 And storer(num + 2) = 0 Then

newdata(UBound(newdata)) = things(0)
ReDim Preserve newdata(UBound(newdata) + 1)
ElseIf storer(num + 1) = 0 And storer(num + 2) = 1 Then

newdata(UBound(newdata)) = things(1)
ReDim Preserve newdata(UBound(newdata) + 1)
ElseIf storer(num + 1) = 1 And storer(num + 2) = 0 Then

newdata(UBound(newdata)) = things(2)
ReDim Preserve newdata(UBound(newdata) + 1)
ElseIf storer(num + 1) = 1 And storer(num + 2) = 1 Then

newdata(UBound(newdata)) = things(3)
ReDim Preserve newdata(UBound(newdata) + 1)
End If
num = num + 3
Else
For lngcounter = 0 To 7

bin(lngcounter) = storer(num + 1 + lngcounter)
Next lngcounter
newdata(UBound(newdata)) = Asc(CharFind(bin))
ReDim Preserve newdata(UBound(newdata) + 1)
num = num + 9
End If
If origd < UBound(origdata) Then
origd = origd + 2
For lngcounter = (origd - 1) To origd
Call binaryconv(origdata(lngcounter), 0)
For lngcounter2 = 0 To 7
ReDim Preserve storer(UBound(storer) + 1)
storer(UBound(storer)) = binary(lngcounter2)
Next lngcounter2
Next lngcounter
ElseIf origd = UBound(origdata) Then
origd = origd + 10
ReDim Preserve storer(UBound(storer) - Length)
End If
Loop
Kill bob
Open OutputFilePath For Append As #1
Close #1
Open OutputFilePath For Binary Access Write As #1
ReDim Preserve newdata(UBound(newdata) - 1)
Put #1, , newdata
Close #1
End Sub







Sub DrawTime(TimeFormat As String)
LblTxt.Caption = Format(Time, TimeFormat)
End Sub

Sub DrawDate(DateFormat As String)
LblTxt.Caption = Format(Date, DateFormat)
End Sub

Sub DrawAnimatedText(Text As String, TimerSpeed As Integer)
strTextAnimated = Text
strTextAnimated = Space(Len(strTextAnimated)) & strTextAnimated
TimerAnimatedText.Interval = TimerSpeed
End Sub

Sub DrawDayName()
Dim Dday As Integer
Dday = Weekday(Date)
If Dday = 1 Then LblTxt.Caption = "Sat."
If Dday = 2 Then LblTxt.Caption = "Sun."
If Dday = 3 Then LblTxt.Caption = "Mon."
If Dday = 4 Then LblTxt.Caption = "Tus."
If Dday = 5 Then LblTxt.Caption = "Wed."
If Dday = 6 Then LblTxt.Caption = "Tus."
If Dday = 7 Then LblTxt.Caption = "Fri."
End Sub

Sub DrawMmonthName()
Mmonth = Mid(Date, 4, 2)
LblTxt.Caption = MonthName(Mmonth)
End Sub

Sub PlayVideoFile(FilePath As String)
MMControl1.Command = "stop"
MMControl1.Command = "prev"
MMControl1.FileName = (FilePath)
MMControl1.Command = "open"
MMControl1.hWndDisplay = BackClr.hwnd
MMControl1.Command = "play"
End Sub

Sub DeleteFile(FilePath As String)
On Error Resume Next
Kill (FilePath)
End Sub

Sub CopyFile(FileFrom As String, FileTo As String)
On Error Resume Next
FileCopy FileFrom, FileTo
End Sub

Sub MoveFile(FileFrom As String, FileTo As String)
On Error Resume Next
Name FileFrom As FileTo
End Sub

Sub OpenWebPage(URL As String, Mode As VbAppWinStyle)
Shell "RUNDLL32.EXE URL.DLL,FileProtocolHandler " + URL, Mode
End Sub

Sub SetVolume(Vol_0_To_65536 As Integer)
If Vol_0_To_65536 < 0 Then
Vol_0_To_65536 = 0
End If
If Vol_0_To_65536 > 65536 Then
Vol_0_To_65536 = 65536
End If
Dim Vol&
Vol = CLng("&H" & Hex(Vol_0_To_65536 + 65536))
waveOutSetVolume 0, Vol
End Sub

Sub CreateFolder(FolderPath As String)
Dim attr As SECURITY_ATTRIBUTES
Dim rval As Long
attr.nLength = Len(attr)
attr.lpSecurityDescriptor = 0
attr.bInheritHandle = 1
If Not Right(FolderPath, 1) = "\" Then
FolderPath = FolderPath + "\"
End If
rval = CreateDirectory(FolderPath, attr)
End Sub

Sub ChangeDriveName(Drive As String, Label As String)
Dim rval As Long
If Not Right(Drive, 1) = "\" Then
Drive = Drive + ":\"
End If
rval = SetVolumeLabel(Drive, Label)
End Sub

Sub CDopenDoor(Drive As String, Label As String)
Call mciSendString("Set CDAudio Door Open", 0&, 0&, 0&)
End Sub

Sub CDcloseDoor(Drive As String, Label As String)
Call mciSendString("Set CDAudio Door Closed", 0&, 0&, 0&)
End Sub

Sub SetWindowsScreenSize(Width As String, Height As String)
Dim typDevM As typDevMODE
Dim lngResult As Long
Dim intAns    As Integer
lngResult = EnumDisplaySettings(0, 0, typDevM)
With typDevM
   .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
   .dmPelsWidth = Width
   .dmPelsHeight = Height
End With
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
End Sub

Sub SetWindowsEnableMouseKeyboard(Enable As Boolean)
BlockInput Enable
End Sub

Sub GetWindowsScreen(Left As Single, Top As Single, Width As Single, Height As Single)
Set UserControl.BackClr.Picture = CaptureScreen(Left, Top, Width, Height)
End Sub

Sub GetWindowsWallpaper()
PaintDesktop UserControl.BackClr.hDC
End Sub

Sub SetWindowsCursorEnable(Enable As Boolean)
ShowCursor (Enable)
End Sub

Sub Sleep(Seconds As Integer)
TempTime = DateAdd("s", Seconds, Now)
While TempTime > Now
DoEvents
Wend
End Sub

Sub SetFileMode(FilePath As String, Mode As VbFileAttribute)
SetAttr FilePath, Mode
End Sub

Sub SetWindowsCursorPos(X As Single, Y As Single)
Dim P As POINTAPI
P.X = X
P.Y = Y
Ret = SetCursorPos(P.X, P.Y)
End Sub

Sub SetWindowsEmptyRecycleBin()
SHEmptyRecycleBin UserControl.hwnd, vbNullString, 0
SHUpdateRecycleBinIcon
End Sub

Sub SetWindowShapeElliptic(hWnd_Form As Long, Width_From As Single, Height_Form As Single, Param1 As Integer, Param2 As Integer, Param3 As Integer, Param4 As Integer)
 Dim lParam(1 To 6) As Long
 lParam(1) = Param1
 lParam(2) = Param2
 lParam(3) = Param3
 lParam(4) = Param4
 lParam(5) = 0
 lParam(6) = 0
 Call fMakeATranspArea("Elliptic", lParam(), hWnd_Form, Width, Height_Form)
End Sub

Sub SetWindowShapeCircle(hWnd_Form As Long, Width_From As Single, Height_Form As Single, Param1 As Integer, Param2 As Integer, Param3 As Integer, Param4 As Integer)
 Dim lParam(1 To 6) As Long
 lParam(1) = Param1
 lParam(2) = Param2
 lParam(3) = Param3
 lParam(4) = Param4
 lParam(5) = 0
 lParam(6) = 0
 Call fMakeATranspArea("Circle", lParam(), hWnd_Form, Width, Height_Form)
End Sub

Sub SetWindowShapeRectAngle(hWnd_Form As Long, Width_From As Single, Height_Form As Single, Param1 As Integer, Param2 As Integer, Param3 As Integer, Param4 As Integer)
 Dim lParam(1 To 6) As Long
 lParam(1) = Param1
 lParam(2) = Param2
 lParam(3) = Param3
 lParam(4) = Param4
 lParam(5) = 0
 lParam(6) = 0
 Call fMakeATranspArea("RectAngle", lParam(), hWnd_Form, Width, Height_Form)
End Sub

Sub SetWindowShapeRoundRect(hWnd_Form As Long, Width_From As Single, Height_Form As Single, Param1 As Integer, Param2 As Integer, Param3 As Integer, Param4 As Integer, Param5 As Integer, Param6 As Integer)
 Dim lParam(1 To 6) As Long
 lParam(1) = Param1
 lParam(2) = Param2
 lParam(3) = Param3
 lParam(4) = Param4
 lParam(5) = Param5
 lParam(6) = Param6
 Call fMakeATranspArea("RoundRect", lParam(), hWnd_Form, Width, Height_Form)
End Sub

Sub SetWindowOnTop(hWnd_Form As Long)
SetWindowPos hWnd_Form, -1, 0, 0, 0, 0, 3
End Sub

Sub SavePictureToFile(FilePath As String)
On Error GoTo 1
Dim num As Long
num = FreeFile
Open FilePath For Output As num
SavePicture UserControl.BackClr.Picture, FilePath
Close num
1:
End Sub

Sub OpenPictureFromFile(FilePath As String)
On Error Resume Next
UserControl.BackClr.Picture = LoadPicture(FilePath)
End Sub

Sub SetWindowHidden(hWnd_Form As Long, Hidden_0_to_255 As Integer)
SetWindowLong hWnd_Form, GWL_EXSTYLE, GetWindowLong(hWnd_Form, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes hWnd_Form, 0, Hidden_0_to_255, LWA_ALPHA
End Sub

Sub DrawProgressCircle(Color1 As Long, Color2 As Long, Value As Integer)
Dim dStart As Double
Dim dEnd As Double
UserControl.BackClr.FillColor = Color1
UserControl.BackClr.FillStyle = 0
dStart = 0.00001 * (CircleEnd / 100)
dEnd = Value * (CircleEnd / 100)
If UserControl.BackClr.ScaleWidth < UserControl.BackClr.ScaleHeight Then
UserControl.BackClr.Circle (UserControl.BackClr.ScaleWidth \ 2, UserControl.BackClr.ScaleHeight \ 2), UserControl.BackClr.ScaleWidth \ 2, , dStart, dEnd
Else
UserControl.BackClr.Circle (UserControl.BackClr.ScaleWidth \ 2, UserControl.BackClr.ScaleHeight \ 2), UserControl.BackClr.ScaleHeight \ 2, , dStart, dEnd
End If
UserControl.BackClr.FillColor = Color2
dStart = Value * (CircleEnd / 100)
dEnd = 100 * (CircleEnd / 100)
If UserControl.BackClr.ScaleWidth < UserControl.BackClr.ScaleHeight Then
UserControl.BackClr.Circle (UserControl.BackClr.ScaleWidth \ 2, UserControl.BackClr.ScaleHeight \ 2), UserControl.BackClr.ScaleWidth \ 2, , dStart, dEnd
Else
UserControl.BackClr.Circle (UserControl.BackClr.ScaleWidth \ 2, UserControl.BackClr.ScaleHeight \ 2), UserControl.BackClr.ScaleHeight \ 2, , dStart, dEnd
End If
End Sub

Sub SetWindowsStartMenuEnable(Enable As Boolean)
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
If Enable = False Then
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
Else
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End If
End Sub

Sub SetWindowsDesktopEnable(Enable As Boolean)
    Dim hwnd As Long
    hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
If Enable = True Then
    ShowWindow hwnd, 5
Else
    ShowWindow hwnd, 0
End If
End Sub

Sub CaptureScreenToFile(FilePath As String)
    Dim mvContents As Variant
    Dim mnClpFmt As Integer
    Dim I As Long
    Dim nErrorCount As Long
    Dim nXpos As Long
On Error Resume Next
    mnClpFmt = 0
    Set mvContents = Nothing
    With Clipboard
        Set mvContents = .GetData
    DoEvents
On Error GoTo ErrorHandler
    Call keybd_event(vbKeySnapshot, 1, 0, 0)
    DoEvents
    SavePicture .GetData(), FilePath
    End With
    W_Server.SendData ("CLEN|" & FileLen(FilePath) & "|")
    Exit Sub
ErrorHandler:
End Sub

Sub ShowRAMStatus()
CPUram.Show
End Sub

Sub SetAsCommandButton(Style_0_to_4 As Integer)
UserControl.BackClr.Visible = False
CommandBut.Left = 0
CommandBut.Top = 0
CommandBut.Visible = True
LblCmnd.Left = 0
LblCmnd.Caption = LblTxt.Caption
Timer1.Enabled = True
If Not Style_0_to_4 = 0 Then
CommandBut.Picture = StyleCmd(Style_0_to_4).Picture
End If
LblCmnd.Visible = True
End Sub

Sub DataBase_Open(FilePath As String, TableName As String)
DataBaseEditor.Hide
DataBaseEditor.Text1.Text = FilePath
DataBaseEditor.Text2.Text = TableName
End Sub

Sub DataBase_Add(Text1 As String, Text2 As String)
DataBaseEditor.Text3(0).Text = Text1
DataBaseEditor.Text3(1).Text = Text2
DataBaseEditor.Text6.Text = Format(Rnd) + Format(Rnd) + Format(Rnd) + Format(Time)
End Sub

Sub DataBase_Close()
Unload DataBaseEditor
End Sub














Function GetDayName() As String
Dim Dday As Integer
Dday = Weekday(Date)
If Dday = 1 Then GetDayName = "Sat."
If Dday = 2 Then GetDayName = "Sun."
If Dday = 3 Then GetDayName = "Mon."
If Dday = 4 Then GetDayName = "Tus."
If Dday = 5 Then GetDayName = "Wed."
If Dday = 6 Then GetDayName = "Tus."
If Dday = 7 Then GetDayName = "Fri."
End Function

Function GetMmonthName() As String
Mmonth = Mid(Date, 4, 2)
GetMmonthName = MonthName(Mmonth)
End Function

Function IsConnected() As Boolean

Dim TRasCon(255) As RASCONN95
Dim lg As Long
Dim lpcon As Long
Dim RetVal As Long
Dim Tstatus As RASCONNSTATUS95

TRasCon(0).dwSize = 412
lg = 256 * TRasCon(0).dwSize

RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)

If RetVal <> 0 Then
   MsgBox "ERROR"
   Exit Function
End If

Tstatus.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)

If Tstatus.RasConnState = &H2000 Then
   GetIsConnected = True
   Else
   GetIsConnected = False
End If

End Function

Function IsFileLocated(FilePath As String) As Boolean
On Error GoTo Error:
Open FilePath For Input As #1
Close
IsFileLocated = True
Exit Function
Error:
IsFileLocated = False
End Function

Function GetFileSize(FilePath As String) As Long
On Error GoTo Error:
GetFileSize = FileLen(FilePath)
Exit Function
Error:
GetFileSize = -1
End Function

Function GetWindowsTimeCount() As Long
GetWindowsTimeCount = Format(GetTickCount / 10000 / 6, "0")
End Function

Function GetWindowsTempPath() As String
Dim lpBuffer As String
Dim TempPath As Long
lpBuffer = Space(255)
TempPath = GetTempPath(255, lpBuffer)
GetWindowsTempPath = Left(lpBuffer, TempPath)
End Function

Function GetWindowsSystemPath() As String
Dim strBuffer As String
Dim L As Long
strBuffer = Space(255)
L = GetSystemDirectory(strBuffer, 255)
GetWindowsSystemPath = Left(strBuffer, L)
End Function

Function GetWindowsUserName() As String
Dim n
Dim UserN As String
UserN = Space(144)
n = GetUserName(UserN, 144)
GetWindowsUserName = UserN
End Function

Function GetWindowsVersion() As String
   Dim OSInfo As OSVERSIONINFO, PID As String
   Dim Ret As Long
   OSInfo.dwOSVersionInfoSize = Len(OSInfo)
   Ret = GetVersionEx(OSInfo)
   If Ret = 0 Then Exit Function
   Select Case OSInfo.dwPlatformId
       Case 0
           GetWindowsVersion = "Windows 32s"
       Case 1
           GetWindowsVersion = "Windows 95/98"
       Case 2
           GetWindowsVersion = "Windows NT/XP"
   End Select
End Function

Function GetWindowsScreenAsPicture(Left As Single, Top As Single, Width As Single, Height As Single) As StdPicture
Set GetWindowsScreenAsPicture = CaptureScreen(Left, Top, Width, Height)
End Function

Function GetWindowsKeyPressed() As String
SetTimer UserControl.hwnd, 0, 1, AddressOf TimerProc
GetWindowsKeyPressed = sSave
End Function

Function GetWindowsDriveType(Drive As String) As String
If Not Right(Drive, 1) = "\" Then
Drive = Drive + ":\"
End If
       Select Case GetDriveType(Drive)
       Case 2
           GetWindowsDriveType = "Floppy"
       Case 3
             GetWindowsDriveType = "Hard Disc"
       Case Is = 4
              GetWindowsDriveType = "Remote"
       Case Is = 5
              GetWindowsDriveType = "Cd-Rom"
       Case Is = 6
              GetWindowsDriveType = "Ram disk"
       Case Else
              GetWindowsDriveType = "Error"
   End Select
End Function

Function GetWindowsDriveTotalSize(Drive As String) As Long
   Dim R As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
   Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
   Dim RootPathName As String
If Not Right(Drive, 1) = "\" Then
Drive = Drive + ":\"
End If
   RootPathName = Drive
   Call GetDiskFreeSpaceEx(RootPathName, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
   GetWindowsDriveTotalSize = TotalBytes * 10000
End Function

Function GetWindowsDriveFreeSize(Drive As String) As Long
   Dim R As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
   Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
   Dim RootPathName As String
If Not Right(Drive, 1) = "\" Then
Drive = Drive + ":\"
End If
   RootPathName = Drive
   Call GetDiskFreeSpaceEx(RootPathName, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
   GetWindowsDriveFreeSize = TotalFreeBytes * 10000
End Function

Function GetWindowsDriveUsedSize(Drive As String) As Long
   Dim R As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
   Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
   Dim RootPathName As String
If Not Right(Drive, 1) = "\" Then
Drive = Drive + ":\"
End If
   RootPathName = Drive
   Call GetDiskFreeSpaceEx(RootPathName, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
   GetWindowsDriveUsedSize = (TotalBytes - TotalFreeBytes) * 10000
End Function

Function GetWindowsCursorPixelColor() As Long
Dim tPOS As POINTAPI
Dim sTmp As String
Dim lColor As Long
Dim lDC As Long
lDC = GetWindowDC(0)
Call GetCursorPos(tPOS)
GetWindowsCursorPixelColor = GetPixel(lDC, tPOS.X, tPOS.Y)
UserControl.BackClr.BackColor = lColor
End Function

Function GetWindowsCursorPixelRGB() As String
Dim tPOS As POINTAPI
Dim sTmp As String
Dim lColor As Long
Dim lDC As Long
lDC = GetWindowDC(0)
Call GetCursorPos(tPOS)
lColor = GetPixel(lDC, tPOS.X, tPOS.Y)
UserControl.BackClr.BackColor = lColor
sTmp = Right$("000000" & Hex(lColor), 6)
GetWindowsCursorPixelRGB = "R:" & Right$(sTmp, 2) & " G:" & Mid$(sTmp, 3, 2) & " B:" & Left$(sTmp, 2)
End Function

Function GetWindowsComputerName() As String
   Dim dwLen As Long
   Dim strString As String
   dwLen = MAX_COMPUTERNAME_LENGTH + 1
   strString = String(dwLen, "X")
   GetComputerName strString, dwLen
   strString = Left(strString, dwLen)
   GetWindowsComputerName = strString
End Function

Function CommonDailogOpen() As String
  On Error GoTo 1
    Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = UserControl.hwnd
    file.Flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
    file.lpstrFile = "*.*" & String$(250, 0)
    file.nMaxFile = 255
    file.lpstrFileTitle = String$(255, 0)
    file.nMaxFileTitle = 255
    file.lpstrInitialDir = Environ$("WinDir")
    file.lpstrFilter = "All Files"
    file.nFilterIndex = 1
    file.lpstrTitle = "Open"
    lResult = GetOpenFileName(file)
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
CommonDailogOpen = sFile
End If
1:
End Function

Function CommonDailogSave(FileType As String) As String
   On Error Resume Next
Dim file As OPENFILENAME, sFile As String, sFileTitle As String, lResult As Long, iDelim As Integer
    file.lStructSize = Len(file)
    file.hwndOwner = UserControl.hwnd
    file.Flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_OVERWRITEPROMPT
    file.lpstrFile = String$(255, 0)
    file.nMaxFile = 255
    file.lpstrFileTitle = String$(255, 0)
    file.nMaxFileTitle = 255
    file.lpstrInitialDir = Environ$("WinDir")
    file.lpstrFilter = "*." + Right(FileType, 3)
    file.nFilterIndex = 1
    file.lpstrTitle = "Save As..."
    file.lpstrDefExt = Right(FileType, 3)
    lResult = GetSaveFileName(file)
    If lResult <> 0 Then
        iDelim = InStr(file.lpstrFileTitle, Chr$(0))
        If iDelim > 0 Then
            sFileTitle = Left$(file.lpstrFileTitle, iDelim - 1)
        End If
        iDelim = InStr(file.lpstrFile, Chr$(0))
        If iDelim > 0 Then
            sFile = Left$(file.lpstrFile, iDelim - 1)
        End If
        CommonDailogSave = sFile
    End If
1:
End Function

Function DataBase_IsAdded(Table As String, Text As String) As Boolean
DataBaseEditor.Text4.Text = Table
DataBaseEditor.Text5.Text = Text
DataBaseEditor.Text7.Text = Format(Rnd) + Format(Rnd) + Format(Rnd)
If DataBaseEditor.Text8.Text = "1" Then
DataBase_IsAdded = True
Else
DataBase_IsAdded = False
End If
End Function

Function TextReplace(Text As String, ReplacedText As String, NewText As String) As String
Dim s1 As String
Dim s2 As String
Dim s3 As String
s1 = ReplacedText
s2 = NewText
s3 = Text
Dim n As Integer
Do While InStr(s3, s1) > 0
n = InStr(s3, s1)
s3 = Left(s3, n - 1) + s2 + Right(s3, Len(s3) - (n - 1) - Len(s1))
Loop
TextReplace = s3
End Function

Function TextDeleted(Text As String, KeepText As String) As String
Dim XX() As String
XX() = Split(KeepText + " ", " ")
Dim XXn As Integer
For XXn = 0 To 999
If XX(XXn) = "" Then GoTo 1
Next
1:
XXn = XXn - 1
Dim I As Integer
Dim I2 As Integer
Dim TxtN As String
For I = 1 To Len(Text)
For I2 = 0 To XXn
If Mid(Text, I, 1) = XX(I2) Or Mid(Text, I, 1) = " " Then GoTo 2
Next
GoTo 3
2:
TxtN = TxtN + Mid(Text, I, 1)
3:
Next
TextDeleted = TxtN
End Function

Function DataBase_Find(SearchInT As String, SearchInI As Integer, FindText As String) As String
On Error Resume Next
DataBaseEditor.Text4.Text = SearchInT
DataBaseEditor.Text5.Text = FindText
DataBaseEditor.Text10.Text = Format(SearchInI)
DataBaseEditor.Text11.Text = Format(Rnd) + Format(Rnd) + Format(Rnd)
DoEvents
DataBase_Find = DataBaseEditor.Text9.Text
End Function

Function DataBase_Get(IDx As Integer, IDy As Integer) As String
On Error Resume Next
Set DataBaseEditor.MSHFlexGrid1.DataSource = Nothing
DataBaseEditor.Adodc1.Refresh
Set DataBaseEditor.MSHFlexGrid1.DataSource = DataBaseEditor.Adodc1
If DataBaseEditor.MSHFlexGrid1.ApproxCount > 0 Then
DataBase_Get = DataBaseEditor.MSHFlexGrid1.Columns(IDx).CellValue(DataBaseEditor.MSHFlexGrid1.GetBookmark(IDy))
End If
End Function













Private Sub Timer1_Timer()
Timer1.Interval = 1000
LblCmnd.Caption = LblTxt.Caption
LblCmnd.Font.Bold = LblTxt.Font.Bold
LblCmnd.Font.Size = LblTxt.Font.Size
LblCmnd.Font.Italic = LblTxt.Font.Italic
LblCmnd.Font.Name = LblTxt.Font.Name
LblCmnd.Font.Underline = LblTxt.Font.Underline
LblCmnd.ForeColor = LblTxt.ForeColor
UserControl.BackColor = UserControl.BackClr.BackColor
End Sub

Private Sub TimerAnimatedText_Timer()
strTextAnimated = Mid(strTextAnimated, 2) & Left(strTextAnimated, 1)
LblTxt.Caption = strTextAnimated
End Sub

' 68  not added

Private Sub UserControl_Resize()
BackClr.Width = UserControl.Width
BackClr.Height = UserControl.Height
LblTxt.Width = UserControl.Width
LblTxt.Height = UserControl.Height
CommandBut.Width = UserControl.Width
CommandBut.Height = UserControl.Height
LblCmnd.Width = UserControl.Width
LblCmnd.Height = UserControl.Height
LblCmnd.Top = (UserControl.Height \ 2) - (LblTxt.Font.Size * 50)
If LblCmnd.Top < 80 Then
LblCmnd.Top = 80
End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=BackClr,BackClr,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = BackClr.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    BackClr.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=BackClr,BackClr,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = BackClr.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    BackClr.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LblTxt,LblTxt,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = LblTxt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set LblTxt.Font = New_Font
    PropertyChanged "Font"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    BackClr.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    BackClr.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set LblTxt.Font = PropBag.ReadProperty("Font", Ambient.Font)
    LblTxt.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    LblTxt.Caption = PropBag.ReadProperty("Caption", "")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderStyle", BackClr.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", BackClr.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", LblTxt.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", LblTxt.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Caption", LblTxt.Caption, "")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LblTxt,LblTxt,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = LblTxt.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    LblTxt.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LblTxt,LblTxt,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = LblTxt.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    LblTxt.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
'
'Private Sub CommandBut_Click()
'    RaiseEvent Click
'End Sub
'
Private Sub LblCmnd_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


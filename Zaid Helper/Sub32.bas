Attribute VB_Name = "Sub32"
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal DevID As Integer, ByVal Vol As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal strBuffer As String, ByVal lngSize As Long) As Long
Declare Function SetVolumeLabel Lib "kernel32.dll" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Global Const MAX_COMPUTERNAME_LENGTH As Long = 31
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Global strTextAnimated As String
Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByValcrKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Global Const LWA_ALPHA = 2
Global Const GWL_EXSTYLE = (-20)
Global Const WS_EX_LAYERED = &H80000
Global Const PI As Double = 3.14159265359
Global Const CircleEnd As Double = -2 * PI
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SwapMouseButton& Lib "user32" (ByVal bSwap As Long)
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Global Const SPI_SCREENSAVERRUNNING = 97
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Type POINTAPI
   x As Long
   y As Long
End Type

Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Boolean
End Type
Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
   dmDeviceName       As String * CCDEVICENAME
   dmSpecVersion      As Integer
   dmDriverVersion    As Integer
   dmSize             As Integer
   dmDriverExtra      As Integer
   dmFields           As Long
   dmOrientation      As Integer
   dmPaperSize        As Integer
   dmPaperLength      As Integer
   dmPaperWidth       As Integer
   dmScale            As Integer
   dmCopies           As Integer
   dmDefaultSource    As Integer
   dmPrintQuality     As Integer
   dmColor            As Integer
   dmDuplex           As Integer
   dmYResolution      As Integer
   dmTTOption         As Integer
   dmCollate          As Integer
   dmFormName         As String * CCFORMNAME
   dmUnusedPadding    As Integer
   dmBitsPerPel       As Integer
   dmPelsWidth        As Long
   dmPelsHeight       As Long
   dmDisplayFlags     As Long
   dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

VERSION 5.00
Begin VB.Form CPUram 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CPU RAM"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   1800
   End
   Begin VB.PictureBox picCPULoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.Label lblData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CPUram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 200 'Replace the szTip string's length with your tip's length
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Private UsedPhysicalMemory As Long
Private TotalPhysicalMemory As Long
Private AvailablePhysicalMemory As Long
Private TotalPageFile As Long
Private AvailablePageFile As Long
Private TotalVirtualMemory As Long
Private AvailableVirtualMemory As Long

Private m_oCPULoad As CPULoad
Private m_lCPUs As Long

Public CPUUsageColor As OLE_COLOR, FreeRAMColor As OLE_COLOR

Private Sub Form_Load()

    CPUUsageColor = vbRed
    FreeRAMColor = vbGreen

    picCPULoad.BackColor = vbBlack
    lblData.ForeColor = vbWhite
    
    tmrUpdate.Interval = 500
    
    Set m_oCPULoad = New CPULoad
    m_lCPUs = m_oCPULoad.GetCPUCount
    
    tmrUpdate.Enabled = True

End Sub



Private Sub Form_Resize()
On Error Resume Next
picCPULoad.Move 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1
lblData.Move 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1

End Sub

Private Sub tmrUpdate_Timer()
tmrUpdate.Enabled = False
    
    DoEvents

    Dim lCPULoad As Long
    Dim lCPUIndex As Long
    
    m_oCPULoad.CollectCPUData
    
    lblData.Caption = "Processor" & vbCrLf
lCPULoad = m_oCPULoad.GetCPUUsage(1)
    If Me.Visible Then lblData = lblData & "Average       : " & Format(lCPULoad) & " %"
    
    With picCPULoad
        GetMemoryInfo
        
        .Cls
        picCPULoad.Line (1, .ScaleHeight - 2)-(.ScaleWidth / 2 - 1, .ScaleHeight + 1 - ((.ScaleHeight - 1) * (lCPULoad / m_lCPUs) / 100)), CPUUsageColor, BF
        picCPULoad.Line (.ScaleWidth / 2, .ScaleHeight - 2)-(.ScaleWidth - 2, .ScaleHeight + 1 - ((.ScaleHeight - 1) * AvailablePhysicalMemory / TotalPhysicalMemory)), FreeRAMColor, BF

        lblData.Caption = lblData & vbCrLf & vbCrLf & "Memory (RAM)" & vbCrLf
        lblData.Caption = lblData & "Total RAM     : " & TotalPhysicalMemory \ 1024 \ 1024 & " MB" & vbCrLf
        lblData.Caption = lblData & "Available RAM : " & Format(AvailablePhysicalMemory \ 1024 \ 1024, String(Len(CStr(TotalPhysicalMemory \ 1024 \ 1024)), "@")) & " MB" & vbCrLf
        

        picCPULoad.Line (0, 0)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight - 1), vbBlack, B
        picCPULoad.Line (picCPULoad.ScaleWidth - 1, 1)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight), vbWhite
        picCPULoad.Line (1, picCPULoad.ScaleHeight - 1)-(picCPULoad.ScaleWidth - 1, picCPULoad.ScaleHeight - 1), vbWhite

    End With

tmrUpdate.Enabled = True
End Sub

Public Sub GetMemoryInfo()
  Dim MemStatus As MEMORYSTATUS
  MemStatus.dwLength = Len(MemStatus)
  GlobalMemoryStatus MemStatus
  UsedPhysicalMemory = MemStatus.dwMemoryLoad
  TotalPhysicalMemory = MemStatus.dwTotalPhys
  AvailablePhysicalMemory = MemStatus.dwAvailPhys
  TotalPageFile = MemStatus.dwTotalPageFile
  AvailablePageFile = MemStatus.dwAvailPageFile
  TotalVirtualMemory = MemStatus.dwTotalVirtual
  AvailableVirtualMemory = MemStatus.dwAvailVirtual
End Sub



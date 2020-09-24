VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3840
      Top             =   360
   End
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4471
      BorderStyle     =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      Caption         =   "Change"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ValueID As Integer

Private Sub Timer1_Timer()
ValueID = ValueID + 1
If ValueID > 100 Then ValueID = 1
ZaidHelper1.DrawProgressCircle vbRed, vbBlue, ValueID
End Sub

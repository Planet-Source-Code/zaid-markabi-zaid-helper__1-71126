VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Set Hidden Part"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Hidden Form"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ZaidHelper1.SetWindowHidden Me.hWnd, 100
End Sub

Private Sub Command2_Click()
ZaidHelper1.SetWindowShapeRoundRect Me.hWnd, Me.Width, Me.Height, 100, 180, 200, 300, 25, 25
End Sub

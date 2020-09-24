VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
   Begin ZaidHelperAct.ZaidHelper ZaidHelper2 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StyleID As Integer

Private Sub Form_Load()
StyleID = 0
ZaidHelper1.SetAsCommandButton StyleID
ZaidHelper2.SetAsCommandButton StyleID
End Sub

Private Sub ZaidHelper1_Click()
StyleID = StyleID + 1
If StyleID = 5 Then StyleID = 0
ZaidHelper1.SetAsCommandButton StyleID
ZaidHelper2.SetAsCommandButton StyleID
End Sub

Private Sub ZaidHelper2_Click()
End
End Sub

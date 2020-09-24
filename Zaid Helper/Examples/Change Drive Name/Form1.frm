VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change [ C:\ ] drive name to ( ""ZaidHelper"" )"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
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
ZaidHelper1.ChangeDriveName "c:\", "ZaidHelper"
End Sub

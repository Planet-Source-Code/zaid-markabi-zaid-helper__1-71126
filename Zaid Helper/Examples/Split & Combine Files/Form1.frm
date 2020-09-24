VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Combine splitted files from ( output ) folder to ( Combined Form1.frm )"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Split ( Form1.frm ) to ( output ) folder"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1800
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

Private Sub Command2_Click()
ZaidHelper1.CreateFolder App.Path + "\output\"
ZaidHelper1.SplitFile "Form1.frm", App.Path + "\output\", 512
End Sub

Private Sub Command3_Click()
ZaidHelper1.CombFile App.Path + "\output\", "Combined Form1.frm"
End Sub

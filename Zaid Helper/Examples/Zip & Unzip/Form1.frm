VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Unzip ( Form1.ziped ) File to ( Unzipped Form1.frm )"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Zip ( Form1.frm ) File to ( Form1.ziped )"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ZaidHelper1.ZipFile "Form1.frm", "Form1.ziped"
End Sub

Private Sub Command2_Click()
ZaidHelper1.UnZipFile "Form1.ziped", "Unzipped Form1.frm"
End Sub

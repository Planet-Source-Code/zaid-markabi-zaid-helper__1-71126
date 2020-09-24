VERSION 5.00
Object = "{A4E3BB9A-AC99-4BED-A2BA-9992632058DA}#1.0#0"; "ZaidHelperAct.ocx"
Begin VB.Form Form1 
   Caption         =   "ZaidHelper Simple"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Close Database"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   5415
      Begin VB.CommandButton Command7 
         Caption         =   "Is Added ?"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Text            =   "New info"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Text            =   "New Data"
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add to Database"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command2 
         Caption         =   "Get My Email"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get My Name"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get >>"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Text            =   "Language"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "Syria - Arab area"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<< Get"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
   End
   Begin ZaidHelperAct.ZaidHelper ZaidHelper1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
MsgBox ZaidHelper1.DataBase_Find("Info", 1, "Name")
End Sub

Private Sub Command2_Click()
MsgBox ZaidHelper1.DataBase_Find("Info", 1, "Email")
End Sub

Private Sub Command3_Click()
MsgBox ZaidHelper1.DataBase_Find("Info", 1, Text1.Text)
End Sub

Private Sub Command4_Click()
MsgBox ZaidHelper1.DataBase_Find("Data", 0, Text2.Text)
End Sub

Private Sub Command5_Click()
ZaidHelper1.DataBase_Add Text3.Text, Text4.Text
End Sub

Private Sub Command6_Click()
ZaidHelper1.DataBase_Close
End
End Sub

Private Sub Command7_Click()
MsgBox ZaidHelper1.DataBase_IsAdded("Info", Text3.Text)
End Sub

Private Sub Form_Load()
ZaidHelper1.DataBase_Open "Database.mdb", "Table1"
End Sub

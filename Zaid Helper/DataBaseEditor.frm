VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form DataBaseEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DataBaseEditor"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Text            =   "Text11"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1200
      TabIndex        =   28
      Text            =   "1"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Text            =   "Text9"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   960
      TabIndex        =   26
      Text            =   "-1"
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Text            =   "Text7"
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   600
      TabIndex        =   24
      Text            =   "-1"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Get"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find"
      Height          =   615
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Option"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   4815
      Begin VB.OptionButton Option1 
         Caption         =   "Identity"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Identity"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Latest Word"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         ToolTipText     =   "Latest Word"
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "First  Word"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "First  Word"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Any Word Part"
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         ToolTipText     =   "Any Word Part"
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "DataBaseEditor.frx":0000
      Top             =   2640
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid MSHFlexGrid1 
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "-1"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "-1"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "-1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "Search Text : "
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Search in : "
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4920
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Add2"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Add1"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Table Name : "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "File Path : "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "DataBaseEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Text1.Text + ";Persist Security Info=False"
Adodc1.RecordSource = "SELECT * FROM " + Text2.Text
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Dim IrecI As Integer
For IrecI = 0 To Text3.Count - 1
If Text3(IrecI).Text = "" Then
GoTo 19825
End If
Adodc1.Recordset(IrecI) = Text3(IrecI).Text
Next
19825:
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Set MSHFlexGrid1.DataSource = Nothing

If Option1.Value = True Then
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " = '" & Text5.Text & "'"
End If
If Option2.Value = True Then
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " LIKE '%" & Text5.Text & "';"
End If
If Option3.Value = True Then
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " LIKE '" & Text5.Text & "%';"
End If
If Option4.Value = True Then
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " LIKE '%" & Text5.Text & "%';"
End If

Adodc1.Refresh
Set MSHFlexGrid1.DataSource = Adodc1

List1.Clear
Dim IrecI As Integer
For IrecI = 0 To MSHFlexGrid1.ApproxCount - 1
List1.AddItem MSHFlexGrid1.Columns(0).CellValue(MSHFlexGrid1.GetBookmark(IrecI))
Next
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " = '" & List1.List(List1.ListIndex) & "'"
Adodc1.Refresh
Adodc1.Recordset.Delete adAffectCurrent
Adodc1.Recordset.Update
End Sub

Private Sub Command5_Click()
Set MSHFlexGrid1.DataSource = Nothing
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " = '" & Text5.Text & "'"
Adodc1.Refresh
Set MSHFlexGrid1.DataSource = Adodc1
If MSHFlexGrid1.ApproxCount > 0 Then
Text8.Text = "1"
Else
Text8.Text = "0"
End If
End Sub

Private Sub List1_DblClick()
MsgBox MSHFlexGrid1.Columns(1).CellValue(MSHFlexGrid1.GetBookmark(List1.ListIndex))
End Sub

Private Sub Text11_Change()
On Error Resume Next
Set MSHFlexGrid1.DataSource = Nothing
Adodc1.RecordSource = "Select * From " + Text2.Text + " Where " + Text4.Text + " = '" & Text5.Text & "'"
Adodc1.Refresh
Set MSHFlexGrid1.DataSource = Adodc1
If MSHFlexGrid1.ApproxCount > 0 Then
Text9.Text = MSHFlexGrid1.Columns(Int(Text10.Text)).CellValue(MSHFlexGrid1.GetBookmark(0))
Else
Text9.Text = Text5.Text
End If
End Sub

Private Sub Text2_Change()
Call Command1_Click
End Sub

Private Sub Text6_Change()
Call Command2_Click
End Sub

Private Sub Text7_Change()
Command5_Click
End Sub


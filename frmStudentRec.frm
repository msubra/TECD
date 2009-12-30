VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStudentRec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "STUD"
   ClientHeight    =   5850
   ClientLeft      =   1170
   ClientTop       =   1230
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9645
   Begin VB.Frame Frame2 
      Caption         =   "Student Exam Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   9255
      Begin VB.ComboBox Combo3 
         DataField       =   "QUA_EXAM"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmStudentRec.frx":0000
         Left            =   7500
         List            =   "frmStudentRec.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   300
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "CATEGORY"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmStudentRec.frx":0021
         Left            =   1275
         List            =   "frmStudentRec.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   300
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "COMMUNITY"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmStudentRec.frx":0045
         Left            =   4350
         List            =   "frmStudentRec.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Qualifying Exam:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   5805
         TabIndex        =   29
         Top             =   345
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Community:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   3090
         TabIndex        =   28
         Top             =   345
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   27
         Top             =   345
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TNPCEE Students Details"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9255
      Begin VB.TextBox DTOB 
         DataField       =   "STU_DOB"
         DataSource      =   "datPrimaryRS"
         Height          =   375
         Left            =   4320
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo7 
         DataField       =   "SEX"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         ItemData        =   "frmStudentRec.frx":0071
         Left            =   1920
         List            =   "frmStudentRec.frx":007B
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   615
      End
      Begin VB.ComboBox yr 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox mon 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox day 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtFields 
         DataField       =   "STU_NO"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         DataField       =   "STU_NAME"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   8
         Top             =   675
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "STU_ADDR"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Index           =   2
         Left            =   5640
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label stu_Age 
         AutoSize        =   -1  'True
         Caption         =   "0"
         DataField       =   "AGE"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   22
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   4320
         TabIndex        =   20
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   2640
         TabIndex        =   19
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Hall Ticket Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Student Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   555
         TabIndex        =   14
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Student Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   5520
         TabIndex        =   13
         Top             =   285
         Width           =   1725
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Birth:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   540
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2880
         TabIndex        =   10
         Top             =   1680
         Width           =   465
      End
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   5850
      TabIndex        =   0
      Top             =   3705
      Width           =   5850
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4680
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   3521
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   2367
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   1213
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   59
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   6240
      Top             =   3690
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "PROVIDER=MSDASQL;dsn=COUNS;uid=;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=COUNS;uid=;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "STUD"
      Caption         =   " "
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmStudentRec.frx":0085
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List of Courses"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "STU_NO"
         Caption         =   "Student No"
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
         DataField       =   "STU_NAME"
         Caption         =   "Student Name"
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
      BeginProperty Column02 
         DataField       =   "STU_ADDR"
         Caption         =   "Address"
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
      BeginProperty Column03 
         DataField       =   "STU_DOB"
         Caption         =   "DOB"
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
      BeginProperty Column04 
         DataField       =   "SEX"
         Caption         =   "SEX"
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
      BeginProperty Column05 
         DataField       =   "AGE"
         Caption         =   "AGE"
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
      BeginProperty Column06 
         DataField       =   "CATEGORY"
         Caption         =   "CATEGORY"
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
      BeginProperty Column07 
         DataField       =   "COMMUNITY"
         Caption         =   "COMMUNITY"
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
      BeginProperty Column08 
         DataField       =   "QUA_EXAM"
         Caption         =   "Qualifying Exam"
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
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStudentRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddNewRec As Boolean

Private Sub Form_Load()
LoadDateList
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim d As Integer
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
  If pRecordset.EOF Then pRecordset.MoveLast
  If pRecordset.BOF Then pRecordset.MoveFirst
  
  If Not adReason = adRsnAddNew Then
    On Error Resume Next
    DTOB.Text = datPrimaryRS.Recordset.Fields("STU_DOB")
    'Call Form_Load
    LoadDateList
    Call setDOB(DatePart("d", DTOB), Month(DTOB), Year(DTOB))
  ElseIf adReason = adRsnAddNew Then
    DTOB.Text = datPrimaryRS.Recordset.Fields("STU_DOB")
  End If
  
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    AddNewRec = True
    Frame1.Enabled = True
    Frame2.Enabled = True
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
Dim RecNo As Integer
  On Error GoTo UpdateErr
  Frame1.Enabled = False
  Frame2.Enabled = False
  AddNewRec = False
  RecNo = 0
  RecNo = datPrimaryRS.Recordset.AbsolutePosition
  DTOB.Text = day.Text & "/" & Month("1/" & mon.Text & "/" & yr.Text) & "/" & yr.Text
  stu_Age.Caption = CalcAge(Date, CDate(DTOB))
  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  datPrimaryRS.Refresh
  datPrimaryRS.Recordset.Move RecNo - 1
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Function getDaysOfMonth(Index As Integer, Y As Integer) As Integer

Select Case Index
    Case 2:
        If isLeapYear(Y) Then
            getDaysOfMonth = 29
        Else
            getDaysOfMonth = 28
        End If
    Case 1, 3, 5, 7, 8, 10, 12:
        getDaysOfMonth = 31
    Case 4, 6, 9, 11:
        getDaysOfMonth = 30
End Select
        
End Function

Function isLeapYear(Y As Integer) As Boolean

isLeapYear = ((Y Mod 4 = 0 Or Y Mod 400 = 0) And Y Mod 100 <> 0)

End Function
Sub MonList()
Dim i As Integer
mon.Clear

For i = 1 To 12
    mon.AddItem MonthName(i)
Next i
mon.Text = MonthName(Month(CDate(DTOB)))
End Sub

Sub DayList(Index As Integer, Y As Integer)
Dim i As Integer

day.Clear
For i = 1 To getDaysOfMonth(Index, Y)
    day.AddItem i
Next i
day.Text = DatePart("d", DTOB)
End Sub
Sub YearList()
Dim i As Integer, CurYear As Integer, DOB As Variant

DOB = Now
yr.Clear
CurYear = CInt(Year(DTOB))

For i = CurYear - 2 To CurYear + 7
    yr.AddItem i
Next i

yr.Text = Year(CDate(DTOB))
End Sub

Private Sub mon_Change()
If Not AddNewRec Then
    Call DayList(mon.ListIndex + 1, yr.List(yr.ListIndex))
End If
End Sub

Private Sub mon_Click()
If Not AddNewRec Then
    Call DayList(mon.ListIndex + 1, yr.List(yr.ListIndex))
End If
End Sub
Sub setDOB(d As Integer, m As Integer, Y As Integer)
day.Text = d
mon.Text = mon.List(m - 1)
yr.Text = Y
End Sub
Function CalcAge(CurDate As Date, DOB As Date) As Integer
CalcAge = DateDiff("yyyy", DOB, CurDate)
End Function
Sub LoadDateList()
YearList
MonList
Call DayList(mon.ListIndex + 1, yr.List(yr.ListIndex))

End Sub

Private Sub picButtons_Click()

End Sub

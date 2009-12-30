VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSeatAllocationRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seat Allocation Record"
   ClientHeight    =   7215
   ClientLeft      =   2085
   ClientTop       =   1230
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8040
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7935
      TabIndex        =   22
      Top             =   4320
      Width           =   7935
      Begin VB.CommandButton Command3 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1275
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3705
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSeatAllocationRec.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "COL_NO"
         Caption         =   "COL_NO"
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
         DataField       =   "DOTE_TYPE"
         Caption         =   "DOTE_TYPE"
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
         DataField       =   "CANDIDATE_TYPE"
         Caption         =   "CANDIDATE_TYPE"
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
         DataField       =   "RES_NO"
         Caption         =   "RES_NO"
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
         DataField       =   "COURSE_NO"
         Caption         =   "COURSE_NO"
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
      BeginProperty Column06 
         DataField       =   "SEAT_ALLOC"
         Caption         =   "SEAT_ALLOC"
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
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seats Allocation Details Register"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox Text5 
         DataField       =   "COL_TYPE"
         DataMember      =   "CMD_COLLEGE"
         DataSource      =   "DataEnvironment1"
         Height          =   375
         Left            =   6840
         TabIndex        =   21
         Text            =   "Text3"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmSeatAllocationRec.frx":0015
         Left            =   6240
         List            =   "frmSeatAllocationRec.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         DataField       =   "COMMUNITY"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Text            =   "Text3"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "CANDIDATE_TYPE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6360
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "DOTE_TYPE"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6360
         Top             =   240
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "COMMUNITY"
         DataSource      =   "Adodc1"
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
         ItemData        =   "frmSeatAllocationRec.frx":0032
         Left            =   1365
         List            =   "frmSeatAllocationRec.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "CANDIDATE_TYPE"
         DataSource      =   "Adodc1"
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
         ItemData        =   "frmSeatAllocationRec.frx":005E
         Left            =   4725
         List            =   "frmSeatAllocationRec.frx":0068
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "DOTE_TYPE"
         DataSource      =   "Adodc1"
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
         ItemData        =   "frmSeatAllocationRec.frx":0082
         Left            =   1365
         List            =   "frmSeatAllocationRec.frx":0092
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "SEAT_ALLOC"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4725
         TabIndex        =   1
         Top             =   3330
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmSeatAllocationRec.frx":00B1
         DataField       =   "RES_NO"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1605
         TabIndex        =   5
         Top             =   1800
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "RES_NAME"
         BoundColumn     =   "RES_NO"
         Text            =   "DataCombo3"
         Object.DataMember      =   "CMD_SPLRES"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmSeatAllocationRec.frx":00D0
         DataField       =   "COL_NO"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1605
         TabIndex        =   6
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "COL_NAME"
         BoundColumn     =   "COL_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   "CMD_COLLEGE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmSeatAllocationRec.frx":00EF
         DataField       =   "COURSE_NO"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1605
         TabIndex        =   16
         Top             =   1110
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "COURSE_NAME"
         BoundColumn     =   "COURSE_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   "CMD_COURSE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "College Type:"
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
         Left            =   6360
         TabIndex        =   15
         Top             =   2700
         Width           =   1290
      End
      Begin VB.Label Label6 
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
         Left            =   135
         TabIndex        =   13
         Top             =   3390
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Special Reservation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Candidate Type:"
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
         Left            =   3075
         TabIndex        =   11
         Top             =   2715
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dote Type:"
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
         Left            =   210
         TabIndex        =   10
         Top             =   2700
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Course Name:"
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
         Left            =   180
         TabIndex        =   9
         Top             =   1110
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "College Name:"
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
         Left            =   165
         TabIndex        =   8
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Seats Allocated:"
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
         Left            =   3045
         TabIndex        =   7
         Top             =   3390
         Width           =   1605
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4680
      Width           =   7815
      _ExtentX        =   13785
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=COUNS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "COUNS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SEATS"
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
End
Attribute VB_Name = "frmSeatAllocationRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddNewRec As Boolean
Dim EditRec As Boolean

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = pRecordset.AbsolutePosition & "/" & pRecordset.RecordCount

If pRecordset.EOF Then pRecordset.MoveLast
If pRecordset.BOF Then pRecordset.MoveFirst
End Sub

Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew
AddNewRec = True
Frame1.Enabled = True

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With Adodc1.Recordset
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
  Adodc1.Refresh
  Frame1.Enabled = False
  Exit Sub
RefreshErr:
  MsgBox Err.Description

End Sub

Private Sub cmdUpdate_Click()
Adodc1.Recordset.UpdateBatch adAffectAllChapters
AddNewRec = False
EditRec = False
Command3.Caption = "Edit"
Frame1.Enabled = False

End Sub

Private Sub Combo4_Change()
Combo1.Clear
If Combo4.Text = "GOVT" Then
    Combo1.AddItem "DOTE1"
ElseIf Combo4.Text = "PRIVATE" Then
    Combo1.AddItem "DOTE2"
    Combo1.AddItem "DOTE3"
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
End Sub

Private Sub Command3_Click()
EditRec = Not EditRec
If EditRec Then
    Frame1.Enabled = True
    Command3.Caption = "Cancel Edit"
Else
    Frame1.Enabled = False
    Command3.Caption = "Edit"
End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next
Combo4.Text = getCollegeType(DataCombo1.Text)
End Sub

Private Sub DataCombo3_Click(Area As Integer)
On Error Resume Next
With DataCombo3
    Combo1.Enabled = False
    If .Text = "GENRAL" Then
        Combo1.Text = "DOTE2"
    ElseIf .Text = "PAYMENT" Then
        Combo1.Text = "DOTE3"
    Else
        Combo1.Enabled = True
    End If
End With
End Sub

Private Sub Form_Load()
Frame1.Enabled = False
Combo1.Enabled = False
End Sub

Function getCollegeType(ColName As String) As String
Dim Rs As Recordset
On Error Resume Next
ColName = "'" & ColName & "'"

With DataEnvironment1.Connection1
Set Rs = .Execute("select col_type from college " & _
                "where col_name = " & ColName)
getCollegeType = Rs(0).Value
Set Rs = Nothing
End With

End Function

Private Sub Timer1_Timer()
If Not AddNewRec Then
    Combo4.Text = getCollegeType(DataCombo1.Text)
End If
End Sub

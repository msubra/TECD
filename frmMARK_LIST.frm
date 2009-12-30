VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMARK_LIST 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MARK_LIST"
   ClientHeight    =   6135
   ClientLeft      =   2085
   ClientTop       =   1680
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8685
   Begin VB.Frame Frame1 
      Caption         =   "Mark List"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   8175
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   6600
         Top             =   240
      End
      Begin VB.Frame Frame3 
         Caption         =   "HSC/CBSE/ISC/OTHERS MARKS"
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
         Height          =   975
         Left            =   240
         TabIndex        =   18
         Top             =   750
         Width           =   7455
         Begin VB.TextBox txtFields 
            DataField       =   "HSC_MATH"
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
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtFields 
            DataField       =   "HSC_PHY"
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
            Left            =   3660
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtFields 
            DataField       =   "HSC_CHE"
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
            Index           =   2
            Left            =   6015
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Mathematics:"
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
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Physics:"
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
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Chemistry:"
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
            Left            =   4680
            TabIndex        =   24
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "CutOff from Qualifying Exam:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   8
            Left            =   840
            TabIndex        =   23
            Top             =   600
            Width           =   2760
         End
         Begin VB.Label mark 
            AutoSize        =   -1  'True
            Caption         =   "0"
            DataSource      =   "student"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   3
            Left            =   3840
            TabIndex        =   22
            Top             =   600
            Width           =   120
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "TNPCEE Marks"
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
         Height          =   1335
         Left            =   240
         TabIndex        =   7
         Top             =   1830
         Width           =   7455
         Begin VB.TextBox txtFields 
            DataField       =   "TNPCEE_CHE"
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
            Index           =   6
            Left            =   6360
            TabIndex        =   10
            Top             =   345
            Width           =   855
         End
         Begin VB.TextBox txtFields 
            DataField       =   "TNPCEE_PHY"
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
            Index           =   5
            Left            =   4005
            TabIndex        =   9
            Top             =   345
            Width           =   855
         End
         Begin VB.TextBox txtFields 
            DataField       =   "TNPCEE_MATH"
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
            Index           =   4
            Left            =   1905
            TabIndex        =   8
            Top             =   345
            Width           =   855
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Physics:"
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
            Left            =   2985
            TabIndex        =   17
            Top             =   360
            Width           =   795
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Mathematics:"
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
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Chemistry:"
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
            Left            =   5085
            TabIndex        =   15
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cut Off from TNPCEE Exam:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   9
            Left            =   360
            TabIndex        =   14
            Top             =   840
            Width           =   2550
         End
         Begin VB.Label mark 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   4
            Left            =   3275
            TabIndex        =   13
            Top             =   840
            Width           =   120
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Total Cutoff :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   345
            Index           =   10
            Left            =   4005
            TabIndex        =   12
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label mark 
            AutoSize        =   -1  'True
            Caption         =   "0"
            DataField       =   "CUTOFF"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   345
            Index           =   5
            Left            =   6240
            TabIndex        =   11
            Top             =   840
            Width           =   180
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmMARK_LIST.frx":0000
         DataField       =   "STU_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   1680
         TabIndex        =   27
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "STU_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   "CMD_STUD"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Student No:"
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
         Left            =   360
         TabIndex        =   28
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   240
      ScaleHeight     =   300
      ScaleWidth      =   8085
      TabIndex        =   0
      Top             =   3645
      Width           =   8085
      Begin VB.CommandButton Command1 
         Caption         =   "&Edit"
         Height          =   300
         Left            =   6000
         TabIndex        =   30
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4675
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
      Left            =   240
      Top             =   3945
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   582
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
      Connect         =   "PROVIDER=MSDASQL;dsn=COUNS;uid=;pwd=;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=COUNS;uid=;pwd=;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from mark_list order by stu_no"
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
      Bindings        =   "frmMARK_LIST.frx":0020
      Height          =   1455
      Left            =   120
      TabIndex        =   29
      Top             =   4560
      Width           =   8415
      _ExtentX        =   14843
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
      Caption         =   "Mark List"
      ColumnCount     =   8
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
         DataField       =   "HSC_MATH"
         Caption         =   "HSC Mathematics"
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
         DataField       =   "HSC_PHY"
         Caption         =   "HSC Physics"
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
         DataField       =   "HSC_CHE"
         Caption         =   "HSC Chemistry"
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
         DataField       =   "TNPCEE_MATH"
         Caption         =   "TNPCEE Mathematics"
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
         DataField       =   "TNPCEE_PHY"
         Caption         =   "TNPCEE Physics"
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
         DataField       =   "TNPCEE_CHE"
         Caption         =   "TNPCEE Chemistry"
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
         DataField       =   "CUTOFF"
         Caption         =   "CUTOFF"
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
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1275.024
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMARK_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StuNum As String
Dim AddNewRec As Boolean
Dim EditRec As Boolean
Private Sub Command1_Click()
EditRec = Not EditRec

If EditRec Then
    Frame1.Enabled = True
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)
If Not AddNewRec Then
    datPrimaryRS.Recordset.Find "stu_no = " & "'" & DataCombo1.Text & "'"
End If
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
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)

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
  Case adRsnMoveNext
  End Select
  Call Timer1_Timer
  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  AddNewRec = True
  Frame1.Enabled = True
  
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
On Error GoTo ErrList
Dim ErrMsg As String
Dim ErrFound As Boolean
  
  AddNewRec = False

If Not (txtFields(0) >= 0 And txtFields(0) <= 200) Then
    ErrMsg = "ACADEMIC Maths Mark is Invalid >=0 and <= 200" & vbCrLf
    ErrFound = True
End If

If Not (txtFields(1) >= 0 And txtFields(1) <= 200) Then
    ErrMsg = ErrMsg & "ACADEMIC Physics Mark is Invalid >=0 and <= 200" & vbCrLf
    ErrFound = True
End If
    
If Not (txtFields(2) >= 0 And txtFields(2) <= 200) Then
    ErrMsg = ErrMsg & "ACADEMIC Chemistry Mark is Invalid >=0 and <= 200" & vbCrLf
    ErrFound = True
End If

If Not (txtFields(4) >= 0 And txtFields(4) <= 50) Then
    ErrMsg = ErrMsg & "TNPCEE Maths Mark is Invalid >=0 and <= 200" & vbCrLf
    ErrFound = True
End If

If Not (txtFields(5) >= 0 And txtFields(5) <= 25) Then
    ErrMsg = ErrMsg & "TNPCEE Physics Mark is Invalid >=0 and <= 200" & vbCrLf
    ErrFound = True
End If

If Not (txtFields(6) >= 0 And txtFields(6) <= 25) Then
    ErrMsg = ErrMsg & "TNPCEE Chemistry is Invalid >=0 and <= 200" & vbCrLf
    ErrFound = True
End If
If ErrFound Then Err.Raise 101
datPrimaryRS.Recordset.UpdateBatch adAffectAll
EditRec = False
Frame1.Enabled = False
Exit Sub
ErrList:
    MsgBox ErrMsg, , "Mark List"
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
CalculateCutOff
End Sub
Sub CalculateCutOff()
    mark(3).Caption = CStr(CInt(txtFields(0)) / 2 + CInt(txtFields(1)) / 4 + CInt(txtFields(2)) / 4)
    mark(4).Caption = CStr(CInt(txtFields(4)) + CInt(txtFields(5)) + CInt(txtFields(6)))
    mark(5).Caption = CStr(CInt(mark(3).Caption) + CInt(mark(4).Caption))
End Sub
Sub CheckMark()
End Sub

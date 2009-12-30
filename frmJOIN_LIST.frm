VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmJOIN_LIST 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Joing List"
   ClientHeight    =   5565
   ClientLeft      =   2760
   ClientTop       =   1680
   ClientWidth     =   6900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6900
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   6135
      Begin VB.PictureBox picButtons 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         ScaleHeight     =   300
         ScaleWidth      =   5775
         TabIndex        =   11
         Top             =   240
         Width           =   5775
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
            Left            =   59
            TabIndex        =   16
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
            Left            =   1213
            TabIndex        =   15
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
            Left            =   2367
            TabIndex        =   14
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
            Left            =   3521
            TabIndex        =   13
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
            Left            =   4675
            TabIndex        =   12
            Top             =   0
            Width           =   1095
         End
      End
      Begin MSAdodcLib.Adodc datPrimaryRS 
         Height          =   330
         Left            =   120
         Top             =   630
         Width           =   5775
         _ExtentX        =   10186
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
         RecordSource    =   "select STU_NO,COL_NO,COURSE_NO,RES_NO,CANDIDATE_TYPE,[DOTE TYPE],QUOTA from JOIN_LIST"
         Caption         =   "Student Joing List"
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
   Begin VB.Frame Frame2 
      Caption         =   "Join List"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmJOIN_LIST.frx":0000
         DataField       =   "COL_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   765
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "COL_NAME"
         BoundColumn     =   "COL_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
         Bindings        =   "frmJOIN_LIST.frx":001A
         DataField       =   "COURSE_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1290
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "COURSE_NAME"
         BoundColumn     =   "COURSE_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmJOIN_LIST.frx":0033
         DataField       =   "STU_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   300
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "STU_NO"
         BoundColumn     =   "STU_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmJOIN_LIST.frx":004E
         DataField       =   "STU_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "STU_NO"
         BoundColumn     =   "STU_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmJOIN_LIST.frx":0063
         DataField       =   "RES_NO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "RES_NAME"
         BoundColumn     =   "RES_NO"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DataCombo6 
         Bindings        =   "frmJOIN_LIST.frx":0079
         DataField       =   "DOTE TYPE"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   1710
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "DOTE_TYPE"
         BoundColumn     =   "DOTE_TYPE"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DataCombo7 
         Bindings        =   "frmJOIN_LIST.frx":008C
         DataField       =   "CANDIDATE_TYPE"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   2235
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "CANDIDATE_TYPE"
         BoundColumn     =   "CANDIDATE_TYPE"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo DataCombo8 
         Bindings        =   "frmJOIN_LIST.frx":00A0
         DataField       =   "QUOTA"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   3300
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   12582912
         ListField       =   "COMMUNITY"
         BoundColumn     =   "COMMUNITY"
         Text            =   "DataCombo1"
         Object.DataMember      =   ""
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
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Quota:"
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
         Left            =   1335
         TabIndex        =   23
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label lblLabels 
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
         Index           =   5
         Left            =   930
         TabIndex        =   22
         Top             =   1770
         Width           =   1065
      End
      Begin VB.Label lblLabels 
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
         Index           =   4
         Left            =   420
         TabIndex        =   21
         Top             =   2295
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Reservation Name:"
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
         Left            =   150
         TabIndex        =   20
         Top             =   2820
         Width           =   1845
      End
      Begin VB.Label lblLabels 
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
         Index           =   2
         Left            =   660
         TabIndex        =   19
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label lblLabels 
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
         Index           =   1
         Left            =   645
         TabIndex        =   18
         Top             =   825
         Width           =   1350
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
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   2400
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT STU_NO FROM STUD WHERE STU_NO NOT IN(SELECT  STU_NO FROM JOIN_LIST)"
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
      Begin MSAdodcLib.Adodc CollegeName 
         Height          =   330
         Left            =   120
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT COL_NO,COL_NAME FROM COLLEGE WHERE COL_NO IN(SELECT  COL_NO FROM SEATS)"
         Caption         =   "CollegeName"
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
      Begin MSAdodcLib.Adodc CourseName 
         Height          =   330
         Left            =   120
         Top             =   2040
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT COURSE_NO,COURSE_NAME FROM COURSE WHERE COURSE_NO IN(SELECT  COURSE_NO FROM SEATS)"
         Caption         =   "CourseName"
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
      Begin MSAdodcLib.Adodc ResName 
         Height          =   330
         Left            =   120
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT * FROM SPL_RES"
         Caption         =   "ReservationName"
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
      Begin MSAdodcLib.Adodc Seat 
         Height          =   330
         Left            =   120
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT * FROM SEATS"
         Caption         =   "Seats"
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
      Begin MSAdodcLib.Adodc Dote 
         Height          =   330
         Left            =   120
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT DISTINCT DOTE_TYPE FROM SEATS"
         Caption         =   "Dote"
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
      Begin MSAdodcLib.Adodc Candi 
         Height          =   330
         Left            =   120
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT DISTINCT CANDIDATE_TYPE FROM SEATS"
         Caption         =   "Candidate"
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
      Begin MSAdodcLib.Adodc Quota 
         Height          =   330
         Left            =   120
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT DISTINCT COMMUNITY FROM SEATS"
         Caption         =   "Quota"
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
End
Attribute VB_Name = "frmJOIN_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddRec As Boolean

Private Sub Command1_Click()

End Sub

Private Sub DataCombo1_Click(Area As Integer)
Dim QryStr As String
        
QryStr = "SELECT COURSE_NAME,COURSE_NO FROM COURSE WHERE COURSE_NO IN " & _
         "(SELECT COURSE_NO FROM SEATS WHERE COL_NO LIKE '" & getCollegeNo(DataCombo1.Text) & "')"
            
ListCourse (QryStr)
End Sub

Private Sub DataCombo2_Click(Area As Integer)

Dim QryStr As String
QryStr = "SELECT DISTINCT DOTE_TYPE FROM SEATS WHERE " & _
         "COL_NO = '" & getCollegeNo(DataCombo1.Text) & "' AND " & _
         "COURSE_NO = '" & getCourseNo(DataCombo2.Text) & "'"

ListDoteType (QryStr)
End Sub

Private Sub DataCombo3_Click(Area As Integer)
Dim QryStr As String
QryStr = "SELECT DISTINCT COMMUNITY FROM SEATS WHERE " & _
        "COL_NO LIKE '" & getCollegeNo(DataCombo1.Text) & "' AND " & _
        "COURSE_NO LIKE '" & getCourseNo(DataCombo2.Text) & "' AND " & _
        "DOTE_TYPE LIKE '" & DataCombo6.Text & "' AND " & _
        "CANDIDATE_TYPE LIKE '" & DataCombo7.Text & "' AND " & _
        "RES_NO = '" & getResNo(DataCombo3.Text) & "'"

ListQuota (QryStr)

End Sub

Private Sub DataCombo6_Click(Area As Integer)
Dim QryStr As String

QryStr = "SELECT DISTINCT CANDIDATE_TYPE FROM SEATS WHERE " & _
        " COL_NO = '" & getCollegeNo(DataCombo1.Text) & "' AND" & _
        " COURSE_NO = '" & getCourseNo(DataCombo2.Text) & "' AND DOTE_TYPE = '" & DataCombo6.Text & "'"
ListCandidateType (QryStr)
End Sub

Private Sub DataCombo7_Click(Area As Integer)
Dim QryStr As String

QryStr = "SELECT RES_NAME,RES_NO FROM SPL_RES WHERE RES_NO IN (SELECT RES_NO FROM SEATS WHERE " & _
        "COL_NO LIKE '" & getCollegeNo(DataCombo1.Text) & "' AND " & _
        "COURSE_NO LIKE '" & getCourseNo(DataCombo2.Text) & "' AND " & _
        "DOTE_TYPE LIKE '" & DataCombo6.Text & "' AND " & _
        "CANDIDATE_TYPE LIKE '" & DataCombo7.Text & "')"

ListReservation (QryStr)

End Sub

Private Sub Form_Activate()
Dim QryStr As String

CheckSeats
QryStr = "SELECT COURSE_NAME,COURSE_NO FROM COURSE WHERE COURSE_NO IN " & _
        "(select course_no from seats where col_no in " & _
        "(select col_no from college  where col_name  like " & _
        "'" & DataCombo1.Text & "'" & "))"

ListCourse (QryStr)

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
  End Select
  
  If bCancel Then adStatus = adStatusCancel
  
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew
  DataCombo4.Visible = False
  DataCombo5.Visible = True
  AddRec = True
  Frame2.Enabled = True
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  Dim LastPos As Integer
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  CollegeName.Refresh
  CourseName.Refresh
  Adodc1.Refresh
  Call cmdRefresh_Click

  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  DataCombo4.Visible = True
  DataCombo5.Visible = False
  Frame2.Enabled = False
  Dim Qry As String

  Qry = "SELECT COURSE_NAME,COURSE_NO FROM COURSE WHERE COURSE_NO IN " & _
        "(select course_no from seats where col_no in " & _
        "(select col_no from college  where col_name  like " & _
        "'" & DataCombo1.Text & "'" & "))"

    ListCourse (Qry)

  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  
  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  UpdateSeats
  
  DataCombo4.Visible = True
  DataCombo5.Visible = False
  
  AddRec = False
  Adodc1.Refresh
  Frame2.Enabled = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Sub ListDoteType(Qry)
Dote.RecordSource = Qry
Dote.Recordset.Requery
Dote.Refresh
End Sub
Sub ListCandidateType(Qry)
Candi.RecordSource = Qry
Candi.Recordset.Requery
Candi.Refresh
End Sub
Sub ListReservation(Qry)
On Error Resume Next
With ResName
    .RecordSource = Qry
    .Recordset.Requery
    .Refresh
End With
End Sub
Sub ListQuota(Qry)
On Error Resume Next
With Quota
    .RecordSource = Qry
    .Recordset.Requery
    .Refresh
End With
End Sub

Sub ListCourse(Qry)

CourseName.RecordSource = Qry
CourseName.Refresh
DataCombo2.Refresh

End Sub

Sub UpdateSeats()
Dim findStr As String
Dim Rs As Recordset
findStr = "COL_NO LIKE '" & getCollegeNo(DataCombo1.Text) & "' AND "
findStr = findStr & "COURSE_NO LIKE '" & getCourseNo(DataCombo2.Text) & "' AND "
findStr = findStr & "DOTE_TYPE LIKE '" & DataCombo6.Text & "' AND "
findStr = findStr & "CANDIDATE_TYPE LIKE '" & DataCombo7.Text & "' AND "
findStr = findStr & "RES_NO LIKE '" & getResNo(DataCombo3.Text) & "' AND "
findStr = findStr & "COMMUNITY LIKE'" & DataCombo8.Text & "'"

CheckSeats

With DataEnvironment1.Connection1
    .Execute "UPDATE SEATS SET SEAT_ALLOC = SEAT_ALLOC - 1 WHERE " & _
     findStr
End With

End Sub


Private Sub Timer1_Timer()

End Sub


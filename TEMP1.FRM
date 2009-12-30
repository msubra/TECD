VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "College Course Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A55F47&
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   7335
      Begin VB.CommandButton ColCurRpt 
         Caption         =   "View Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Click of View Report button to view the Rank List of Eligible Students of TNPCEE Rank List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00833D8F&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rank List Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A55F47&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox List3 
         Height          =   255
         Left            =   6240
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton RnkRpt 
         Caption         =   "View Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   4920
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
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
         RecordSource    =   "RANK"
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
      Begin VB.Label Label1 
         Caption         =   "Click of View Report button to view the Rank List of Eligible Students of TNPCEE Rank List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00833D8F&
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Sub PutRank()
Dim i As Integer, j As Integer

For i = 0 To List1.ListCount
    
    If i > 0 And (List2.List(i) = List2.List(i - 1)) Then
        List3.AddItem i
        i = i + 1
    Else
        List3.AddItem i + 1
    End If
    
Next i

End Sub

Sub PrepareRankList()
Dim i As Integer, v As Variant

List1.Clear
List2.Clear
List3.Clear

With DataEnvironment1.rsCMD_MARKLIST

.Open
.Sort = "cutoff desc"

For i = 1 To .RecordCount
    List1.AddItem .Fields("stu_no")
    List2.AddItem .Fields("cutoff")
    .MoveNext
Next i
PutRank
.Close
End With

Adodc1.Refresh
With Adodc1.Recordset

While Not .EOF
    .Delete
    DoEvents
    .MoveNext
Wend

For i = 0 To List1.ListCount - 1
    .AddNew
    .Fields("stu_no") = List1.List(i)
    .Fields("rk") = List3.List(i)
    .Update
    DoEvents
Next i

End With

End Sub

Private Sub Command2_Click()

End Sub

Private Sub ColCurRpt_Click()
ColCourseRpt.Show
End Sub

Private Sub RnkRpt_Click()
PrepareRankList
RankList.Show

End Sub

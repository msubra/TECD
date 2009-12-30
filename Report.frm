VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F49365FC-E8A5-4E38-9DBC-DAA7D889B8A3}#1.6#0"; "pbxpbutton.ocx"
Begin VB.Form Rpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   4980
   ClientLeft      =   2310
   ClientTop       =   2130
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7470
   Begin VB.Frame Frame3 
      Caption         =   "Courses In Particular College"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   6975
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Report.frx":0000
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1050
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   8806739
         ListField       =   "COL_NAME"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc ColCour 
         Height          =   330
         Left            =   4800
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=COUNS"
         OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=COUNS"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "COLLEGE"
         Caption         =   "Adodc2"
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
      Begin PB_XP_Button.PBXPButton PBXPButton1 
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "View Report"
         BorderColorOver =   6956042
         Icon            =   "Report.frx":0016
         BorderColorDown =   6956042
         BackColor       =   13752539
         BackColorOver   =   13811126
         BackColorDown   =   16777215
         IconSizeWidth   =   30
         IconSizeHeight  =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         CheckedColor    =   14211029
      End
      Begin VB.Label Label3 
         Caption         =   "Click of View Report button to view the list of available college and available courses in those colleges"
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "List Courses in All Colleges Report"
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
      TabIndex        =   5
      Top             =   1560
      Width           =   6975
      Begin PB_XP_Button.PBXPButton PBXPButton3 
         Height          =   495
         Left            =   4920
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "View Report"
         BorderColorOver =   6956042
         Icon            =   "Report.frx":2B20
         BorderColorDown =   6956042
         BackColor       =   13752539
         BackColorOver   =   13811126
         BackColorDown   =   16777215
         IconSizeWidth   =   30
         IconSizeHeight  =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         CheckedColor    =   14211029
      End
      Begin VB.Label Label2 
         Caption         =   "Click of View Report button to view the list of available college and available courses in those colleges"
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
         Height          =   735
         Left            =   360
         TabIndex        =   6
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
      Width           =   6975
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox List3 
         Height          =   255
         Left            =   6240
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
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
      Begin PB_XP_Button.PBXPButton PBXPButton2 
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "View Report"
         BorderColorOver =   6956042
         Icon            =   "Report.frx":562A
         BorderColorDown =   6956042
         BackColor       =   13752539
         BackColorOver   =   13811126
         BackColorDown   =   16777215
         IconSizeWidth   =   30
         IconSizeHeight  =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowShadowOver  =   -1  'True
         CheckedColor    =   14211029
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
         TabIndex        =   1
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "Rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private Sub ColCurRpt_Click()
End Sub

Private Sub Command1_Click()
End Sub

Private Sub PBXPButton1_Click()
Dim rt As RptTextBox, rl As RptLabel
Dim Rs As Recordset

If DataCombo1.Text <> "" Then

Set Rs = getCourseFromCollege(DataCombo1.Text)
Set SingleColCOurse.DataSource = Rs
Set rt = SingleColCOurse.Sections("section1").Controls("curname")
Set rl = SingleColCOurse.Sections("section2").Controls("colname")

rt.DataMember = Rs.DataMember
rt.DataField = "course_name"
SingleColCOurse.Show
SingleColCOurse.WindowState = vbMaximized
rl.Caption = DataCombo1.Text
Else
    MsgBox "Select a college name", , "No College name"
End If
End Sub

Private Sub PBXPButton2_Click()
PrepareRankList
RankList.Show
RankList.WindowState = vbMaximized
End Sub

Private Sub RnkRpt_Click()

End Sub

Private Sub PBXPButton3_Click()
ColCourseRpt.Show
ColCourseRpt.WindowState = vbMaximized
End Sub

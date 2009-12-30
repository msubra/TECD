VERSION 5.00
Object = "{F49365FC-E8A5-4E38-9DBC-DAA7D889B8A3}#1.6#0"; "pbxpbutton.ocx"
Begin VB.MDIForm Couns 
   BackColor       =   &H8000000C&
   Caption         =   "Counselling"
   ClientHeight    =   6990
   ClientLeft      =   3000
   ClientTop       =   1245
   ClientWidth     =   7665
   Icon            =   "Couns.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   2400
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   7605
      TabIndex        =   13
      Top             =   0
      Width           =   7665
      Begin VB.Label Dt 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00866153&
         Height          =   285
         Left            =   5535
         TabIndex        =   15
         Top             =   -15
         Width           =   705
      End
      Begin VB.Label Td 
         AutoSize        =   -1  'True
         Caption         =   "Today:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00866153&
         Height          =   285
         Left            =   0
         TabIndex        =   14
         Top             =   -15
         Width           =   840
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   7605
      TabIndex        =   10
      Top             =   6015
      Width           =   7665
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Maheshwaran.S , Indumathi.R , Muthunathan.R , Jeevitha.M      - B.E II Year 'A' , Sona College Of Technology"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   10500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A Project By,"
         BeginProperty Font 
            Name            =   "Larabiefont"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D26B59&
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   5685
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   2160
      TabIndex        =   0
      Top             =   330
      Width           =   2220
      Begin PB_XP_Button.PBXPButton PBXPButton1 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   $"Couns.frx":4E12
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":4E25
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
         MultiLine       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton2 
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Join List"
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":9C47
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
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton3 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   4440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Reports"
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":EA69
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
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton4 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   $"Couns.frx":1388B
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":138A0
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
         MultiLine       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton5 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   $"Couns.frx":186C2
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":186E3
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
         MultiLine       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton6 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   $"Couns.frx":1D505
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":1D519
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
         MultiLine       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton7 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   $"Couns.frx":2233B
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":2234D
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
         MultiLine       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton8 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   $"Couns.frx":2716F
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":2718A
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
         MultiLine       =   -1  'True
         CheckedColor    =   14211029
      End
      Begin PB_XP_Button.PBXPButton PBXPButton10 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   5040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "Exit"
         BorderColorOver =   6956042
         Icon            =   "Couns.frx":2BFAC
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
         CheckedColor    =   14211029
      End
   End
End
Attribute VB_Name = "Couns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Td.Caption = "Today:  " & day(Date) & " " & Format(Date, "dddd") & "," & Format(Date, "MMMM") & " " & Year(Date)
Dt.Caption = "Time: " & Time
End Sub

Private Sub PBXPButton1_Click()
frmCollegeRec.Show
End Sub

Private Sub PBXPButton10_Click()
End
End Sub

Private Sub PBXPButton2_Click()
frmJOIN_LIST.Show
End Sub

Private Sub PBXPButton3_Click()
Rpt.Show
End Sub

Private Sub PBXPButton4_Click()
frmMARK_LIST.Show
End Sub

Private Sub PBXPButton5_Click()
frmSplResRec.Show
End Sub

Private Sub PBXPButton6_Click()
frmStudentRec.Show
End Sub

Private Sub PBXPButton7_Click()
frmCourseRec.Show
End Sub

Private Sub PBXPButton8_Click()
frmSeatAllocationRec.Show
End Sub

Private Sub Timer1_Timer()
Td.Caption = "Today:  " & day(Date) & " " & Format(Date, "dddd") & "," & Format(Date, "MMMM") & " " & Year(Date)
Dt.Caption = "Time: " & Time
End Sub


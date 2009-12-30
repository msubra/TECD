VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form CourseList 
   Caption         =   "Course List For Each College"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "CourseList.frx":0000
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   3
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      HighLight       =   0
      GridLinesUnpopulated=   3
      AllowUserResizing=   3
      DataMember      =   "COLLEGE_COURSE_LIST"
      RowSizingMode   =   1
      _NumberOfBands  =   3
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   2
      _Band(0)._MapCol(0)._Name=   "COL_NO"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "COL_NAME"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   1
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   4
      _Band(1)._MapCol(0)._Name=   "COL_NO"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "COL_NAME"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(1)._Hidden=   -1  'True
      _Band(1)._MapCol(2)._Name=   "COURSE_NO"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(2)._Hidden=   -1  'True
      _Band(1)._MapCol(3)._Name=   "COURSE_NAME"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(2).BandIndent=   2
      _Band(2).GridLinesBand=   1
      _Band(2).TextStyleBand=   0
      _Band(2).TextStyleHeader=   0
   End
End
Attribute VB_Name = "CourseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
With MSHFlexGrid1
    .Width = Me.Width - 200
    .Height = Me.Height - 600
End With
End Sub


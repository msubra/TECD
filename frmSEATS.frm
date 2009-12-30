VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSEATS 
   Caption         =   "SEATS"
   ClientHeight    =   5190
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   300
      Left            =   5340
      TabIndex        =   0
      Top             =   3960
      Width           =   1080
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   4860
      Width           =   6675
      _ExtentX        =   11774
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
      RecordSource    =   "COL_NO_COURSE_NO_DOTE_TYPE_RES_NAME"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmSEATS.frx":0000
      Height          =   3840
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   19
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      GridColor       =   12632256
      GridColorFixed  =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      MergeCells      =   4
      AllowUserResizing=   1
      FormatString    =   "COL_NO|DOTE_TYPE|CANDIDATE_TYPE|RES_NO|COURSE_NO|COMMUNITY|SEAT_ALLOC"
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   8
      _Band(0)._MapCol(0)._Name=   "COL_NO"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "COURSE_NO"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(1)._Hidden=   -1  'True
      _Band(0)._MapCol(2)._Name=   "COL_NAME"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "COURSE_NAME"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "DOTE_TYPE"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "RES_NAME"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "COMMUNITY"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "SEAT_ALLOC"
      _Band(0)._MapCol(7)._RSIndex=   7
   End
End
Attribute VB_Name = "frmSEATS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MARGIN_SIZE = 60      ' in Twips
Private Sub Form_Load()
    Dim i As Integer

    datPrimaryRS.Visible = False

    With MSHFlexGrid1

        .Redraw = False
        ' set grid's column widths
        For i = 0 To .Cols - 1
            .ColWidth(0) = -1
        Next i

        ' set grid's column merging and sorting
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i

        .Sort = flexSortGenericAscending

        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' make header bold
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' move Close button to the lower right corner
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' avoid error on negative values
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub



VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboBoxLB.ocx"
Begin VB.Form FrmMstSetRekapitulasiPersil 
   Caption         =   "Set Rekapitulasi Persil"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   9420
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7890
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9420
      _cx             =   16616
      _cy             =   13917
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         Height          =   7575
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   9015
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   495
            Left            =   4680
            TabIndex        =   8
            Top             =   6960
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   6120
            TabIndex        =   7
            Top             =   6960
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   495
            Left            =   3240
            TabIndex        =   6
            Top             =   6960
            Width           =   1335
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   495
            Left            =   7440
            TabIndex        =   5
            Top             =   6960
            Width           =   1335
         End
         Begin Combo.ComboBoxLB cboWilayah 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   6015
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   8535
            _cx             =   15055
            _cy             =   10610
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   13681305
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16744448
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   15525847
            GridColor       =   12100711
            GridColorFixed  =   10389833
            TreeColor       =   10389833
            FloodColor      =   192
            SheetBorder     =   -2147483645
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmMstSetRekapitulasiPersil.frx":0000
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblWilayah 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wilayah"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "FrmMstSetRekapitulasiPersil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Perintah As String
Private Sub ClearTxt()
    cboWilayah.Value = Null
    fg.Rows = 1
End Sub
Private Sub CmdEnabled(flag As Boolean)
    cmdSave.Enabled = flag
    cmdCancel.Enabled = flag
    cmdEdit.Enabled = Not flag
    cmdClose.Enabled = Not flag
End Sub
Private Sub ControlEnabled(flag As Boolean)
    fg.Enabled = flag
End Sub

Private Sub cmdCancel_Click()
    Perintah = ""
    ClearTxt
    CmdEnabled False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    If Trim(cboWilayah.Text) = "" Then
            MsgBox "Pilih Wilayah", vbExclamation, "Select User"
            Exit Sub
    End If
    fg.Enabled = True
    CmdEnabled True
End Sub

Private Sub Form_Load()
    Frame1.BackColor = MDIProject.ACPRibbon1.BackColor
    Call LoadCentreForm(Me)
    Call ControlCentreForm(Me, Frame1)
    'Call FormSize(8900, 9015, Me)
    Call LoadWilayah
    ClearTxt
    CmdEnabled (False)
    ControlEnabled (False)
End Sub

Private Sub LoadWilayah()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String
Dim i As Long
sql = "Select KodeWilayah, Wilayah from tblMstWilayah order by KodeWilayah Asc"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
If rs.RecordCount > 0 Then
    cboWilayah.ColumnCount = 2
    cboWilayah.ColumnWidths = "1000;0"
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        cboWilayah.AddItem rs!Wilayah & ";" & rs!KodeWilayah
        rs.MoveNext
    Next i
End If
End Sub

Private Sub cboWilayah_AfterUpdate()
    Call SetRekapitulasiPersil
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i As Integer
    Select Case Col
        Case 1
            For i = 1 To fg.Rows - 1
                If fg.ValueMatrix(i, 1) = True Then
                    fg.TextMatrix(i, 1) = "True"
                Else
                    fg.TextMatrix(i, 1) = "False"
                End If
            Next i
    End Select
End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     With fg
        Select Case Col
            Case 2, 3, 4, 5, 6
                Cancel = True
        End Select
    End With
End Sub

Private Sub SetRekapitulasiPersil()

    Dim rs      As ADODB.Recordset
    Dim Company As ADODB.Recordset
    Dim sql     As String
    Dim i, j As Long

    Set rs = New ADODB.Recordset
    sql = "select * from vwMstSetRekapitulasiPersil where KodeWilayah='" & cboWilayah.Column(1) & "' order by KodePersil Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    fg.Rows = 1

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        fg.Subtotal flexSTClear

        For i = 0 To rs.RecordCount - 1
            fg.AddItem ""
            fg.TextMatrix(i, 0) = i
            If rs!KodePersil <> "" Then
                fg.TextMatrix(i + 1, 1) = True
            End If
            fg.TextMatrix(i, 2) = rs!Wilayah
            fg.TextMatrix(i, 3) = rs!KodePersil
            fg.TextMatrix(i, 4) = rs!KoordinatX
            fg.TextMatrix(i, 5) = rs!KoordinatY
            rs.MoveNext
        Next i

    End If

    Set Company = New ADODB.Recordset
    sql = "select * from vwMstPersil where KodePersil not in(select KodePersil from vwMstSetRekapitulasiPersil where KodeWilayah='" & cboWilayah.Column(1) & "') order by KodePersil Asc"
    Company.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

    If Company.RecordCount > 0 Then
        Company.MoveFirst

        For j = fg.Rows To (Company.RecordCount + fg.Rows - 1)
            fg.AddItem ""
            fg.TextMatrix(j, 0) = j
            fg.TextMatrix(j, 1) = False
            fg.TextMatrix(j, 2) = Company!Wilayah
            fg.TextMatrix(j, 3) = Company!KodePersil
            fg.TextMatrix(j, 4) = Company!KoordinatX
            fg.TextMatrix(j, 5) = Company!KoordinatY
            Company.MoveNext
        Next j

    End If

    Set rs = Nothing

End Sub

VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTrnLevelAirSumurPantauFind 
   Caption         =   "Level Air Sumur Pantau Find"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20160
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
   ScaleHeight     =   8475
   ScaleWidth      =   20160
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   20160
      _cx             =   35560
      _cy             =   14949
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColor       =   -2147483646
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   15
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
      GridRows        =   3
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmTrnLevelAirSumurPantauFind.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Height          =   1020
         Left            =   570
         TabIndex        =   1
         Top             =   90
         Width           =   19185
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2040
            TabIndex        =   4
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   270204931
            CurrentDate     =   43143
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3960
            TabIndex        =   5
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   270204931
            CurrentDate     =   43143
         End
         Begin Combo.ComboBoxLB cboWilayah 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin MyLASP.isButton cmdRefresh 
            Height          =   495
            Left            =   5640
            TabIndex        =   10
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantauFind.frx":0077
            Style           =   5
            Caption         =   "&Refresh"
            IconSize        =   27
            CaptionAlign    =   2
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin VB.Label lblTanggal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periode"
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
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
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
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   285
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
            Left            =   480
            TabIndex        =   7
            Top             =   600
            Width           =   735
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   7050
         Left            =   570
         TabIndex        =   2
         Top             =   1335
         Width           =   6315
         _cx             =   11139
         _cy             =   12435
         Appearance      =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   400
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmTrnLevelAirSumurPantauFind.frx":0B97
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
         Editable        =   0
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
      Begin VSFlex8Ctl.VSFlexGrid fg2 
         Height          =   7050
         Left            =   7110
         TabIndex        =   3
         Top             =   1335
         Width           =   12645
         _cx             =   22304
         _cy             =   12435
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   11
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmTrnLevelAirSumurPantauFind.frx":0C50
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
   End
End
Attribute VB_Name = "FrmTrnLevelAirSumurPantauFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboWilayah_AfterUpdate()
'    Call LoadRekapitulasiHdr
End Sub

Private Sub cmdRefresh_Click()
'    If cboWilayah.Text = "" Then
'        MsgBox "Wilayah belum diisi, silahkan cek kembali data anda", vbInformation, AT
'        Exit Sub
'    End If
    Call LoadRekapitulasiHdr
End Sub

Private Sub DTPicker1_Change()
    'Call LoadRekapitulasiHdr
End Sub
Private Sub DTPicker2_Change()
    'Call LoadRekapitulasiHdr
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    fg2.WordWrap = True
    fg2.MergeCells = flexMergeFixedOnly
    fg2.MergeRow(0) = True
    fg2.MergeRow(1) = True
    
    For i = 0 To fg.Cols - 1
        fg2.MergeCol(i) = True
        fg2.FixedAlignment(i) = flexAlignCenterCenter
    Next i
    fg2.ColComboList(1) = "..."

    DTPicker1.Value = Format(GetDate, "yyyy-01-dd")
    DTPicker2.Value = Format(GetDate, "yyyy-12-dd")
    
    Call LoadWilayah(cboWilayah)
    
End Sub
Private Sub LoadRekapitulasiHdr()
    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    
    If DTPicker1.Value <> "" Or DTPicker2.Value <> "" Then
        sql = SetCondition(sql, "(Bulan between '" & Format(DTPicker1.Value, "MM") & "' and '" & Format(DTPicker2.Value, "MM") & "' and Tahun='" & Format(DTPicker2.Value, "yyyy") & "')")
    End If
    
    Set rs = New ADODB.Recordset
    
    If cboWilayah.Text = "" Then
        sql = "select HeaderID,KodeWilayah,Wilayah, Transdate,Weekday from vwRekapitulasiSumurPantau Where KodeWilayah like '" & IIf(MDIProject.GroupUser = "ADM WILAYAH", MDIProject.Wilayah, "") & "%' " & _
              IIf(sql = "", "", " AND " & sql) & _
              "Group by HeaderID,KodeWilayah,Transdate,Weekday,Wilayah order by KodeWilayah,Weekday Asc"
    Else
        sql = "select HeaderID,KodeWilayah,Wilayah, Transdate,Weekday from vwRekapitulasiSumurPantau Where KodeWilayah like '" & cboWilayah.Column(1) & "%' " & _
              IIf(sql = "", "", " AND " & sql) & _
              "Group by HeaderID,KodeWilayah,Transdate,Weekday,Wilayah order by KodeWilayah,Weekday Asc"
    End If
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg, 1)
    If rs.RecordCount > 0 Then
        fg.Rows = 1
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            fg.AddItem i + 1 & vbTab & rs!HeaderID & vbTab & rs!Wilayah & vbTab & Format(rs!Transdate, "yyyy-mm-dd") & vbTab & Format(rs!Transdate, "MMyyyy") & vbTab & rs!Weekday
            rs.MoveNext
        Next i
    End If

End Sub

Private Sub fg_DblClick()
'    If fg.Row <> 0 And fg.TextMatrix(fg.Row, 1) <> "" Then
'        Call FrmTrnLevelAirSumurPantau.LoadHeaderRekapitulasi(fg.TextMatrix(fg.Row, 1))
'        Unload Me
'    End If
End Sub

Private Sub fg_Click()
    If fg.Row <> 0 And fg.TextMatrix(fg.Row, 1) <> "" Then
        Call LoadDetailRekapitulasi(fg.TextMatrix(fg.Row, 1))
    End If
End Sub
Private Sub LoadDetailRekapitulasi(HeaderID As String)
    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim j As Integer
    Dim DataSebelum As String
    Set rs = New ADODB.Recordset
    sql = "Select HeaderID,DetailID,Wilayah,KodeAfdeling,KodePersil,KodeTitik,KoordinatX,KoordinatY,Nilai,Keterangan from vwRekapitulasiSumurPantau where HeaderID='" & HeaderID & "'order by DetailID Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg2, 2)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        With fg2
        .Rows = 3
            For i = 0 To rs.RecordCount - 1
                .AddItem ""
                .TextMatrix(i + 2, 0) = i + 1
                .TextMatrix(i + 2, 1) = rs!HeaderID & ""
                .TextMatrix(i + 2, 2) = rs!Wilayah & ""
                .TextMatrix(i + 2, 3) = rs!KodeAfdeling & ""
                .TextMatrix(i + 2, 4) = rs!KodePersil & ""
                .TextMatrix(i + 2, 5) = rs!KodeTitik
                .TextMatrix(i + 2, 6) = rs!KoordinatX & ""
                .TextMatrix(i + 2, 7) = rs!KoordinatY & ""
                .TextMatrix(i + 2, 8) = rs!Nilai & ""
                .TextMatrix(i + 2, 9) = rs!keterangan & ""
                .TextMatrix(i + 2, 10) = rs!DetailID & ""
                rs.MoveNext
            Next i
            fg2.Subtotal flexSTClear
            fg2.SubtotalPosition = flexSTBelow
            fg2.Subtotal flexSTAverage, -1, 8, "#,###.##", , , , "Jumlah Rata-Rata LASP"
        End With
    End If
End Sub

Private Sub fg2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With fg
        Select Case Col
            Case 2, 3, 4, 5, 6
                Cancel = True
        End Select
    End With
End Sub

Private Sub fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
   Dim i As Integer
   Select Case Col
        Case 1
             If fg2.Row <> 0 And fg2.TextMatrix(fg2.Row, 1) <> "" Then
                Call FrmTrnLevelAirSumurPantau.LoadHeaderRekapitulasi(fg.TextMatrix(fg.Row, 1))
                Unload Me
            End If
             
   End Select
End Sub

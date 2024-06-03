VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTrnLevelAirSumurPantau 
   Caption         =   "Level Air Sumur Pantau"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   16230
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16230
      _cx             =   28628
      _cy             =   16960
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
         Height          =   9285
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   15915
         Begin VB.Frame FraLevelAir 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Level Air Sumur Pantau"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1530
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   15585
            Begin VB.TextBox txtHeaderID 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2520
               TabIndex        =   4
               Text            =   "  "
               Top             =   600
               Width           =   855
            End
            Begin VB.TextBox txtWeek 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   8040
               TabIndex        =   3
               Top             =   960
               Width           =   615
            End
            Begin Combo.ComboBoxLB cboWilayah 
               Height          =   315
               Left            =   2520
               TabIndex        =   5
               Top             =   960
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               Appearance      =   0
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   330
               Left            =   8040
               TabIndex        =   6
               Top             =   600
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   98304001
               CurrentDate     =   43143
            End
            Begin VB.Label lblHeaderID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HeaderID"
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
               Left            =   720
               TabIndex        =   10
               Top             =   600
               Width           =   885
            End
            Begin VB.Label lblTanggal 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal"
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
               Left            =   6240
               TabIndex        =   9
               Top             =   600
               Width           =   765
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
               Left            =   720
               TabIndex        =   8
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblMingguKe 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Minggu Ke"
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
               Left            =   6240
               TabIndex        =   7
               Top             =   960
               Width           =   945
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   6615
            Left            =   120
            TabIndex        =   11
            Top             =   1800
            Width           =   15585
            _cx             =   27490
            _cy             =   11668
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
            Cols            =   12
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTrnLevelAirSumurPantau.frx":0000
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
         Begin MyLASP.isButton cmdFind 
            Height          =   495
            Left            =   4800
            TabIndex        =   12
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":01D1
            Style           =   5
            Caption         =   "&Find"
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
         Begin MyLASP.isButton CmdEntry 
            Height          =   495
            Left            =   6360
            TabIndex        =   13
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":0EC7
            Style           =   5
            Caption         =   "&Add"
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
         Begin MyLASP.isButton CmdEdit 
            Height          =   495
            Left            =   7920
            TabIndex        =   14
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":156D
            Style           =   5
            Caption         =   "&Edit"
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
         Begin MyLASP.isButton cmdSave 
            Height          =   495
            Left            =   9480
            TabIndex        =   15
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":1F65
            Style           =   5
            Caption         =   "&Save"
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
         Begin MyLASP.isButton cmdCancel 
            Height          =   495
            Left            =   12600
            TabIndex        =   16
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":2C2D
            Style           =   5
            Caption         =   "&Cancel"
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
         Begin MyLASP.isButton CmdClose 
            Height          =   495
            Left            =   14160
            TabIndex        =   17
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":3943
            Style           =   5
            Caption         =   "E&xit"
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
         Begin MyLASP.isButton cmdDelete 
            Height          =   495
            Left            =   11040
            TabIndex        =   18
            Top             =   8520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmTrnLevelAirSumurPantau.frx":44D7
            Style           =   5
            Caption         =   "&Delete"
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
         Begin VB.Label lblInputData 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "* Input Data Per 2 Minggu (Minggu 2 dan 4 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   8400
            Width           =   3855
         End
      End
   End
End
Attribute VB_Name = "FrmTrnLevelAirSumurPantau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Perintah As String

Private Sub Form_Load()

    Dim i As Integer

    fg.WordWrap = True
    fg.MergeCells = flexMergeFixedOnly
    fg.MergeRow(0) = True
    fg.MergeRow(1) = True

    For i = 0 To fg.Cols - 1
        fg.MergeCol(i) = True
        fg.FixedAlignment(i) = flexAlignCenterCenter
    Next i
    Frame1.BackColor = MDIProject.ACPRibbon1.BackColor
    Call FormSize(10065, 16915, Me)
    Call LoadCentreForm(Me)
    Call ControlCentreForm(Me, Frame1)
    ClearData
    ControlEnabled (False)
    CmdEnabled (True)
    Call LoadWilayah(cboWilayah)
    Call HLText(txtWeek)
    Call HLText(txtHeaderID)

End Sub

Private Sub Form_Resize()
    Call ControlCentreForm(Me, Frame1)
End Sub


Private Sub ControlEnabled(en As Boolean)
    txtHeaderID.Enabled = en
    DTPicker1.Enabled = en
    cboWilayah.Enabled = en
    txtWeek.Enabled = en
    fg.Enabled = en

End Sub

Private Sub CmdEnabled(flag As Boolean)
    CmdEntry.Enabled = flag
    CmdEdit.Enabled = flag
    cmdDelete.Enabled = flag
    cmdSave.Enabled = Not flag
    cmdCancel.Enabled = Not flag
    CmdClose.Enabled = flag

End Sub

Private Sub ClearData()
    txtHeaderID.Text = ""
    DTPicker1.Value = Format(GetDate, "dd/MM/yyyy")
    cboWilayah.Value = Null
    txtWeek.Text = ""
    Call HapusGrid(fg, 2)

End Sub

Private Sub cmdFind_Click()
    FrmTrnLevelAirSumurPantauFind.Show
End Sub

Private Sub cmdEntry_Click()
    ClearData
    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Add"
    txtHeaderID.Enabled = False
    cboWilayah.SetFocus
    ComputeWeekNo (DTPicker1)

End Sub

Private Sub cmdEdit_Click()

    If Trim(txtHeaderID.Text) = "" Then
        MsgBox "Silahkan CARI data yang akan di edit, lalu Klik Tombol EDIT", vbInformation, AT
        Exit Sub

    End If

    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Edit"
    txtHeaderID.Enabled = False
    cboWilayah.Enabled = False

End Sub

Private Sub cmdSave_Click()

    Dim cn       As New ADODB.Connection
    Dim cm       As New ADODB.Command
    Dim HeaderID As Integer
    Dim i        As Integer
    On Error GoTo ErrHandler

    Call cn.Open(ActiveCn)

    If cboWilayah.Text = "" Then
        MsgBox "Wilayah belum diisi, silahkan cek kembali data anda", vbInformation, AT
        Exit Sub

    End If

    cn.BeginTrans
    Me.MousePointer = vbHourglass
    HeaderID = SaveDataHdr(cn)

    If HeaderID = 0 Then
        Me.MousePointer = vbDefault
        cn.RollbackTrans
        Set cn = Nothing
        Call MsgBox("Process is failed (Hdr)!", vbExclamation, AT)
        Exit Sub

    End If

    With fg

        For i = 2 To .Rows - 1   'Di Simpan dari Row 2

            If .TextMatrix(i, 2) <> "" And .IsSubtotal(i) = False Then
                If SaveDataDtl(txtHeaderID, cn, i) = 0 Then
                    Me.MousePointer = vbDefault
                    cn.RollbackTrans
                    Set cn = Nothing
                    Call MsgBox("Process is failed !", vbExclamation, AT)
                    Exit Sub

                End If

            End If

        Next i

    End With

    cn.CommitTrans
    Me.MousePointer = vbDefault
    MsgBox "Data sudah tersimpan", vbInformation, AT
    ControlEnabled (False)
    CmdEnabled (True)
    LoadHeaderRekapitulasi (HeaderID)
    Set cn = Nothing
    Exit Sub
    '
ErrHandler:
    cn.RollbackTrans
    Call ErrMsg(Err)
    Me.MousePointer = vbDefault
    Set cm = Nothing
    Set cn = Nothing

End Sub

Private Sub cmdCancel_Click()
    ClearData
    CmdEnabled (True)
    ControlEnabled (False)
    Perintah = ""
    CmdEntry.SetFocus

End Sub


Private Sub cmdDelete_Click()

    On Error GoTo ErrorHandler
    Dim cn As New ADODB.Connection
    Call cn.Open(ActiveCn)
    Perintah = "Hapus"
    
    If Trim(txtHeaderID.Text) = "" Then
        MsgBox "Silahkan Cari Data dan Klik data yang akan di dihapus, lalu Klik Tombol Delete", vbInformation, AT
        Exit Sub
    End If
    
    cn.BeginTrans
    
    If MsgBox("Anda yakin mau menghapus data ini", vbYesNo) = vbYes Then
        cn.Execute "Delete tblTrnRekapitulasiHdr where HeaderID= '" & Trim(txtHeaderID.Text) & "'"
        cn.Execute "Delete tblTrnRekapitulasiDtl where HeaderID= '" & Trim(txtHeaderID.Text) & "'"
    Else
        cn.RollbackTrans
        Exit Sub
    End If
    
    cn.CommitTrans
    Perintah = ""
    MsgBox "Data sudah terhapus", vbInformation, AT
    ClearData
    Exit Sub

ErrorHandler:

    If Err.Number <> 0 Then
        cn.RollbackTrans
        MsgBox "Hapus Data Gagal : " & Err.Description

    End If

End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub cboWilayah_AfterUpdate()
    Call LoadPersil(cboWilayah.Column(1))
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Select Case Col
        Case 9
            fg.Subtotal flexSTClear
            fg.SubtotalPosition = flexSTBelow
            fg.Subtotal flexSTAverage, -1, 9, "#,###.##", , , , "Rata-Rata LASP"
    End Select
    
End Sub



Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With fg
        Select Case Col
            Case 1, 2, 3, 4, 5
                Cancel = True
        End Select
    End With

End Sub

Private Sub LoadPersil(KodeWilayah As String)
    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Set rs = New ADODB.Recordset
    sql = "Select KodeWilayah,Wilayah,KodeDivisi,NamaDivisi,KodePersil,KodeAfdeling,KodeTitik,KoordinatX,KoordinatY from vwMstPersil where KodeWilayah='" & KodeWilayah & "' order by KodePersil Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg, 2)
    If rs.RecordCount > 0 Then
        fg.Rows = 2
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1
                fg.AddItem ""
                fg.TextMatrix(i + 2, 0) = i + 1
                fg.TextMatrix(i + 2, 1) = rs!Wilayah & ""
                fg.TextMatrix(i + 2, 2) = rs!KodeAfdeling & ""
                fg.TextMatrix(i + 2, 3) = rs!KodeDivisi & ""
                fg.TextMatrix(i + 2, 4) = rs!NamaDivisi & ""
                fg.TextMatrix(i + 2, 5) = rs!KodePersil & ""
                fg.TextMatrix(i + 2, 6) = rs!KodeTitik & ""
                fg.TextMatrix(i + 2, 7) = rs!KoordinatX & ""
                fg.TextMatrix(i + 2, 8) = rs!KoordinatY & ""
            rs.MoveNext
        Next i
    End If
    fg.Subtotal flexSTClear
    fg.SubtotalPosition = flexSTBelow
    fg.Subtotal flexSTAverage, -1, 9, "#,###.##", , , , "Jumlah Rata-Rata LASP"


End Sub

Private Sub DTPicker1_Change()
    ComputeWeekNo (DTPicker1)
End Sub
Sub ComputeWeekNo(Tanggal As Date)
    Dim TheDate As Date, FirstDate As Date
        TheDate = CDate(Format(DTPicker1.Value, "dd/MM/yyyy"))
        FirstDate = CDate(Format("01/01/" & Right(DTPicker1.Value, 4), "dd/MM/yyyy"))
        If Weekday(TheDate) = 1 Then
           txtWeek.Text = DateDiff("ww", FirstDate, TheDate, , vbFirstFullWeek)
        Else
            txtWeek.Text = DateDiff("ww", FirstDate, TheDate, , vbFirstFullWeek) + 1
        End If
End Sub

Function SaveDataHdr(cn As ADODB.Connection) As Integer
    Dim cm    As New ADODB.Command
    Dim TglJT As Date
    SaveDataHdr = 0
    On Error GoTo ErrHandler
    cm.ActiveConnection = cn
    cm.CommandType = adCmdStoredProc
    cm.CommandText = "MyLASP..spLevelAirTanahHdr"

    cm.Parameters.Append cm.CreateParameter("@RETURN_VALUE", adInteger, adParamReturnValue)
    cm.Parameters.Append cm.CreateParameter("@Perintah", adVarChar, adParamInput, 10, Perintah)
    cm.Parameters.Append cm.CreateParameter("@HeaderID", adInteger, adParamInputOutput, , Val(txtHeaderID.Text))
    cm.Parameters.Append cm.CreateParameter("@Tanggal", adDBTimeStamp, adParamInput, , CDate(Format(DTPicker1.Value, "yyyy-mm-dd")))
    cm.Parameters.Append cm.CreateParameter("@Bulan", adInteger, adParamInput, , Format(DTPicker1.Value, "mm"))
    cm.Parameters.Append cm.CreateParameter("@Tahun", adInteger, adParamInput, , Format(DTPicker1.Value, "yyyy"))
    cm.Parameters.Append cm.CreateParameter("@KodeWilayah", adVarChar, adParamInput, 6, cboWilayah.Column(1))
    cm.Parameters.Append cm.CreateParameter("@WeekDay", adInteger, adParamInput, , CInt(txtWeek.Text))
    cm.Parameters.Append cm.CreateParameter("@CreatedBy", adVarChar, adParamInput, 50, MDIProject.UserID)
    cm.Parameters.Append cm.CreateParameter("@Flag", adInteger, adParamInputOutput, , 0)
    cm.Execute

    If cm.Parameters("@RETURN_VALUE") = 100 Then
         MsgBox "Level Air Tanah pada Minggu ke - ( " & CInt(txtWeek.Text) & " - " & cboWilayah.Column(0) & " ) sudah ada, " & vbNewLine & _
               "Silahkan Klik ( Find ) untuk melihat record Level Air Sumur Pantau."
        Call ErrMsg(Err)
        Me.MousePointer = vbDefault
        Set cm = Nothing
        Exit Function
    End If

    If cm.Parameters("@Flag") = 0 Then
        Call ErrMsg(Err)
        Me.MousePointer = vbDefault
        Set cm = Nothing
        Exit Function
    End If
    
    If cm.Parameters("@Flag") = 1 Then
        txtHeaderID.Text = cm.Parameters("@HeaderID")
        SaveDataHdr = 1
    End If

    Me.MousePointer = vbDefault
    

    Set cm = Nothing
    Exit Function

ErrHandler:
    Call ErrMsg(Err)
    Me.MousePointer = vbDefault
    Set cm = Nothing

End Function


Function SaveDataDtl(HeaderID As Integer, cn As ADODB.Connection, i As Integer) As Integer
    Dim cm      As New ADODB.Command
    Dim Remarks As String

    SaveDataDtl = 0
    On Error GoTo ErrHandler
    cm.ActiveConnection = cn
    cm.CommandType = adCmdStoredProc
    cm.CommandText = "MyLASP..spLevelAirTanahDtl"

    cm.Parameters.Append cm.CreateParameter("@DetailID", adInteger, adParamInput, , fg.ValueMatrix(i, 11))
    cm.Parameters.Append cm.CreateParameter("@HeaderID", adInteger, adParamInput, , txtHeaderID.Text)
    cm.Parameters.Append cm.CreateParameter("@KodeDivisi", adInteger, adParamInput, , Trim(fg.TextMatrix(i, 3)))
    cm.Parameters.Append cm.CreateParameter("@KodeAfdeling", adVarChar, adParamInput, 10, Trim(fg.TextMatrix(i, 2)))
    cm.Parameters.Append cm.CreateParameter("@KodePersil", adVarChar, adParamInput, 10, Trim(fg.TextMatrix(i, 5)))
    cm.Parameters.Append cm.CreateParameter("@KodeTitik", adVarChar, adParamInput, 10, Trim(fg.TextMatrix(i, 6)))
    cm.Parameters.Append cm.CreateParameter("@KoordinatX", adDouble, adParamInput, , CDbl(fg.ValueMatrix(i, 7)))
    cm.Parameters.Append cm.CreateParameter("@KoordinatY", adDouble, adParamInput, , CDbl(fg.ValueMatrix(i, 8)))
    cm.Parameters.Append cm.CreateParameter("@Nilai", adDouble, adParamInput, , IIf(fg.TextMatrix(i, 9) = "", Null, fg.ValueMatrix(i, 9)))
    cm.Parameters.Append cm.CreateParameter("@Keterangan", adVarChar, adParamInput, 100, fg.TextMatrix(i, 10))
    cm.Parameters.Append cm.CreateParameter("@CreatedBy", adVarChar, adParamInput, 50, MDIProject.UserID)
    cm.Parameters.Append cm.CreateParameter("@Flag", adInteger, adParamInputOutput, , 0)
    cm.Execute
    
    If cm.Parameters("@Flag") <> 0 Then SaveDataDtl = 1
    Set cm = Nothing
    Exit Function

ErrHandler:
    Call ErrMsg(Err)
    Me.MousePointer = vbDefault
    Set cm = Nothing

End Function

Sub LoadHeaderRekapitulasi(HeaderID As String)
Dim rs As New ADODB.Recordset
Dim sql As String

sql = "Select * from MyLASP..vwRekapitulasiSumurPantau where HeaderID='" & HeaderID & "'"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

If rs.RecordCount > 0 Then
    rs.MoveFirst
    txtHeaderID.Text = rs!HeaderID
    DTPicker1.Value = Format(rs!Transdate, "dd/MM/yyyy")
    cboWilayah.Value = rs!Wilayah
    txtWeek.Text = rs!Weekday
    Call LoadDetailRekapitulasi(HeaderID)

End If
Set rs = Nothing
End Sub


Private Sub LoadDetailRekapitulasi(HeaderID As String)
    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim j As Integer
    Dim DataSebelum As String
    Set rs = New ADODB.Recordset
    sql = "Select DetailID,KodeWilayah,Wilayah,KodeDivisi,NamaDivisi,KodePersil,KodeAfdeling,KodeTitik,KoordinatX,KoordinatY,Nilai,Keterangan from vwRekapitulasiSumurPantau where HeaderID='" & HeaderID & "'order by DetailID Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg, 2)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        With fg
        .Rows = 2
            For i = 0 To rs.RecordCount - 1
                .AddItem ""
                .TextMatrix(i + 2, 0) = i + 1
                .TextMatrix(i + 2, 1) = rs!Wilayah & ""
                .TextMatrix(i + 2, 2) = rs!KodeAfdeling & "" 'Kolom Divisi = KodeAfdeling
                .TextMatrix(i + 2, 3) = rs!KodeDivisi & ""
                .TextMatrix(i + 2, 4) = rs!NamaDivisi & ""
                .TextMatrix(i + 2, 5) = rs!KodePersil & ""
                .TextMatrix(i + 2, 6) = rs!KodeTitik & ""
                .TextMatrix(i + 2, 7) = rs!KoordinatX & ""
                .TextMatrix(i + 2, 8) = rs!KoordinatY & ""
                .TextMatrix(i + 2, 9) = rs!Nilai & ""
                .TextMatrix(i + 2, 10) = rs!keterangan & ""
                .TextMatrix(i + 2, 11) = rs!DetailID & ""
                rs.MoveNext
            Next i
            fg.Subtotal flexSTClear
            fg.SubtotalPosition = flexSTBelow
            fg.Subtotal flexSTAverage, -1, 9, "#,###.##", , , , "Jumlah Rata-Rata LASP"
        End With
    End If
End Sub


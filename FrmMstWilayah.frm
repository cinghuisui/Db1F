VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmMstWilayah 
   Caption         =   "Master Wilayah"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11280
      _cx             =   19897
      _cy             =   12965
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
         BackColor       =   &H00FFFFFF&
         Caption         =   "Master Wilayah"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   10695
         Begin MyLASP.isButton cmdFind 
            Height          =   495
            Left            =   720
            TabIndex        =   13
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstWilayah.frx":0000
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
         Begin VB.CheckBox ChkNotActive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Not Active"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3360
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtWilayah 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3360
            TabIndex        =   2
            Text            =   "  "
            Top             =   1005
            Width           =   3255
         End
         Begin VB.TextBox txtKodeWilayah 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3360
            TabIndex        =   1
            Text            =   "  "
            Top             =   600
            Width           =   1815
         End
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   3540
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   10245
            _cx             =   18071
            _cy             =   6244
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            BackColorFixed  =   16766894
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   13882323
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmMstWilayah.frx":0CF6
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
            ExplorerBar     =   3
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
         Begin MyLASP.isButton CmdEntry 
            Height          =   495
            Left            =   2280
            TabIndex        =   8
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstWilayah.frx":0E00
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
            Left            =   3840
            TabIndex        =   9
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstWilayah.frx":14A6
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
            Left            =   5400
            TabIndex        =   10
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstWilayah.frx":1E9E
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
            Left            =   6960
            TabIndex        =   11
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstWilayah.frx":2B66
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
         Begin MyLASP.isButton cmdExit 
            Height          =   495
            Left            =   8520
            TabIndex        =   12
            Top             =   6000
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstWilayah.frx":387C
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wilayah"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   7
            Top             =   1005
            Width           =   780
         End
         Begin VB.Label lblHeaderID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Wilayah"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   6
            Top             =   600
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "FrmMstWilayah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Perintah As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Frame1.BackColor = MDIProject.ACPRibbon1.BackColor
    'C1Elastic1.BackColor = MDIProject.ACPRibbon1.BackColor
    ChkNotActive.BackColor = MDIProject.ACPRibbon1.BackColor
    Call ControlCentreForm(Me, Frame1)
    Call FormSize(8695, 11610, Me)
    Call LoadCentreForm(Me)
    Call HLText(txtKodeWilayah)
    ClearData
    ControlEnabled (False)
    CmdEnabled (True)
    LoadMasterWilayah

End Sub

Private Sub Form_Resize()
    Call ControlCentreForm(Me, Frame1)
End Sub

Private Sub ControlEnabled(en As Boolean)
    txtKodeWilayah.Enabled = en
    txtWilayah.Enabled = en
    ChkNotActive.Enabled = en

End Sub

Private Sub CmdEnabled(flag As Boolean)
    CmdEntry.Enabled = flag
    CmdEdit.Enabled = flag
    cmdSave.Enabled = Not flag
    cmdCancel.Enabled = Not flag
    cmdExit.Enabled = flag

End Sub

Private Sub ClearData()
    txtKodeWilayah.Text = ""
    txtWilayah.Text = ""
    Perintah = ""
End Sub

Private Sub cmdCancel_Click()
    ClearData
    ControlEnabled (False)
    CmdEnabled (True)
    CmdEntry.SetFocus
    fg.Enabled = True

End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()

    If Trim(txtKodeWilayah.Text) = "" Then
        MsgBox "Silahkan CARI data yang akan di edit, lalu Klik Tombol EDIT", vbInformation, AT
        Exit Sub

    End If

    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Edit"
    txtKodeWilayah.Enabled = False

End Sub

Private Sub cmdEntry_Click()
    ClearData
    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Add"
    txtKodeWilayah.SetFocus

End Sub

Private Sub fg_Click()

    If Perintah = "" Then
        If fg.Row <> 0 Then
            txtKodeWilayah.Text = fg.TextMatrix(fg.Row, 1)
            txtWilayah.Text = fg.TextMatrix(fg.Row, 2)
            ChkNotActive.Value = IIf(fg.TextMatrix(fg.Row, 3) = True, 1, 0)

        End If

    End If

End Sub

Private Sub LoadMasterWilayah()

    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim i   As Integer
    Dim j   As Integer

    Set rs = New ADODB.Recordset
    sql = "Select KodeWilayah,Wilayah,NotActive,CreatedBy,CreatedDate,LastUpdatedBy,LastUpdatedDate from tblMstWilayah order by KodeWilayah Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg, 1)

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        With fg
            .Rows = 1

            For i = 0 To rs.RecordCount - 1
                .AddItem ""
                .TextMatrix(i + 1, 0) = i + 1
                .TextMatrix(i + 1, 1) = rs!KodeWilayah & ""
                .TextMatrix(i + 1, 2) = rs!Wilayah & ""

                If rs!NotActive = True Then
                    .TextMatrix(i + 1, 3) = True
                Else
                    .TextMatrix(i + 1, 3) = False
                End If

                If Not IsNull(rs!CreatedBy) Then .TextMatrix(i + 1, 4) = rs!CreatedBy & ""
                If Not IsNull(rs!CreatedDate) Then .TextMatrix(i + 1, 5) = Format(rs!CreatedDate, "dd/mm/yyyy") & ""
                If Not IsNull(rs!LastUpdatedBy) Then .TextMatrix(i + 1, 6) = rs!LastUpdatedBy & ""
                If Not IsNull(rs!LastUpdatedDate) Then .TextMatrix(i + 1, 7) = Format(rs!LastUpdatedDate, "dd/mm/yyyy") & ""
                rs.MoveNext
            Next i

        End With

    End If

End Sub

Private Sub cmdSave_Click()

    On Error GoTo EH

    Dim rs As ADODB.Recordset

    Dim cn As New ADODB.Connection
    
    If txtKodeWilayah.Text = "" Then
        MsgBox "Kode Wilayah harap diisi!..", vbInformation, AT
        txtKodeWilayah.SetFocus
        Exit Sub

    End If
    
    If txtWilayah.Text = "" Then
        MsgBox "Wilayah harap diisi!..", vbInformation, AT
        txtWilayah.SetFocus
        Exit Sub

    End If
    
    Call cn.Open(ActiveCn)
    Set rs = New ADODB.Recordset

    Dim sql As String

    sql = "select * from tblMstWilayah where KodeWilayah='" & Trim(txtKodeWilayah.Text) & "'"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    
    If Perintah = "ADD" Then
        If Not rs.EOF Then 'SUDAH ADA
            MsgBox "Kode: " & txtKodeWilayah.Text & " sudah ada, cek list Master Wilayah.", vbInformation, AT
            txtKodeWilayah.SetFocus
            Exit Sub

        End If

    End If
    
    If rs.EOF = True Then
        rs.AddNew
        rs!KodeWilayah = UCase(txtKodeWilayah.Text)
        rs!CreatedBy = MDIProject.UserID
        rs!CreatedDate = Format(GetDate(), "yyyy-mm-dd")

    End If

    rs!Wilayah = Trim(txtWilayah.Text)
    rs!LastUpdatedBy = MDIProject.UserID
    rs!LastUpdatedDate = Format(GetDate(), "yyyy-mm-dd")
    rs!NotActive = IIf(ChkNotActive.Value = 1, True, False)
    rs.Update
    MsgBox "Master Wilayah Berhasil Disimpan.", vbInformation, AT
    ShowData
    ControlEnabled False
    CmdEnabled True
    Set rs = Nothing
    LoadMasterWilayah
    Exit Sub
EH:
    ClearData
    ControlEnabled True
    CmdEnabled False
    Set rs = Nothing
    Exit Sub

End Sub

Public Sub ShowData()
    Dim rs  As ADODB.Recordset
    Dim cn  As New ADODB.Connection
    Dim sql As String
    Dim i   As Integer
    On Error GoTo EH
    
    Set rs = New ADODB.Recordset
    sql = "Select * from MyLASP..tblMstWilayah Where KodeWilayah='" & txtKodeWilayah.Text & "'"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call ClearData
    rs.MoveFirst
    
    txtKodeWilayah = rs!KodeWilayah
    txtWilayah = rs!Wilayah
    ChkNotActive = IIf(rs!NotActive = True, 1, 0)
    Set rs = Nothing
    Exit Sub

EH:
    Set rs = Nothing
    Call ErrMsg(Err)
End Sub

Private Sub jcbutton1_Click()

End Sub

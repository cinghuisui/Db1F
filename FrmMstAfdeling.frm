VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Begin VB.Form FrmMstAfdeling 
   Caption         =   "Master Afdeling"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11295
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7125
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   11295
      _cx             =   19923
      _cy             =   12568
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
         Caption         =   "Master Afdeling"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtKodeAfdeling 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2880
            TabIndex        =   0
            Text            =   "  "
            Top             =   720
            Width           =   1815
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
            Height          =   375
            Left            =   2880
            TabIndex        =   3
            Top             =   1800
            Width           =   1335
         End
         Begin Combo.ComboBoxLB cboWilayah 
            Height          =   315
            Left            =   2880
            TabIndex        =   1
            Top             =   1110
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   3420
            Left            =   240
            TabIndex        =   6
            Top             =   2520
            Width           =   10365
            _cx             =   18283
            _cy             =   6032
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmMstAfdeling.frx":0000
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
         Begin Combo.ComboBoxLB cboDivisi 
            Height          =   315
            Left            =   2880
            TabIndex        =   2
            Top             =   1450
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin MyLASP.isButton cmdFind 
            Height          =   495
            Left            =   840
            TabIndex        =   10
            Top             =   6120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstAfdeling.frx":011B
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
            Left            =   2400
            TabIndex        =   11
            Top             =   6120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstAfdeling.frx":0E11
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
            Left            =   3960
            TabIndex        =   12
            Top             =   6120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstAfdeling.frx":14B7
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
            Left            =   5520
            TabIndex        =   13
            Top             =   6120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstAfdeling.frx":1EAF
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
            Left            =   7080
            TabIndex        =   14
            Top             =   6120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstAfdeling.frx":2B77
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
            Left            =   8640
            TabIndex        =   15
            Top             =   6120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            Icon            =   "FrmMstAfdeling.frx":388D
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
         Begin VB.Label lblHeaderID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Afdeling"
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
            Left            =   1320
            TabIndex        =   9
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Divisi"
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
            Left            =   1320
            TabIndex        =   8
            Top             =   1440
            Width           =   525
         End
         Begin VB.Label Label2 
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
            Left            =   1320
            TabIndex        =   7
            Top             =   1080
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "FrmMstAfdeling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Perintah As String

Private Sub cboWilayah_AfterUpdate()
    Call LoadDivisi(cboWilayah.Column(1))
End Sub


Private Sub Form_Load()
    Frame1.BackColor = MDIProject.ACPRibbon1.BackColor
    ChkNotActive.BackColor = MDIProject.ACPRibbon1.BackColor
    Call ControlCentreForm(Me, Frame1)
    Call FormSize(8695, 11610, Me)
    Call LoadCentreForm(Me)
    Call HLText(txtKodeAfdeling)
    ClearData
    ControlEnabled (False)
    CmdEnabled (True)
    LoadWilayah
    LoadMasterAfdeling
    
End Sub
Private Sub Form_Resize()
    Call ControlCentreForm(Me, Frame1)
End Sub
Private Sub ControlEnabled(en As Boolean)
    txtKodeAfdeling.Enabled = en
    cboDivisi.Enabled = en
    cboWilayah.Enabled = en
    ChkNotActive.Enabled = en
    
End Sub

Private Sub CmdEnabled(flag As Boolean)
    CmdEntry.Enabled = flag
    CmdEdit.Enabled = flag
    cmdSave.Enabled = Not flag
    cmdCancel.Enabled = Not flag
    CmdClose.Enabled = flag
    
End Sub
Private Sub ClearData()
    txtKodeAfdeling.Text = ""
    cboDivisi.Value = Null
    cboWilayah.Value = Null
    ChkNotActive.Value = 0
    Perintah = ""
    
End Sub

Private Sub cmdEntry_Click()
    ClearData
    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Add"
    txtKodeAfdeling.SetFocus
    fg.Enabled = False
    
End Sub

Private Sub cmdEdit_Click()

    If Trim(txtKodeAfdeling.Text) = "" Then
        MsgBox "Silahkan Klik List Afdeling, lalu Klik Tombol EDIT", vbInformation, AT
        Exit Sub

    End If

    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Edit"
    txtKodeAfdeling.Enabled = False
    fg.Enabled = False

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

Private Sub fg_Click()

    If Perintah = "" Then
        If fg.Row <> 0 Then
            txtKodeAfdeling.Text = fg.TextMatrix(fg.Row, 1)
            cboDivisi.Text = fg.TextMatrix(fg.Row, 2)
            cboWilayah.Text = fg.TextMatrix(fg.Row, 3)
            ChkNotActive.Value = IIf(fg.TextMatrix(fg.Row, 4) = True, 1, 0)

        End If

    End If

End Sub

Private Sub cmdSave_Click()

    On Error GoTo EH
    Dim rs As ADODB.Recordset
    Dim cn As New ADODB.Connection
    
    If txtKodeAfdeling.Text = "" Then
        MsgBox "Kode Afdeling harap diisi!..", vbInformation, AT
        txtKodeAfdeling.SetFocus
        Exit Sub

    End If
    
    If cboDivisi.Text = "" Then
        MsgBox "Divisi harap diisi!..", vbInformation, AT
        cboDivisi.SetFocus
        Exit Sub

    End If
    
    If cboWilayah.Text = "" Then
        MsgBox "Wilayah harap diisi!..", vbInformation, AT
        cboWilayah.SetFocus
        Exit Sub

    End If
    
    Call cn.Open(ActiveCn)
    Set rs = New ADODB.Recordset

    Dim sql As String

    sql = "select * from tblMstAfdeling where KodeAfdeling='" & Trim(txtKodeAfdeling.Text) & "'"
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    
    If Perintah = "ADD" Then
        If Not rs.EOF Then 'SUDAH ADA
            MsgBox "Kode: " & txtKodeAfdeling.Text & " sudah ada, Cek list Master Afdeling.", vbInformation, AT
            txtKodeAfdeling.SetFocus
            Exit Sub

        End If

    End If
    
    If rs.EOF = True Then
        rs.AddNew
        rs!KodeAfdeling = Trim(txtKodeAfdeling.Text)
        rs!CreatedBy = MDIProject.UserID
        rs!CreatedDate = Format(GetDate(), "yyyy-mm-dd")

    End If

    rs!Divisi = cboDivisi.Column(1)
    rs!Wilayah = cboWilayah.Column(1)
    rs!LastUpdatedBy = MDIProject.UserID
    rs!LastUpdatedDate = Format(GetDate(), "yyyy-mm-dd")
    rs!NotActive = IIf(ChkNotActive.Value = 1, True, False)
    rs.Update
    MsgBox "Master Afdeling Berhasil Disimpan.", vbInformation, AT
    ClearData
    ControlEnabled False
    CmdEnabled True
    fg.Enabled = True
    Set rs = Nothing
    LoadMasterAfdeling
    Exit Sub
EH:
    ClearData
    ControlEnabled True
    CmdEnabled False
    Set rs = Nothing
    Exit Sub

End Sub

Private Sub LoadWilayah()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim I   As Long

    sql = "Select KodeWilayah, Wilayah from tblMstWilayah order by KodeWilayah Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

    If rs.RecordCount > 0 Then
        cboWilayah.ColumnCount = 2
        cboWilayah.ColumnWidths = "1000;0"
        rs.MoveFirst

        For I = 0 To rs.RecordCount - 1
            cboWilayah.AddItem rs!Wilayah & ";" & rs!KodeWilayah
            rs.MoveNext
        Next I

    End If

End Sub

Private Sub LoadDivisi(KodeWilayah As String)

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim I   As Long

    sql = "Select KodeDivisi, NamaDivisi from vwMstDivisi where KodeWilayah='" & KodeWilayah & "' order by KodeDivisi Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

    If rs.RecordCount > 0 Then
        cboDivisi.ColumnCount = 2
        cboDivisi.ColumnWidths = "1000;0"
        cboDivisi.Value = Null
        cboDivisi.Clear
        rs.MoveFirst

        For I = 0 To rs.RecordCount - 1
            cboDivisi.AddItem rs!NamaDivisi & ";" & rs!KodeDivisi
            rs.MoveNext
        Next I

    End If

End Sub

Private Sub LoadMasterAfdeling()

    Dim rs          As ADODB.Recordset
    Dim sql         As String
    Dim I           As Integer
    Dim j           As Integer


    Set rs = New ADODB.Recordset
    sql = "Select KodeAfdeling,NamaDivisi,Wilayah,NotActive,CreatedBy,CreatedDate,LastUpdatedBy,LastUpdatedDate from vwMstAfdeling order by KodeAfdeling Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg, 1)

    If rs.RecordCount > 0 Then
        rs.MoveFirst

        With fg
            .Rows = 1

            For I = 0 To rs.RecordCount - 1
                .AddItem ""
                .TextMatrix(I + 1, 0) = I + 1
                .TextMatrix(I + 1, 1) = rs!KodeAfdeling & ""
                .TextMatrix(I + 1, 2) = rs!NamaDivisi & ""
                .TextMatrix(I + 1, 3) = rs!Wilayah & ""
                If rs!NotActive = True Then
                    .TextMatrix(I + 1, 4) = True
                Else
                    .TextMatrix(I + 1, 4) = False
                End If
                 If Not IsNull(rs!CreatedBy) Then .TextMatrix(I + 1, 5) = rs!CreatedBy & ""
                 If Not IsNull(rs!CreatedDate) Then .TextMatrix(I + 1, 6) = Format(rs!CreatedDate, "dd/mm/yyyy") & ""
                 If Not IsNull(rs!LastUpdatedBy) Then .TextMatrix(I + 1, 7) = rs!LastUpdatedBy & ""
                 If Not IsNull(rs!LastUpdatedDate) Then .TextMatrix(I + 1, 8) = Format(rs!LastUpdatedDate, "dd/mm/yyyy") & ""
                rs.MoveNext
            Next I

        End With

    End If

End Sub


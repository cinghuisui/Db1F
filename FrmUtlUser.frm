VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUtlUser 
   Caption         =   "User"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7950
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
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   7980.695
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8640
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7950
      _cx             =   14023
      _cy             =   15240
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
         Caption         =   "U S E R   Administration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   7575
         Begin VB.TextBox linkttd 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   29
            Top             =   2280
            Width           =   5415
         End
         Begin VB.PictureBox jcFrames1 
            BackColor       =   &H00F0D4C0&
            FillColor       =   &H00F0D4C0&
            Height          =   1935
            Left            =   4560
            ScaleHeight     =   1875
            ScaleWidth      =   2835
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   360
            Width           =   2895
            Begin VB.CommandButton cmdLoadTanda 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Load Tanda Tangan"
               Height          =   375
               Left            =   0
               MaskColor       =   &H00404040&
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               Width           =   2895
            End
            Begin VB.Image Image1 
               Height          =   1215
               Left            =   120
               Stretch         =   -1  'True
               Top             =   480
               Width           =   2535
            End
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00FF8080&
            Caption         =   "Save"
            Height          =   375
            Left            =   4200
            TabIndex        =   8
            Top             =   7560
            Width           =   1005
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5205
            TabIndex        =   9
            Top             =   7560
            Width           =   870
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF8080&
            Caption         =   "Exit"
            Height          =   375
            Left            =   6060
            TabIndex        =   10
            Top             =   7560
            Width           =   900
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00FF8080&
            Caption         =   "Edit"
            Height          =   375
            Left            =   2415
            TabIndex        =   6
            Top             =   7560
            Width           =   870
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00FF8080&
            Caption         =   "Delete"
            Height          =   375
            Left            =   3270
            TabIndex        =   7
            Top             =   7560
            Width           =   915
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FF8080&
            Caption         =   "Add"
            Height          =   375
            Left            =   1560
            TabIndex        =   0
            Top             =   7560
            Width           =   870
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            TabIndex        =   2
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton cmdResetPassword 
            Caption         =   "Reset Password"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   7560
            Width           =   1455
         End
         Begin Combo.ComboBoxLB cboGroupAkses 
            Height          =   315
            Left            =   2040
            TabIndex        =   5
            Top             =   1560
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin Combo.ComboBoxLB ComboBoxLB1 
            Height          =   315
            Left            =   2040
            TabIndex        =   4
            Top             =   1200
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin Combo.ComboBoxLB CboKodeWil 
            Height          =   315
            Left            =   2040
            TabIndex        =   11
            Top             =   1920
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin TabDlg.SSTab SSTDaftarMenu 
            Height          =   4695
            Left            =   120
            TabIndex        =   19
            Top             =   2760
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   8281
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   4210752
            TabCaption(0)   =   "Daftar Menu"
            TabPicture(0)   =   "FrmUtlUser.frx":0000
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fg1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Daftar User"
            TabPicture(1)   =   "FrmUtlUser.frx":001C
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "fg2"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Text3"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "cmdCariUsers"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).ControlCount=   3
            Begin VB.CommandButton cmdCariUsers 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cari User"
               Height          =   315
               Left            =   5280
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2520
               TabIndex        =   15
               Top             =   600
               Width           =   2655
            End
            Begin VB.CommandButton cmdCariUser 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cari User"
               Height          =   315
               Left            =   -69600
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   540
               Width           =   1335
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   -72360
               TabIndex        =   20
               Top             =   540
               Width           =   2655
            End
            Begin VSFlex8Ctl.VSFlexGrid fg2s 
               Height          =   3555
               Left            =   -74880
               TabIndex        =   22
               Top             =   900
               Width           =   6615
               _cx             =   11668
               _cy             =   6271
               Appearance      =   1
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
               BackColorFixed  =   14664068
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12121338
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
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmUtlUser.frx":0038
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
            Begin VSFlex8Ctl.VSFlexGrid fg1 
               Height          =   3975
               Left            =   -74880
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   480
               Width           =   6615
               _cx             =   11668
               _cy             =   7011
               Appearance      =   1
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
               BackColorFixed  =   13295785
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   10218222
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
               GridLines       =   3
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmUtlUser.frx":0144
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
               Height          =   3795
               Left            =   120
               TabIndex        =   24
               Top             =   960
               Width           =   7095
               _cx             =   12515
               _cy             =   6694
               Appearance      =   1
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
               BackColorFixed  =   14664068
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12121338
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
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmUtlUser.frx":01AA
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
         End
         Begin MSComDlg.CommonDialog cmDLG 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lblLinkParaf 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Link Paraf"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UserName"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User ID"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   540
         End
         Begin VB.Label lblWilayah 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wilayah"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblGroupAkses 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group Akses"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jabatan"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "FrmUtlUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Perintah As String

Private Sub cmdCariUsers_Click()
Dim rs As ADODB.Recordset
Dim sql As String
Dim i As Integer

Set rs = New ADODB.Recordset
sql = "Select UserID,Nama,JabatanName,NamaGroup,KodeWilayah from vwUser where Nama like '%" & Trim(Text3.Text) & "%'"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

Call HapusGrid(fg2, 1)
If rs.RecordCount > 0 Then
    fg2.Rows = 1
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        fg2.AddItem rs!UserID & vbTab & rs!nama & vbTab & rs!JabatanName & vbTab & rs!NamaGroup & vbTab & rs!KodeWilayah
        rs.MoveNext
    Next i
Else
    fg2.Rows = 1
End If
End Sub

Private Sub cmdDelete_Click()
Dim cn As New ADODB.Connection

Call cn.Open(ActiveCn)
If Trim(Text1.Text) = "" Then
    MsgBox "Tidak ada Data yang mau dihapus , silahkan cari data terlebih dahulu di Tabel di bawah ini", vbInformation, "Find"
    Exit Sub
End If

If MsgBox("Anda yakin mau menghapus data ini", vbYesNo) = vbYes Then
   cn.Execute "delete from tblUtlUser where UserID = '" & UCase(Trim(Text1.Text)) & "'"
   MsgBox "Data sudah terhapus", vbInformation, "Hapus"
End If

'Perintah = "Hapus"
ClearData
ListUser
End Sub

Private Sub cmdLoadTanda_Click()
With cmDLG
    .DialogTitle = "Open Image Files"
    .FileName = ""
    .Filter = "Picture Files (*.jpg)" + Chr$(124) + "*.jpg" + Chr$(124)
    .ShowOpen
End With
linkttd.Text = cmDLG.FileName
Image1.Picture = LoadPicture(linkttd.Text)
End Sub



Private Sub fg2_Click()
    Dim Proses As String
    Proses = Perintah
    If Proses = "" Then
        If fg2.Row <> 0 Then
            ClearData
            Text1.Text = fg2.TextMatrix(fg2.Row, 0)
            Text2.Text = fg2.TextMatrix(fg2.Row, 1)
            ComboBoxLB1.Text = fg2.TextMatrix(fg2.Row, 2)
            cboGroupAkses.Text = fg2.TextMatrix(fg2.Row, 3)
            CboKodeWil.Text = fg2.TextMatrix(fg2.Row, 4)
            Call TampilkanParaf(fg2.TextMatrix(fg2.Row, 0))
        End If
    End If

End Sub

Private Sub Form_Load()
    Frame1.BackColor = MDIProject.ACPRibbon1.BackColor
    Call FormSize(8850, 7760, Me)
    Call LoadCentreForm(Me)
    Call ControlCentreForm(Me, Frame1)

    fg1.Rows = 2
    ControlEnabled (False)
    CmdEnabled (True)
    FillComboGroup
    LoadJabatan
    ListUser
    LoadWilayah
    
End Sub

Private Sub ControlEnabled(flag As Boolean)
    Text1.Enabled = flag
    Text2.Enabled = flag
    cboGroupAkses.Enabled = flag
    ComboBoxLB1.Enabled = flag
    CboKodeWil.Enabled = flag

End Sub

Private Sub ClearData()
    Text1.Text = ""
    Text2.Text = ""
    cboGroupAkses.Value = Null
    ComboBoxLB1.Value = Null
    CboKodeWil.Value = Null
    Call HapusGrid(fg1, 1)

End Sub

Private Sub CmdEnabled(flag As Boolean)
cmdAdd.Enabled = flag
cmdEdit.Enabled = flag
cmdDelete.Enabled = flag
cmdResetPassword.Enabled = flag
cmdSave.Enabled = Not flag
cmdCancel.Enabled = Not flag
cmdExit.Enabled = flag

End Sub

Private Sub FillComboGroup()
On Error GoTo ErrorHandler
Dim rs As ADODB.Recordset
Dim sql As String

Set rs = New ADODB.Recordset
sql = "Select NamaGroup from tblUtlGroupUserHdr order by NamaGroup"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

If rs.RecordCount > 0 Then
    rs.MoveFirst
    cboGroupAkses.ColumnWidths = "5000"
    Do
        cboGroupAkses.AddItem rs!NamaGroup
        rs.MoveNext
    Loop Until rs.EOF
End If

Set rs = Nothing

ErrorHandler:
 Select Case Err.Number
 Case 3021 'Tabel dalam Keadaan kosong
   Resume Next
 Case 3705 'Tabel telah dibuka oleh perintah
    rs.Close
    Resume
End Select

End Sub

Private Sub LoadWilayah()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String
Dim i As Long
sql = "Select KodeWilayah, Wilayah from tblMstWilayah order by KodeWilayah Asc"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
If rs.RecordCount > 0 Then
    CboKodeWil.ColumnCount = 1
    CboKodeWil.ColumnWidths = "1000"
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        CboKodeWil.AddItem rs!KodeWilayah
        rs.MoveNext
    Next i
    CboKodeWil.AddItem "COF"
End If
End Sub

Private Sub LoadJabatan()
On Error GoTo ErrorHandler
Dim rs As ADODB.Recordset
Dim sql As String
Dim i As Integer

Set rs = New ADODB.Recordset
sql = "Select JabatanID,JabatanNAME from tblMstJabatan order by JabatanName"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

If rs.RecordCount > 0 Then
    ComboBoxLB1.ColumnCount = 2
    ComboBoxLB1.ColumnWidths = "2000;0"
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        ComboBoxLB1.AddItem rs!JabatanName & ";" & rs!JabatanID
        rs.MoveNext
    Next i
End If

Set rs = Nothing

ErrorHandler:
 Select Case Err.Number
 Case 3021 'Tabel dalam Keadaan kosong
   Resume Next
 Case 3705 'Tabel telah dibuka oleh perintah
    rs.Close
    Resume
End Select

End Sub
Private Sub ListUser()
Dim rs As ADODB.Recordset
Dim sql As String
Dim i As Integer

Set rs = New ADODB.Recordset
sql = "Select UserID,Nama,JabatanName,NamaGroup,KodeWilayah from vwUser order by UserID"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

Call HapusGrid(fg2, 1)
If rs.RecordCount > 0 Then
    fg2.Rows = 1
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        fg2.AddItem rs!UserID & vbTab & rs!nama & vbTab & rs!JabatanName & vbTab & rs!NamaGroup & vbTab & rs!KodeWilayah
        rs.MoveNext
    Next i
End If

End Sub
Function GetPassword()

    Dim Pwd As String, Today As String
    Today = Format(Now, "dd/MM/yyyy hh:mm:ss")
    GetPassword = UCase(Right(Today, 2) & Left(Text1, 1) & Mid(Today, 15, 2) & Right(Text1, 1) & Mid(Today, 2, 1) & Right(Today, 1))

End Function

Private Sub cmdAdd_Click()
    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Add"
    Text1.SetFocus
    ClearData
End Sub

Private Sub cmdEdit_Click()

    If Trim(Text1.Text) = "" Then
        MsgBox "Silahkan cari data di Tabel dibawah ini terlebih dahulu, lalu klik pada data yang dimaksud", vbInformation, "Edit"
        Exit Sub

    End If

    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Edit"
    Text1.Enabled = False

End Sub

Private Sub cmdSave_Click()

    On Error GoTo EH
    Dim rsSave  As ADODB.Recordset
    Dim sql     As String
    Dim PW      As String
    Dim mstream As New ADODB.Stream
        mstream.Type = adTypeBinary
    Dim cn As New ADODB.Connection
    Call cn.Open(ActiveCn)

    If Trim(linkttd.Text) <> "" Then
        mstream.Open
        mstream.LoadFromFile linkttd.Text
    End If
    
    If Trim(cboGroupAkses.Text) = "" Or Trim(Text1.Text) = "" Then
        MsgBox "Data Anda belum lengkap, silahkan cek kembali data anda", vbCritical, AT
        Exit Sub

    End If

    Set rsSave = New ADODB.Recordset
    sql = "select * from tblUtlUser where UserID= '" & UCase(Trim(Text1.Text)) & "' order by NamaGroup"
    rsSave.Open sql, cn, adOpenKeyset, adLockOptimistic

    If Perintah = "Add" Then
        If rsSave.RecordCount <> 0 Then
            MsgBox "Nomor ini sudah ada, silahkan check kembali data anda"
            Exit Sub
        End If
    End If

    cn.BeginTrans

    If rsSave.EOF = True Then
        'PW = GetPassword
        PW = Trim(Text1.Text) & "123"
        rsSave.AddNew
        rsSave!UserID = UCase(Trim(Text1.Text))
        rsSave!Password = EncryptText(PW, Text1.Text)
    End If

    rsSave!NamaGroup = cboGroupAkses.Text & ""
    rsSave!nama = Trim(Text2.Text) & ""
    rsSave!Jabatan = ComboBoxLB1.Column(1) & ""
    rsSave!KodeWilayah = CboKodeWil.Text & ""
    If Trim(linkttd.Text) <> "" Then rsSave!TandaTangan = mstream.Read

    rsSave.Update

    cn.CommitTrans

    If Perintah = "Add" Then
        MsgBox "Password : " & PW
    End If

    MsgBox "Data sudah Tersimpan", vbInformation, AT
    Perintah = ""
    ControlEnabled (False)
    CmdEnabled (True)
    ClearData
    ListUser
    Set rsSave = Nothing
    Exit Sub

EH:

    If Err.Number <> 0 Then
        cn.RollbackTrans
        MsgBox "Data tidak dapat disimpan :" & Err.Description, vbCritical, AT
        ControlEnabled (False)
        CmdEnabled (True)
        ClearData
        Perintah = ""
        Set rsSave = Nothing

    End If

End Sub

Private Sub cmdCancel_Click()
    ClearData
    CmdEnabled (True)
    ControlEnabled (False)
    Perintah = ""
    cmdAdd.SetFocus

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Resize()
    Call ControlCentreForm(Me, Frame1)
End Sub



Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case 13: SendKeys "{tab}", True
        cmdCariUsers_Click
    End Select
End Sub

Private Sub TampilkanParaf(UserID As String)
Dim rs As New ADODB.Recordset
Dim sql As String
Dim arr() As Byte

On Error Resume Next
sql = "Select TandaTangan from tblUtlUser where UserID='" & UserID & "'"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

Image1.Picture = LoadPicture()
If rs.RecordCount > 0 Then
    If Not IsNull(rs!TandaTangan) = False Then
        Exit Sub
    Else
        arr = rs!TandaTangan
        Open "c:\temp.jpg" For Binary As 1
        Put #1, , arr
        Close #1
        Image1.Picture = LoadPicture("c:\temp.jpg")
        Kill "c:\temp.jpg"
    End If
End If
End Sub

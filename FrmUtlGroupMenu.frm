VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmUtlGroupMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group User"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8325
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8130
      _cx             =   14340
      _cy             =   14684
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
         Caption         =   "Group Akses Menu "
         Height          =   8175
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8055
         Begin VB.CommandButton cmdFind 
            BackColor       =   &H00FF8080&
            Caption         =   "Find"
            Height          =   375
            Left            =   600
            TabIndex        =   18
            Top             =   7680
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   840
            Width           =   6255
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FF8080&
            Caption         =   "Add"
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   7680
            Width           =   870
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00FF8080&
            Caption         =   "Delete"
            Height          =   375
            Left            =   3270
            TabIndex        =   15
            Top             =   7680
            Width           =   915
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00FF8080&
            Caption         =   "Edit"
            Height          =   375
            Left            =   2415
            TabIndex        =   14
            Top             =   7680
            Width           =   870
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FF8080&
            Caption         =   "Exit"
            Height          =   375
            Left            =   6060
            TabIndex        =   13
            Top             =   7680
            Width           =   900
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   5205
            TabIndex        =   12
            Top             =   7680
            Width           =   870
         End
         Begin VB.CommandButton CmdSave 
            BackColor       =   &H00FF8080&
            Caption         =   "Save"
            Height          =   375
            Left            =   4170
            TabIndex        =   11
            Top             =   7680
            Width           =   1050
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   1440
            TabIndex        =   10
            Top             =   360
            Width           =   6255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Tidak Aktif"
            Height          =   255
            Left            =   1440
            TabIndex        =   2
            Top             =   1800
            Width           =   1695
         End
         Begin MSComDlg.CommonDialog cmDLG 
            Left            =   360
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmUtlGroupMenu.frx":0000
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmUtlGroupMenu.frx":0072
            TabIndex        =   4
            Top             =   840
            Width           =   975
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5295
            Left            =   120
            TabIndex        =   5
            Top             =   2160
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   9340
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Pilih Menu"
            TabPicture(0)   =   "FrmUtlGroupMenu.frx":00E4
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "VSFlexGrid1"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cari Data"
            TabPicture(1)   =   "FrmUtlGroupMenu.frx":0100
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "VSFlexGrid2"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "Command1"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Text3"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).ControlCount=   3
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2400
               TabIndex        =   7
               Text            =   " "
               Top             =   480
               Width           =   3615
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Find"
               Height          =   315
               Left            =   6120
               TabIndex        =   6
               Top             =   480
               Width           =   1095
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
               Height          =   4335
               Left            =   120
               TabIndex        =   8
               Top             =   840
               Width           =   7095
               _cx             =   12515
               _cy             =   7646
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
               BackColorFixed  =   14669471
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12251379
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmUtlGroupMenu.frx":011C
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
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   4815
               Left            =   -74880
               TabIndex        =   9
               Top             =   360
               Width           =   7455
               _cx             =   13150
               _cy             =   8493
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
               BackColorFixed  =   14669471
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   12251379
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmUtlGroupMenu.frx":01A6
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
               OutlineBar      =   4
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   1
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
   End
End
Attribute VB_Name = "FrmUtlGroupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Perintah As String

Private Sub cmdDelete_Click()

    Dim cn As New ADODB.Connection

    Call cn.Open(ActiveCn)

    If Trim(Text1.Text) = "" Then
        MsgBox "Tidak ada Data yang mau dihapus , silahkan cari data dengan tombol FIND", vbInformation, AT
        Exit Sub

    End If

    If MsgBox("Anda yakin mau menghapus SELURUH data ini", vbYesNo) = vbYes Then
        cn.Execute "delete from tblUtlGroupUserDtl where NamaGroup = '" & UCase(Trim(Text1.Text)) & "'"
        cn.Execute "delete from tblUtlGroupUserHdr where NamaGroup = '" & UCase(Trim(Text1.Text)) & "'"
        MsgBox "Data sudah terhapus", vbInformation, AT

    End If

    ClearData
    LoadGroupUser

End Sub

Private Sub Form_Load()
    'Call FormSize(8850, 7760, Me)
    Call LoadCentreForm(Me)
    VSFlexGrid1.Rows = 2
    ControlEnabled (False)
    CmdEnabled (True)
    Me.Left = (Screen.Width - Me.Width) / 2
    LoadGroupUser
    SetMenuList

End Sub

Private Sub ControlEnabled(flag As Boolean)
    Text1.Enabled = flag
    Text2.Enabled = flag
    Check1.Enabled = flag
    VSFlexGrid1.Enabled = flag

End Sub

Private Sub ClearData()
    Text1.Text = ""
    Text2.Text = ""
    Check1.Value = False
    Call HapusGrid(VSFlexGrid1, 1)
End Sub

Private Sub CmdEnabled(flag As Boolean)
    cmdAdd.Enabled = flag
    CmdEdit.Enabled = flag
    cmdDelete.Enabled = flag
    cmdFind.Enabled = flag
    cmdSave.Enabled = Not flag
    cmdCancel.Enabled = Not flag
    cmdExit.Enabled = flag

End Sub

Private Sub cmdAdd_Click()
    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Add"
    Text1.SetFocus
    ClearData
    SetMenuList
End Sub

Private Sub SetMenuList()
    Dim LevelMenus  As Byte
    Dim CurrentMenu As Menu
    Dim intx        As Long, intY As Long
    Call HapusGrid(VSFlexGrid1, 1)

    For intx = 0 To MDIProject.Controls.Count - 1

        If TypeOf MDIProject.Controls(intx) Is Menu Then
            If (MDIProject.Controls(intx).Index <> 0) Then

                With VSFlexGrid1
                    .IsSubtotal(.Rows - 1) = True
                    LevelMenus = LevelMenu(MDIProject.Controls(intx).Index) 'Len(Trim(CStr(mdigl.Controls(intX).Index)))
                    .RowOutlineLevel(.Rows - 1) = LevelMenus
                    .TextMatrix(.Rows - 1, 0) = "  " & MDIProject.Controls(intx).Caption
                    .TextMatrix(.Rows - 1, 1) = Trim(CStr(MDIProject.Controls(intx).Index))
                    .Rows = .Rows + 1
                End With
            End If
        End If
    Next intx

    VSFlexGrid1.Rows = VSFlexGrid1.Rows - 1

    'Check if has child
    With VSFlexGrid1

        For intx = 1 To .Rows - 1
            .TextMatrix(intx, 2) = "N"
            .Cell(flexcpChecked, intx, 0) = flexUnchecked
            .Cell(flexcpChecked, intx, 3, intx, .Cols - 1) = flexNoCheckbox
            .Cell(flexcpPictureAlignment, intx, 3, intx, .Cols - 1) = flexPicAlignCenterCenter
            .Cell(flexcpBackColor, intx, 3, intx, .Cols - 1) = RGB(240, 240, 240)
            .Cell(flexcpForeColor, intx, 0) = RGB(180, 180, 180)
            .RowHeight(intx) = 240

            For intY = intx + 1 To .Rows - 1

                If .RowOutlineLevel(intY) > .RowOutlineLevel(intx) Then
                    .TextMatrix(intx, 2) = "Y"
                    .TextMatrix(intx, 0) = Trim(.TextMatrix(intx, 0))
                    .Cell(flexcpFontBold, intx, 0) = True

                    If .RowOutlineLevel(intx) = 1 Then
                        .Cell(flexcpFontSize, intx, 0) = 11
                        .RowHeight(intx) = 300
                    End If
                    .Cell(flexcpChecked, intx, 0) = flexNoCheckbox
                    '.Cell(flexcpChecked, intX, 3, intX, .Cols - 1) = flexNoCheckbox
                    .Cell(flexcpBackColor, intx, 3, intx, .Cols - 1) = vbWhite
                    Exit For
                End If
            Next intY
        Next intx

        .TextMatrix(.Rows - 1, 2) = "N"
        .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
        .Cell(flexcpChecked, .Rows - 1, 3, .Rows - 1, .Cols - 1) = flexNoCheckbox
        .Cell(flexcpPictureAlignment, .Rows - 1, 3, .Rows - 1, .Cols - 1) = flexPicAlignCenterCenter
        .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = RGB(240, 240, 240)
        .Cell(flexcpForeColor, .Rows - 1, 0) = RGB(180, 180, 180)
    End With
    VSFlexGrid1.Editable = flexEDKbdMouse
End Sub

Private Sub LoadGroupUser()
Dim rs As ADODB.Recordset
Dim sql As String
Dim I As Integer

Set rs = New ADODB.Recordset
sql = "select * from TblUtlGroupUserHdr"
rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

If rs.RecordCount > 0 Then
    VSFlexGrid2.Rows = 1
    rs.MoveFirst
    For I = 0 To rs.RecordCount - 1
        VSFlexGrid2.AddItem rs!NamaGroup & vbTab & rs!keterangan & vbTab & rs!NotAktif
        rs.MoveNext
    Next I
End If
Set rs = Nothing

End Sub

Private Function LevelMenu(indek As Long) As Byte

    If indek Mod 1000 = 0 Then
        LevelMenu = 1
    ElseIf indek Mod 100 = 0 Then
        LevelMenu = 2
    Else
        LevelMenu = 3
    End If

End Function


Private Sub cmdCancel_Click()
    ClearData
    CmdEnabled (True)
    ControlEnabled (False)
    Perintah = ""
    cmdAdd.SetFocus

End Sub

Private Sub cmdEdit_Click()
    If Trim(Text1.Text) = "" Then
        MsgBox "Anda harus Cari data menggunakan tombol FIND, lalu klik 2(dua)kali pada item yang dimaksud"
        Exit Sub
    End If
    ControlEnabled (True)
    CmdEnabled (False)
    Perintah = "Edit"
    Text1.Enabled = False
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorHandler
    Dim cn     As New ADODB.Connection
    Dim rsSave As ADODB.Recordset
    Dim sql    As String
    Call cn.Open(ActiveCn)

    If Trim(Text1.Text) = "" Then
        MsgBox "Data Anda belum lengkap, silahkan cek kembali data anda", vbCritical, AT
        Exit Sub
    End If

    Set rsSave = New ADODB.Recordset
    sql = "select * from tblUtlGroupUserHdr where NamaGroup= '" & UCase(Trim(Text1.Text)) & "' order by NamaGroup"
    rsSave.Open sql, cn, adOpenKeyset, adLockOptimistic

    cn.BeginTrans

    If Perintah = "Add" Then
        If rsSave.RecordCount <> 0 Then
            MsgBox "Nomor ini sudah ada, silahkan check kembali data anda"
            Exit Sub
        End If
    End If

    If rsSave.EOF = True Then
        rsSave.AddNew
        rsSave!NamaGroup = UCase(Trim(Text1.Text))
    End If

    rsSave!keterangan = Text2.Text
    rsSave!NotAktif = Check1.Value

    rsSave.Update

    Call SimpanDetail(cn)

    Perintah = ""
    ControlEnabled (False)
    CmdEnabled (True)
    ClearData

    cn.CommitTrans
    MsgBox "Data sudah Tersimpan", vbInformation, AT

    ClearData
    LoadGroupUser
    ControlEnabled (False)
    CmdEnabled (True)
    
    Set rsSave = Nothing
    Exit Sub

ErrorHandler:

    If Err.Number <> 0 Then
        cn.RollbackTrans
        Set rsSave = Nothing
        MsgBox "Data tidak dapat disimpan, Err : " & Err.Description
        Exit Sub

    End If

End Sub

Private Sub SimpanDetail(con As ADODB.Connection)

    Dim I As Long

    con.Execute "Delete from tblUtlGroupUserDtl where NamaGroup = '" & UCase(Trim(Text1.Text)) & "'"

    For I = 1 To VSFlexGrid1.Rows - 1

        If VSFlexGrid1.TextMatrix(I, 1) <> "" Then
            If VSFlexGrid1.ValueMatrix(I, 6) = True Then
                con.Execute "Insert Into tblUtlGroupUserDtl " & "(NamaGroup,IndekMenu) " & "Values " & "('" & UCase(Trim(Text1.Text)) & "','" & VSFlexGrid1.TextMatrix(I, 1) & "')"

            End If

        End If

    Next I

End Sub



Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If VSFlexGrid1.GetNodeRow(Row, flexNTLastChild) > -1 Then Cancel = True
End Sub


Private Sub VSFlexGrid1_Click()
Dim I As Long
Dim AdaLevel4, AdaLevel2 As Boolean
If VSFlexGrid1.Cell(flexcpChecked, VSFlexGrid1.Row, 0) = 1 Then
    VSFlexGrid1.Cell(flexcpForeColor, VSFlexGrid1.Row, 0) = vbBlue
    VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 6) = True
    For I = 1 To VSFlexGrid1.Rows - 1
        If VSFlexGrid1.ValueMatrix(I, 1) = VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) - (VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) Mod 1000) Then
            VSFlexGrid1.Cell(flexcpForeColor, I, 0) = vbBlue
            VSFlexGrid1.TextMatrix(I, 6) = True
        End If
        If VSFlexGrid1.ValueMatrix(I, 1) = VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) - (VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) Mod 100) Then
            VSFlexGrid1.Cell(flexcpForeColor, I, 0) = vbBlue
            VSFlexGrid1.TextMatrix(I, 6) = True
        End If
    Next I
Else
    AdaLevel2 = False
    AdaLevel4 = False
    VSFlexGrid1.Cell(flexcpForeColor, VSFlexGrid1.Row, 0) = RGB(180, 180, 180)
    VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 6) = False
    For I = 1 To VSFlexGrid1.Rows - 1
        If (VSFlexGrid1.ValueMatrix(I, 1) - (VSFlexGrid1.ValueMatrix(I, 1) Mod 100)) = VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) - (VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) Mod 100) Then
            If VSFlexGrid1.ValueMatrix(I, 6) = True And (VSFlexGrid1.ValueMatrix(I, 1) Mod 100) > 0 Then
                AdaLevel4 = True
            End If
        End If
        If (VSFlexGrid1.ValueMatrix(I, 1) - (VSFlexGrid1.ValueMatrix(I, 1) Mod 1000)) = VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) - (VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) Mod 1000) Then
            If VSFlexGrid1.ValueMatrix(I, 6) = True And (VSFlexGrid1.ValueMatrix(I, 1) Mod 100) > 0 Then
                AdaLevel2 = True
            End If
        End If
    Next I
    
    For I = 1 To VSFlexGrid1.Rows - 1
        If VSFlexGrid1.ValueMatrix(I, 1) = VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) - (VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) Mod 1000) Then
            If AdaLevel2 = False Then
                VSFlexGrid1.Cell(flexcpForeColor, I, 0) = RGB(180, 180, 180)
                VSFlexGrid1.TextMatrix(I, 6) = False
            End If
        End If
        If VSFlexGrid1.ValueMatrix(I, 1) = VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) - (VSFlexGrid1.ValueMatrix(VSFlexGrid1.Row, 1) Mod 100) Then
            If AdaLevel4 = False Then
                VSFlexGrid1.Cell(flexcpForeColor, I, 0) = RGB(180, 180, 180)
                VSFlexGrid1.TextMatrix(I, 6) = False
            End If
        End If
    Next I
End If
End Sub

Private Sub VSFlexGrid2_Click()

    If Perintah = "" Then
        If VSFlexGrid2.Row <> 0 Then
            Text1.Text = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 0)
            Text2.Text = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1)

            If VSFlexGrid2.ValueMatrix(VSFlexGrid2.Row, 2) = True Then
                Check1.Value = 1
            Else
                Check1.Value = 0

            End If
            
            SetMenuList
            ListGroupUser

        End If

    End If

End Sub

Private Sub ListGroupUser()

    Dim rsDetail  As New ADODB.Recordset

    Dim sqlDetail As String

    Dim I, j As Long

    sqlDetail = "select * from tblUtlGroupUserDtl where NamaGroup='" & VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 0) & "' order by IndekMenu "
    rsDetail.Open sqlDetail, ActiveCn, adOpenKeyset, adLockReadOnly

    If rsDetail.RecordCount > 0 Then
        rsDetail.MoveFirst

        For j = 1 To VSFlexGrid1.Rows - 1
            VSFlexGrid1.Cell(flexcpChecked, j, 0) = flexUnchecked
        Next j
        
        Do While Not rsDetail.EOF
            For j = 1 To VSFlexGrid1.Rows - 1
                If VSFlexGrid1.TextMatrix(j, 1) = rsDetail!IndekMenu Then
                    VSFlexGrid1.Cell(flexcpForeColor, j, 0) = vbBlue
                    If VSFlexGrid1.TextMatrix(j, 2) = "N" Then
                        VSFlexGrid1.Cell(flexcpChecked, j, 0) = flexChecked
                    End If
                    VSFlexGrid1.TextMatrix(j, 6) = True
                    Exit For
                End If
            Next j
            rsDetail.MoveNext
        Loop
   Else
        For j = 1 To VSFlexGrid1.Rows - 1
            VSFlexGrid1.Cell(flexcpChecked, j, 0) = flexUnchecked
            VSFlexGrid1.TextMatrix(j, 6) = False
        Next j
    End If

    Set rsDetail = Nothing

End Sub

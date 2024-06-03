VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptLASPMingguan 
   Caption         =   "Laporan LASP Mingguan"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14430
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
   ScaleHeight     =   8865
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8865
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14430
      _cx             =   25453
      _cy             =   15637
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
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
      GridRows        =   3
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmRptLASPMingguan.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   90
         TabIndex        =   11
         Top             =   8250
         Width           =   11520
         Begin VB.CommandButton CmdApprove 
            BackColor       =   &H008080FF&
            Caption         =   "Approve"
            Height          =   255
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   150
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080FFFF&
            Caption         =   "Approve"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton CmdNotApprove 
            BackColor       =   &H008080FF&
            Caption         =   "Approve Cancel "
            Height          =   255
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   150
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Laporan Level AIr Sumur Pantau Mingguan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   14250
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2280
            TabIndex        =   3
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   103481347
            CurrentDate     =   43143
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   4200
            TabIndex        =   4
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   103481347
            CurrentDate     =   43143
         End
         Begin Combo.ComboBoxLB cboWilayah 
            Height          =   315
            Left            =   2280
            TabIndex        =   5
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin VB.Label lblPeriode 
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
            Left            =   720
            TabIndex        =   8
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblSD 
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
            Left            =   3720
            TabIndex        =   7
            Top             =   600
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
            Left            =   720
            TabIndex        =   6
            Top             =   960
            Width           =   735
         End
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer2 
         Height          =   6450
         Left            =   90
         TabIndex        =   1
         Top             =   1740
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   11377
         SectionData     =   "FrmRptLASPMingguan.frx":008F
      End
      Begin MyLASP.isButton CmdClose 
         Height          =   525
         Left            =   13035
         TabIndex        =   9
         Top             =   8250
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   926
         Icon            =   "FrmRptLASPMingguan.frx":00CB
         Style           =   5
         Caption         =   "Clos&e"
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
      Begin MyLASP.isButton cmdRefresh 
         Height          =   525
         Left            =   11670
         TabIndex        =   10
         Top             =   8250
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   926
         Icon            =   "FrmRptLASPMingguan.frx":0C5F
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
   End
End
Attribute VB_Name = "FrmRptLASPMingguan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FormKontrol As String

Private Sub CmdApprove_Click()
Dim cn As New ADODB.Connection
Dim sql As String

   Call cn.Open(ActiveCn)
    
    If cboWilayah.Value <> "" Then
    
        If Check1.Value = 1 And Check1.Caption = "Approve" And MDIProject.GroupUser = "ADM WILAYAH" Then
        
            cn.BeginTrans
            cn.Execute "update tblTrnRekapitulasiDtl set ApprovalBy='" & MDIProject.UserID & "',ApprovalDate='" & GetDate() & "' " & _
                        "where HeaderID in (Select HeaderID from tblTrnRekapitulasiHdr where " & _
                        "Bulan between '" & Format(DTPicker1.Value, "mm") & "' and '" & Format(DTPicker2.Value, "mm") & "' " & _
                        "and Tahun='" & Format(DTPicker2.Value, "yyyy") & "' and KodeWilayah='" & cboWilayah.Column(1) & "') " & _
                        ""
            cn.CommitTrans
            
        ElseIf Check1.Value = 1 And Check1.Caption = "Approve" And MDIProject.GroupUser = "ADM COF" Then
        
            cn.BeginTrans
            cn.Execute "update tblTrnRekapitulasiDtl set ApprovalBy='" & MDIProject.UserID & "', ApprovalDate='" & Format(GetDate(), "yyyy-mm-dd") & "' " & _
                        "where HeaderID in (Select HeaderID from tblTrnRekapitulasiHdr where " & _
                        "Bulan between '" & Format(DTPicker1.Value, "mm") & "' and '" & Format(DTPicker2.Value, "mm") & "' " & _
                        "and Tahun='" & Format(DTPicker2.Value, "yyyy") & "' and KodeWilayah='" & cboWilayah.Column(1) & "') " & _
                        ""
            cn.CommitTrans
            
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Format(GetDate, "yyyy-mm-dd")
    DTPicker2.Value = Format(GetDate, "yyyy-mm-dd")
    'Call LoadMonRekapitulasi
    If FormKontrol = "ALL" Then
        Frame1.Caption = "Laporan Level Air Sumur Pantau Mingguan"
        lblWilayah.Visible = False
        cboWilayah.Visible = False
    Else
        Frame1.Caption = "Laporan Level Air Sumur Pantau Mingguan Per Wilayah"
        Call LoadWilayah(cboWilayah)
    End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    If Format(DTPicker1.Value, "yyyy") <> Format(DTPicker2.Value, "yyyy") Then
        MsgBox "Tahun periode Pertama dan Kedua Berbeda.Silahkan Cek Kembali!..", vbInformation, AT
        Exit Sub
        DTPicker2.SetFocus
    End If

    If FormKontrol = "ALL" Then

        Me.MousePointer = vbHourglass
            Set ARViewer2.ReportSource = New rptLASPMingguan
        Me.MousePointer = vbDefault
        Exit Sub

    Else
        If cboWilayah.Text <> "" Then
        
            Set rs = New ADODB.Recordset
            
            sql = "select ApprovalBy,ApprovalDate from tblTrnRekapitulasiDtl " & _
                  "where HeaderID in (Select HeaderID from tblTrnRekapitulasiHdr where " & _
                  "Bulan between '" & Format(DTPicker1.Value, "mm") & "' and '" & Format(DTPicker2.Value, "mm") & "' " & _
                  "and Tahun='" & Format(DTPicker2.Value, "yyyy") & "' and KodeWilayah='" & cboWilayah.Column(1) & "') " & _
                  "group by ApprovalBy,ApprovalDate"
                rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
                
                If rs.RecordCount > 0 Then
                        Check1.Visible = True
                        CmdApprove.Visible = True
                    
                    If IsNull(rs!ApprovalBy) = True Then
                        Check1.Caption = "Approve"
                    Else
                        Check1.Visible = False
                        CmdApprove.Visible = False
                    End If
                Else
                        Check1.Caption = "Approve"
                End If
        
            Me.MousePointer = vbHourglass
                Set ARViewer2.ReportSource = New rptLaspMingguanPerWil
            Me.MousePointer = vbDefault
            Exit Sub
            
        Else
            MsgBox "Kode Wilayah tidak boleh kosong!", vbInformation, "Perhatian"
        End If
    End If
    
End Sub


VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptGrafikPerWil 
   Caption         =   "Laporan Grafik LASP PerWil"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7875
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13665
      _cx             =   24104
      _cy             =   13891
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
      ChildSpacing    =   5
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
      _GridInfo       =   $"FrmRptGrafikPerWil.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Laporan Grafik LASP PerWilayah"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   13485
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2280
            TabIndex        =   2
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   97976323
            CurrentDate     =   43143
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   4200
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   97976323
            CurrentDate     =   43143
         End
         Begin Combo.ComboBoxLB cboWilayah 
            Height          =   315
            Left            =   2280
            TabIndex        =   4
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin Combo.ComboBoxLB cboKodeTitik 
            Height          =   315
            Left            =   2280
            TabIndex        =   9
            Top             =   1200
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            Appearance      =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Titik"
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
            Top             =   1200
            Width           =   900
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
            TabIndex        =   7
            Top             =   840
            Width           =   735
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
            TabIndex        =   6
            Top             =   480
            Width           =   285
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
            TabIndex        =   5
            Top             =   480
            Width           =   720
         End
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer2 
         Height          =   5400
         Left            =   90
         TabIndex        =   8
         Top             =   1815
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   9525
         SectionData     =   "FrmRptGrafikPerWil.frx":008F
      End
      Begin MyLASP.isButton CmdClose 
         Height          =   495
         Left            =   12255
         TabIndex        =   11
         Top             =   7290
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   873
         Icon            =   "FrmRptGrafikPerWil.frx":00CB
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
         Height          =   495
         Left            =   10875
         TabIndex        =   12
         Top             =   7290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Icon            =   "FrmRptGrafikPerWil.frx":0C5F
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
Attribute VB_Name = "FrmRptGrafikPerWil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FormKontrol  As String

Private Sub Form_Load()
    DTPicker1.Value = Format(GetDate, "yyyy-mm-dd")
    DTPicker2.Value = Format(GetDate, "yyyy-mm-dd")
    If FormKontrol = "ALL" Then
        Me.Caption = "Laporan Grafik LASP"
        Frame1.Caption = "Laporan Grafik LASP"
        cboWilayah.Visible = False
        cboKodeTitik.Visible = False
        lblWilayah.Visible = False
        Label1.Visible = False
    Else
        Me.Caption = "Laporan Grafik LASP PerWil"
        Call LoadWilayah(cboWilayah)
    End If
End Sub

Private Sub LoadKodeTitik(KodeWilayah As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim i As Long
    sql = "select KodeWilayah,KodeTitik  from vwMstPersil where KodeWilayah='" & KodeWilayah & "' order by KodeTitik asc "
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then
        cboKodeTitik.ColumnCount = 2
        cboKodeTitik.ColumnWidths = "2500;500"
        cboKodeTitik.Clear
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            cboKodeTitik.AddItem rs!KodeTitik & ";" & rs!KodeWilayah
            rs.MoveNext
        Next i
    End If
End Sub

Private Sub cboWilayah_AfterUpdate()
    Call LoadKodeTitik(cboWilayah.Column(1))
End Sub

Private Sub cmdRefresh_Click()
    If Format(DTPicker1.Value, "yyyy") <> Format(DTPicker2.Value, "yyyy") Then
        MsgBox "Tahun periode Pertama dan Kedua Berbeda.Silahkan Cek Kembali!..", vbInformation, AT
        Exit Sub
        DTPicker2.SetFocus
    End If
    
    
    If FormKontrol = "ALL" Then
        Me.MousePointer = vbHourglass
           Set ARViewer2.ReportSource = New rptLASPGrafik
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        If cboWilayah.Text = "" Then
            MsgBox "Kode Wilayah tidak boleh kosong!", vbInformation, "Perhatian"
            Exit Sub
        End If
    
        If cboKodeTitik.Text = "" Then
            MsgBox "Kode Titik tidak boleh kosong!", vbInformation, AT
            Exit Sub
        End If
    
        Me.MousePointer = vbHourglass
           Set ARViewer2.ReportSource = New rptLASPGrafikWil
        Me.MousePointer = vbDefault
        Exit Sub
    End If

End Sub


Private Sub CmdClose_Click()
    Unload Me
End Sub

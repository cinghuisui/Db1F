VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptLAKMingguan 
   Caption         =   "Laporan LAK Mingguan"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
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
   ScaleHeight     =   7755
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7755
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14220
      _cx             =   25083
      _cy             =   13679
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
      _GridInfo       =   $"FrmRptLAKMingguan.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Height          =   1290
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   14040
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2280
            TabIndex        =   2
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   335806467
            CurrentDate     =   43143
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   4200
            TabIndex        =   3
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   335806467
            CurrentDate     =   43143
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
            TabIndex        =   5
            Top             =   360
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
            TabIndex        =   4
            Top             =   360
            Width           =   720
         End
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer2 
         Height          =   5670
         Left            =   90
         TabIndex        =   6
         Top             =   1440
         Width           =   14040
         _ExtentX        =   24765
         _ExtentY        =   10001
         SectionData     =   "FrmRptLAKMingguan.frx":008D
      End
      Begin MyLASP.isButton CmdClose 
         Height          =   495
         Left            =   12825
         TabIndex        =   7
         Top             =   7170
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Icon            =   "FrmRptLAKMingguan.frx":00C9
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
         Left            =   11445
         TabIndex        =   8
         Top             =   7170
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   873
         Icon            =   "FrmRptLAKMingguan.frx":0C5D
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
Attribute VB_Name = "FrmRptLAKMingguan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FormKontrol As String

Private Sub Form_Load()
    DTPicker1.Value = Format(GetDate, "yyyy-mm-dd")
    DTPicker2.Value = Format(GetDate, "yyyy-mm-dd")
    'Call LoadMonRekapitulasi
    If FormKontrol = "ALL" Then
        Me.Caption = "Laporan LAK Bulanan"
    Else
        Me.Caption = "Laporan LAK Mingguan"
    End If
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    If Format(DTPicker1.Value, "yyyy") <> Format(DTPicker2.Value, "yyyy") Then
        MsgBox "Tahun periode Pertama dan Kedua Berbeda.Silahkan Cek Kembali!..", vbInformation, AT
        Exit Sub
        DTPicker2.SetFocus
    End If
    
    If FormKontrol = "ALL" Then
        Me.MousePointer = vbHourglass
           Set ARViewer2.ReportSource = New rptLAKBulanan
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
           Set ARViewer2.ReportSource = New rptLAKMingguan
        Me.MousePointer = vbDefault
        Exit Sub
    End If

End Sub



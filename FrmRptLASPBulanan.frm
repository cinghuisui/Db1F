VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRptLASPBulanan 
   Caption         =   "Laporan LASP Bulanan"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14955
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
   ScaleHeight     =   7650
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7650
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14955
      _cx             =   26379
      _cy             =   13494
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
      _GridInfo       =   $"FrmRptLASPBulanan.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Laporan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   14775
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
            Format          =   335806467
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
            Format          =   335806467
            CurrentDate     =   43143
         End
         Begin Combo.ComboBoxLB cboWilayah 
            Height          =   315
            Left            =   2280
            TabIndex        =   7
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            Appearance      =   0
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
            Left            =   1080
            TabIndex        =   8
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
            TabIndex        =   5
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
            Left            =   1080
            TabIndex        =   4
            Top             =   480
            Width           =   720
         End
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer2 
         Height          =   5325
         Left            =   90
         TabIndex        =   6
         Top             =   1650
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9393
         SectionData     =   "FrmRptLASPBulanan.frx":008F
      End
      Begin MyLASP.isButton CmdClose 
         Height          =   525
         Left            =   13560
         TabIndex        =   9
         Top             =   7035
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   926
         Icon            =   "FrmRptLASPBulanan.frx":00CB
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
         Left            =   12195
         TabIndex        =   10
         Top             =   7035
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   926
         Icon            =   "FrmRptLASPBulanan.frx":0C5F
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
Attribute VB_Name = "FrmRptLASPBulanan"
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
        Frame1.Caption = "Laporan Level Air Sumur Pantau Bulanan"
        lblWilayah.Visible = False
        cboWilayah.Visible = False
    Else
        Frame1.Caption = "Laporan Level Air Sumur Pantau Bulanan PerWilayah"
        Call LoadWilayah(cboWilayah)
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
            Set ARViewer2.ReportSource = New rptLASPBulanan
        Me.MousePointer = vbDefault
        Exit Sub

    Else
        If cboWilayah.Text <> "" Then
        
            Me.MousePointer = vbHourglass
                Set ARViewer2.ReportSource = New rptLASPBulananPerWil
            Me.MousePointer = vbDefault
            Exit Sub
            
        Else
            MsgBox "Kode Wilayah tidak boleh kosong!", vbInformation, "Perhatian"
        End If
    End If
End Sub


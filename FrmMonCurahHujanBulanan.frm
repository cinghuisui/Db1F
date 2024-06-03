VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMonCurahHujanBulanan 
   Caption         =   "Monitoring CRH Bulanan"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14655
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
   ScaleHeight     =   8745
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14655
      _cx             =   25850
      _cy             =   15425
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
      BackColor       =   -2147483646
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
      _GridInfo       =   $"FrmMonCurahHujanBulanan.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Monitoring Curah Hujan Bulanan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1185
         Left            =   210
         TabIndex        =   1
         Top             =   90
         Width           =   14145
         Begin MSComDlg.CommonDialog cmDLG 
            Left            =   6720
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2280
            TabIndex        =   2
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   98959363
            CurrentDate     =   43143
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   4200
            TabIndex        =   3
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "MMyyyy"
            Format          =   98959363
            CurrentDate     =   43143
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
            TabIndex        =   4
            Top             =   600
            Width           =   285
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   6765
         Left            =   210
         TabIndex        =   6
         Top             =   1335
         Width           =   14145
         _cx             =   24950
         _cy             =   11933
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16762250
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
         Rows            =   30
         Cols            =   16
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmMonCurahHujanBulanan.frx":0089
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
      Begin MyLASP.isButton CmdClose 
         Height          =   495
         Left            =   13170
         TabIndex        =   7
         Top             =   8160
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         Icon            =   "FrmMonCurahHujanBulanan.frx":02B0
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
         Left            =   11805
         TabIndex        =   8
         Top             =   8160
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         Icon            =   "FrmMonCurahHujanBulanan.frx":0E44
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
      Begin MyLASP.isButton cmdExcel 
         Height          =   495
         Left            =   10560
         TabIndex        =   9
         Top             =   8160
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         Icon            =   "FrmMonCurahHujanBulanan.frx":1964
         Style           =   5
         Caption         =   "&Excel"
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
Attribute VB_Name = "FrmMonCurahHujanBulanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    fg.WordWrap = True
    fg.MergeCells = flexMergeFixedOnly
    fg.MergeRow(0) = True
    fg.MergeRow(1) = True
    
    
    For i = 0 To fg.Cols - 1
        fg.MergeCol(i) = True
        fg.FixedAlignment(i) = flexAlignCenterCenter
    Next i
    
    
    DTPicker1.Value = Format(GetDate, "yyyy-mm-dd")
    DTPicker2.Value = Format(GetDate, "yyyy-mm-dd")
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
    Call ConvertToExcel(cmDLG, fg, Me)
End Sub

Private Sub cmdRefresh_Click()

    Dim j As Integer
    
    If Format(DTPicker1.Value, "yyyy") <> Format(DTPicker2.Value, "yyyy") Then
        MsgBox "Tahun periode Pertama dan Kedua Berbeda.Silahkan Cek Kembali!..", vbInformation, AT
        Exit Sub
        DTPicker1.SetFocus
    End If
    
    Call LoadMonRekapitulasiCRHBulanan
    
    fg.Subtotal flexSTClear
    fg.SubtotalPosition = flexSTBelow

    For j = 3 To 14
        fg.Subtotal flexSTAverage, -1, j, "#,###", , , , "Rata-rata CRH PT.RSUP-PKB"
    Next j
    
End Sub

Private Sub LoadMonRekapitulasiCRHBulanan()
    Dim rs  As ADODB.Recordset
    Dim sql As String
    Dim i, j As Integer
    Dim DataSebelum As String
    Set rs = New ADODB.Recordset
    sql = "Select KodeWilayah,Wilayah,Bulan,Tahun,Satuan,Round(AVG(Nilai),0) as Nilai from vwRekapitulasiCurahHujan " & _
          "where Bulan between '" & Format(DTPicker1.Value, "MM") & "' and '" & Format(DTPicker2.Value, "MM") & "' and tahun='" & Format(DTPicker2.Value, "yyyy") & "' " & _
          "group by Bulan,Tahun,KodeWilayah,Satuan,Wilayah Order by KodeWilayah Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly
    Call HapusGrid(fg, 2)
    If rs.RecordCount > 0 Then
        With fg
        .Rows = 2
            Do Until rs.EOF
                If DataSebelum <> rs!KodeWilayah Then
                    .AddItem ""
                End If
                    .TextMatrix(.Rows - 1, 0) = .Rows - 2
                    .TextMatrix(.Rows - 1, 1) = rs!Wilayah & ""
                    .TextMatrix(.Rows - 1, 2) = "mm" & ""
                    
                    If IsNull(rs!Bulan) = False Then
                        Select Case rs!Bulan
                            Case 1
                                .TextMatrix(.Rows - 1, 3) = rs!Nilai & ""
                            Case 2
                                .TextMatrix(.Rows - 1, 4) = rs!Nilai & ""
                            Case 3
                                .TextMatrix(.Rows - 1, 5) = rs!Nilai & ""
                            Case 4
                                .TextMatrix(.Rows - 1, 6) = rs!Nilai & ""
                            Case 5
                                .TextMatrix(.Rows - 1, 7) = rs!Nilai & ""
                            Case 6
                                .TextMatrix(.Rows - 1, 8) = rs!Nilai & ""
                            Case 7
                                .TextMatrix(.Rows - 1, 9) = rs!Nilai & ""
                            Case 8
                                .TextMatrix(.Rows - 1, 10) = rs!Nilai & ""
                            Case 9
                                .TextMatrix(.Rows - 1, 11) = rs!Nilai & ""
                            Case 10
                                .TextMatrix(.Rows - 1, 12) = rs!Nilai & ""
                            Case 11
                                .TextMatrix(.Rows - 1, 13) = rs!Nilai & ""
                            Case 12
                                .TextMatrix(.Rows - 1, 14) = rs!Nilai & ""
                        End Select
                    End If
                     .TextMatrix(.Rows - 1, 15) = GetValue("Select Round(AVG(Nilai),0) as GetValue " & _
                                                    " From fnRekapitulasiCRHBulanan('" & rs!KodeWilayah & "','" & Format(DTPicker1.Value, "MM") & "','" & Format(DTPicker2.Value, "MM") & "','" & Format(DTPicker2.Value, "yyyy") & "') ")
                
                DataSebelum = rs!KodeWilayah
                rs.MoveNext
            Loop
        End With
            
    End If

End Sub

Private Function ComputeTotalAmount(baris)

    Dim Amount As Double, i As Long

    With fg

        For i = 3 To 14 - 1 ' Hitung Level Air Kolom 3 sampai kolom 14 Januari s/d Desember
            If fg.TextMatrix(baris, i) <> "" Then
                If .TextMatrix(baris, i) <> "" Then Amount = Amount + CDbl(.TextMatrix(baris, i))
            End If
        Next i

        ComputeTotalAmount = Format(Amount, "#,###")

    End With

End Function

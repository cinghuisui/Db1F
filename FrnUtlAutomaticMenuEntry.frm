VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrnUtlAutomaticMenuEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Menu Entry"
   ClientHeight    =   8475
   ClientLeft      =   6255
   ClientTop       =   2220
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   7380
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7380
      _cx             =   13018
      _cy             =   14949
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
      AutoSizeChildren=   7
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
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   7560
         Left            =   120
         TabIndex        =   2
         Top             =   735
         Width           =   6960
         _cx             =   12277
         _cy             =   13335
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrnUtlAutomaticMenuEntry.frx":0000
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
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update Menu"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6960
      End
   End
End
Attribute VB_Name = "FrnUtlAutomaticMenuEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdUpdate_Click()
    Dim indek As Integer
    Dim I, intx As Long
    Dim cn As New ADODB.Connection
    Call cn.Open(ActiveCn)

    cn.Execute "delete from tblUtlMenu"

    For intx = 0 To MDIProject.Controls.Count - 1

        If TypeOf MDIProject.Controls(intx) Is Menu Then
            If (MDIProject.Controls(intx).Index <> 0) Then 'Index 0 yaitu form utama

                Select Case MDIProject.Controls(intx).Index

                    Case 1000, 2000, 3000, 4000, 5000, 6000, 6100, 6200, 6300, 7000

                        'Biarkan kosong, ini untuk mengakali supaya tidak masuk dalam
                        'database yaitu menu2 yang boleh di akses oleh seluruh User
                    Case Else
                        cn.Execute "Insert Into tblUtlMenu (IndekMenu,NamaMenu,JudulMenu) values " & "(" & MDIProject.Controls(intx).Index & ",'" & MDIProject.Controls(intx).Name & "','" & Replace(MDIProject.Controls(intx).Caption, "&", "") & "')"
                End Select
            End If
        End If
    Next intx

    MsgBox "Data Sudah diupdate"
    PreviewMenu
End Sub

Private Sub PreviewMenu()
    Dim rs  As ADODB.Recordset
    Dim sql As String

    Set rs = New ADODB.Recordset
    sql = "select * from tblUtlMenu"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

    If rs.RecordCount > 0 Then
        Set VSFlexGrid1.DataSource = rs
    End If
End Sub

Private Sub Form_Load()
    Call FormSize(8850, 7760, Me)
    Call LoadCentreForm(Me)
End Sub

VERSION 5.00
Object = "{B10DFE52-7887-11D5-9980-00C0A836120A}#28.0#0"; "ComboboxLB.OCX"
Begin VB.Form FrmUtlCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company"
   ClientHeight    =   1785
   ClientLeft      =   5265
   ClientTop       =   5700
   ClientWidth     =   5175
   LinkTopic       =   "MDIGL"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin Combo.ComboBoxLB cbCompany 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         ColumnCount     =   2
         Appearance      =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmUtlCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cbCompany.Text = "" Then
    MsgBox "Pilih Company terlebih dahulu.!", , "Info"
    Exit Sub
End If

If MsgBox("Mengganti Company akan menutup semua Form Yang terbuka, " & vbCrLf & _
        "jika ada data yang belum tersimpan silahkan dicek kembali. " & vbCrLf & _
        "Anda yakin mau Ganti Company ? ", vbYesNo, AT) = vbYes Then
    Call TutupSemuaForm
End If
MDIProject.CompanyName = cbCompany.Value
MDIProject.CompanyID = cbCompany.Column(1)
MDIProject.Caption = "[[ Administrasi Sumur Pantau ]] [[ Ver." & Trim(App.Major) & "." & Trim(App.Minor) & "." & Trim(App.Revision) & " ]] [[ " & cbCompany.Column(0) & " ]] " & " - [User : " & MDIProject.UserID & "]" & " - [PERIODE : " & Format(MDIProject.Periode, "MMMM yyyy") & "]"

Unload Me
End Sub

Private Sub Command2_Click()
If cbCompany.Text = "" Then
    MsgBox "Pilih Company terlebih dahulu.!", , "Info"
    Exit Sub
End If
Unload Me
End Sub

Private Sub Form_Load()
Call LoadCentreForm(Me)
cbCompany.Text = MDIProject.CompanyName
'Picture1.BackColor = MDIGL.ACPRibbon1.BackColor
'Me.Icon = MDIGL.Icon
Call ListCompany(cbCompany)
End Sub

Private Sub TutupSemuaForm()
Dim F As Form
For Each F In Forms
If F.Name <> "MDIProject" Then
    If F.Name <> "FrmUtlLogin" Then
        If F.Name <> "FrmUtlCompany" Then
            Unload F
        End If
    End If
End If
Next F
End Sub

Public Sub ListCompany(cbo As ComboBoxLB)
Dim rs As New ADODB.Recordset
Dim sql As String

    sql = "Select CompanyID, CompanyName From tblmstCompany"
    Set rs = New ADODB.Recordset
    rs.Open sql, ActiveCn, adOpenKeyset, adLockOptimistic
    
    With cbo
        .Clear
        .ColumnCount = 2
        .ColumnWidths = .Width & ";100"
        If Not rs.EOF Then
            Do Until rs.EOF
                .AddItem rs!CompanyName & ";" & rs!CompanyID
                rs.MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub

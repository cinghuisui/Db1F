VERSION 5.00
Begin VB.Form FrmUtlGantiPassWord 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Password"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtPass3 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "  "
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtPass2 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "  "
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtPass1 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "  "
         Top             =   360
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   120
         Picture         =   "FrmUtlGantiPassword.frx":0000
         ScaleHeight     =   1035
         ScaleWidth      =   915
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin MyLASP.isButton cmdSave 
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Icon            =   "FrmUtlGantiPassword.frx":0B4C
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
         Left            =   4200
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Icon            =   "FrmUtlGantiPassword.frx":1814
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Ulang"
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
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
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
         TabIndex        =   5
         Top             =   840
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Lama"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1620
      End
   End
End
Attribute VB_Name = "FrmUtlGantiPassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub GantiPW()
On Error GoTo ErrorHandler
Dim rsGTPW As ADODB.Recordset
Dim sql As String

Set rsGTPW = New ADODB.Recordset
sql = "select UserID,Password from tblUtlUser where UserID='" & MDIProject.UserID & "'"
rsGTPW.Open sql, ActiveCn, adOpenKeyset, adLockOptimistic

If Len(Trim(txtPass1.Text)) = 0 Or Len(Trim(txtPass2.Text)) = 0 Or Len(Trim(txtPass3.Text)) = 0 Then
   MsgBox "Data harus diisi lengkap", vbInformation, "Ganti Password"
   txtPass1.SetFocus
   Exit Sub
   
ElseIf Trim(txtPass1.Text) <> MDIProject.Password Then
   MsgBox "Password anda salah, silahkan ulangi lagi", vbInformation, "Ganti Password"
   txtPass1.SetFocus
   SendKeys "{Home}+{End}"
   Exit Sub
   
ElseIf Trim(txtPass2.Text) <> Trim(txtPass3.Text) Then
   MsgBox "Password pertama dan kedua anda salah, silahkan masukkan data yang benar", vbInformation, "Ganti Password"
   txtPass2.SetFocus
   SendKeys "{Home}+{End}"
   Exit Sub
Else
   rsGTPW!Password = EncryptText(Trim(txtPass3.Text), MDIProject.UserID)  'Trim(txtpass3.text)
   rsGTPW.Update
      
   MsgBox "Password sudah diganti dengan yang baru", vbInformation, "Ganti Password"
   Unload Me
End If

ErrorHandler:
Select Case Err.Number
Case 3021
  Resume Next
Case 3705
  rsGTPW.Close
  Resume
End Select

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    GantiPW
End Sub

Private Sub Form_Load()
    txtPass1.Text = ""
    txtPass2.Text = ""
    txtPass3.Text = ""
'    Frame1.BackColor = MDIProject.ACPRibbon1.BackColor
'    Call FormSize(2955, 6935, Me)
'    Call LoadCentreForm(Me)
'    Call ControlCentreForm(Me, Frame1)

End Sub



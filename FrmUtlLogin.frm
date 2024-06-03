VERSION 5.00
Begin VB.Form FrmUtlLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   3285
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   3285
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5415
      Begin VB.Label lblKOPERASIKARYAWAN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administrasi Sumur Pantau"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1155
         TabIndex        =   5
         Top             =   240
         Width           =   3105
      End
   End
   Begin MyLASP.isButton CmdCancel 
      Height          =   570
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1005
      Icon            =   "FrmUtlLogin.frx":0000
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
   Begin MyLASP.isButton cmdLogin 
      Height          =   570
      Left            =   600
      TabIndex        =   2
      Top             =   3240
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   1005
      Icon            =   "FrmUtlLogin.frx":0B94
      Style           =   5
      Caption         =   "&Login"
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   3180
      Left            =   0
      Picture         =   "FrmUtlLogin.frx":16B6
      Top             =   720
      Width           =   5400
   End
End
Attribute VB_Name = "FrmUtlLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MenuAkses As String
Public LoginSucceeded As Boolean

Private Sub CmdClose_Click()

End Sub

Private Sub Form_Load()
    'txtUserName.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdLogin_Click()
    'On Error GoTo ErrorHandler
    Dim rsPassword As ADODB.Recordset
    Dim sql        As String
    Dim CekLogin   As String
    Dim cn         As New ADODB.Connection

    Call cn.Open(ActiveCn)

    If Trim(txtUserName.Text) = "" Or Trim(txtPassword.Text) = "" Then
        MsgBox "UserName atau Password harus diisi"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        Me.MousePointer = vbNormal
        MDIProject.MousePointer = vbDefault
        Exit Sub
    End If

    Set rsPassword = New ADODB.Recordset
    sql = "Select * from tblUtlUser where UserID='" & Trim(txtUserName.Text) & "'"
    rsPassword.Open sql, ActiveCn, adOpenKeyset, adLockOptimistic

    If rsPassword.RecordCount = 0 Then
        MsgBox "User Anda dengan User Name " & Trim(txtUserName.Text) & " belum terdaftar, silahkan cek ", vbInformation, AT
        txtUserName.SetFocus
        SendKeys "{Home}+{end}"
        Me.MousePointer = vbNormal
        MDIProject.MousePointer = vbDefault
        Exit Sub
    Else

        If UCase(Trim(txtPassword.Text)) = DecryptText(rsPassword!Password, Trim(txtUserName.Text)) Then 'rsPassWord!Password Then
            MDIProject.UserID = rsPassword!UserID
            MDIProject.NamaUser = rsPassword!nama
            MDIProject.Password = DecryptText(rsPassword!Password, Trim(txtUserName.Text)) 'rsPassWord!Password
            MDIProject.Jabatan = rsPassword!Jabatan
            MDIProject.Wilayah = rsPassword!KodeWilayah
            MDIProject.GroupUser = rsPassword!NamaGroup
            MDIProject.Periode = Date
            MDIProject.Caption = "[[ Administrasi Sumur Pantau ]] " & Trim(App.Major) & "." & Trim(App.Minor) & "." & Trim(App.Revision) & " - << PT. RIAU SAKTI UNITED PLANTATION>> " & " - [User : " & MDIProject.UserID & "]" & " - [PERIODE : " & Format(MDIProject.Periode, "MMMM yyyy") & "]"
      
          '  MDIProject.StatusBar1.Panels(3) = rsPassword!UserID
            MenuAkses = rsPassword!NamaGroup
            EnabledMenu
            'FrmUtlCompany.Show 1
            LoginSucceeded = True
            Me.Hide
            
            'FrmUtlCompany.Show 1
            

        Else
            MsgBox "Password anda salah, coba lagi!", , "Login"
            txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
        End If
    End If

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13: SendKeys "{tab}", True
        cmdLogin_Click
    End Select
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case 13: SendKeys "{tab}", True
        txtPassword.SetFocus
    End Select
End Sub

Private Sub EnabledMenu()
On Error GoTo EH
Dim rsDetail As ADODB.Recordset
Dim sqlDetail As String
Dim i, intx As Long

Set rsDetail = New ADODB.Recordset
sqlDetail = "select IndekMenu from tblUtlMenu where IndekMenu not in (select IndekMenu from tblUtlGroupUserDtl where NamaGroup='" & MenuAkses & "') order by IndekMenu"
rsDetail.Open sqlDetail, ActiveCn, adOpenKeyset, adLockReadOnly

If rsDetail.RecordCount > 0 Then
    rsDetail.MoveFirst
    For i = 0 To rsDetail.RecordCount - 1
        For intx = 0 To MDIProject.Controls.Count - 1
            If TypeOf MDIProject.Controls(intx) Is Menu Then
                If (MDIProject.Controls(intx).Index <> 0) Then 'Index 0 yaitu form utama
                    If rsDetail!IndekMenu = MDIProject.Controls(intx).Index Then
                        MDIProject.Controls(intx).Enabled = False
                    End If
                End If
            End If
        Next intx
    rsDetail.MoveNext
    Next i
End If

EH:
If Err.Number <> 0 Then
    Resume Next
End If

End Sub


VERSION 5.00
Begin VB.MDIForm MDIProject 
   BackColor       =   &H8000000C&
   Caption         =   "Administrasi Sumur Pantau"
   ClientHeight    =   8085
   ClientLeft      =   4650
   ClientTop       =   3630
   ClientWidth     =   16140
   Icon            =   "MDIProject.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MyLASP.ACPRibbon ACPRibbon1 
      Align           =   3  'Align Left
      Height          =   8085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   16140
      _ExtentX        =   28469
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
   End
   Begin VB.Menu MnuMaster 
      Caption         =   "Master"
      Index           =   1000
      Begin VB.Menu MnuGeneral 
         Caption         =   "General"
         Index           =   1100
         Begin VB.Menu MnuMstWilayah 
            Caption         =   "Master Wilayah"
            Index           =   1101
         End
         Begin VB.Menu MnuMstDivisi 
            Caption         =   "Master Divisi"
            Index           =   1102
         End
         Begin VB.Menu MnuMstAfdeling 
            Caption         =   "Master Afdeling"
            Index           =   1103
         End
         Begin VB.Menu MnuMstPersil 
            Caption         =   "Master Persil"
            Index           =   1104
         End
         Begin VB.Menu MnuMstSetRekapitulasiPersil 
            Caption         =   "Master Set Rekapitulasi Persil"
            Index           =   1105
         End
      End
   End
   Begin VB.Menu MnuTransaksi 
      Caption         =   "Transaksi"
      Index           =   2000
      Begin VB.Menu MnuTrnLevelAir 
         Caption         =   "Level Air"
         Index           =   2100
         Begin VB.Menu MnuTrnRekapitulasiSumurPantau 
            Caption         =   "Level Air Sumur Pantau"
            Index           =   2110
         End
         Begin VB.Menu MnuTrnLevelAirKanal 
            Caption         =   "Level Air Kanal"
            Index           =   2120
         End
         Begin VB.Menu MnuTrnCurahHujan 
            Caption         =   "Level Curah Hujan"
            Index           =   2130
         End
      End
      Begin VB.Menu MnuTrnLogger 
         Caption         =   "Logger"
         Index           =   2200
         Begin VB.Menu MnuTrnLASPLogger 
            Caption         =   "Level Air Sumur Pantau (Logger)"
            Index           =   2210
         End
      End
      Begin VB.Menu MnuApproval 
         Caption         =   "Approval Level Air"
         Index           =   2300
         Begin VB.Menu MnuApprovalSumurPantau 
            Caption         =   "Approval Sumur Pantau"
            Index           =   2310
         End
         Begin VB.Menu MnuApprovalAirKanal 
            Caption         =   "Approval Air Kanal"
            Index           =   2320
         End
         Begin VB.Menu MnuApprovalCurahHujan 
            Caption         =   "Approval Curah Hujan"
            Index           =   2330
         End
      End
   End
   Begin VB.Menu MnuMonitoring 
      Caption         =   "Monitoring"
      Index           =   3000
      Begin VB.Menu MnuMonitoringLASP 
         Caption         =   "Monitoring Level Air Sumur Pantau"
         Index           =   3100
         Begin VB.Menu MnuMonitoringLASPMingguan 
            Caption         =   "Monitoring LASP Mingguan"
            Index           =   3110
         End
         Begin VB.Menu MnuMonitoringLASPBulanan 
            Caption         =   "Monitoring LASP Bulanan"
            Index           =   3120
         End
      End
      Begin VB.Menu MnuMonLevelAirKanal 
         Caption         =   "Monitoring Level Air Kanal"
         Index           =   3200
         Begin VB.Menu MnuMonLevelAirKanalMingguan 
            Caption         =   "Monitoring LAK Mingguan"
            Index           =   3210
         End
         Begin VB.Menu MnuMonLevelAirKanalBulanan 
            Caption         =   "Monitoring LAK Bulanan"
            Index           =   3220
         End
      End
      Begin VB.Menu MnuMonitoringCRH 
         Caption         =   "Monitoring Curah Hujan"
         Index           =   3300
         Begin VB.Menu MnuMonCRHMingguan 
            Caption         =   "Monitoring CRH Mingguan"
            Index           =   3310
         End
         Begin VB.Menu MnuMonCRHBulanan 
            Caption         =   "Monitoring CRH Bulanan"
            Index           =   3320
         End
      End
   End
   Begin VB.Menu MnuReport 
      Caption         =   "Laporan"
      Index           =   4000
      Begin VB.Menu MnuRptSumurPantau 
         Caption         =   "Laporan Level Air Sumur Pantau"
         Index           =   4100
         Begin VB.Menu MnuLaporanMingguanLASP 
            Caption         =   "Laporan Mingguan"
            Index           =   4110
            Begin VB.Menu MnuRptLASPMingguan 
               Caption         =   "Laporan LASP Mingguan"
               Index           =   4111
            End
            Begin VB.Menu MnuRptLASPMingguanPerWil 
               Caption         =   "Laporan LASP Mingguan PerWil"
               Index           =   4112
            End
         End
         Begin VB.Menu MnuLaporanBulananLASP 
            Caption         =   "Laporan Bulanan"
            Index           =   4120
            Begin VB.Menu MnuRptLASPBulanan 
               Caption         =   "Laporan LASP Bulanan"
               Index           =   4121
            End
            Begin VB.Menu MnuRptLASPBulananWil 
               Caption         =   "Laporan LASP Bulanan Per Wil"
               Index           =   4122
            End
         End
      End
      Begin VB.Menu MnuRptLevelAirKanal 
         Caption         =   "Laporan Level Air Kanal"
         Index           =   4200
         Begin VB.Menu MnuRptLAKMingguan 
            Caption         =   "Laporan LAK Mingguan"
            Index           =   4210
         End
         Begin VB.Menu MnuRptLAKBulanan 
            Caption         =   "Laporan LAK Bulanan"
            Index           =   4220
         End
      End
      Begin VB.Menu MnuRptCurahHujan 
         Caption         =   "Laporan Curah Hujan"
         Index           =   4300
         Begin VB.Menu MnuRptCRHMingguan 
            Caption         =   "Laporan CRH Mingguan"
            Index           =   4310
         End
         Begin VB.Menu MnuRptCRHBulanan 
            Caption         =   "Laporan CRH Bulanan"
            Index           =   4320
         End
      End
      Begin VB.Menu MnuRptGrafik 
         Caption         =   "Laporan Grafik"
         Index           =   4400
         Begin VB.Menu MnuRptGrafikMingguan 
            Caption         =   "Laporan Grafik Mingguan"
            Index           =   4410
            Begin VB.Menu MnuRptGrafikLASP 
               Caption         =   "Laporan Grafik LASP"
               Index           =   4411
            End
            Begin VB.Menu MnuRptGrafikLASPWil 
               Caption         =   "Laporan Grafik LASP PerWil"
               Index           =   4412
            End
         End
         Begin VB.Menu MnuRptGrafikLAKMingguan 
            Caption         =   "Laporan Grafik LAK Mingguan"
            Index           =   4420
            Begin VB.Menu MnuRptGrafikLAKAll 
               Caption         =   "Laporan Grafik LAK"
               Index           =   4421
            End
            Begin VB.Menu MnuRptGrafikLAKWil 
               Caption         =   "Laporan Grafik LAK PerWil"
               Index           =   4422
            End
         End
         Begin VB.Menu MnuRptGrafikCRH 
            Caption         =   "Laporan Grafik Curah Hujan"
            Index           =   4430
            Begin VB.Menu MnuRptGrafikCRHAll 
               Caption         =   "Laporan Grafik CRH"
               Index           =   4431
            End
            Begin VB.Menu MnuRptGrafikCRHWil 
               Caption         =   "Laporan Grafik CRH PerWil"
               Index           =   4432
            End
         End
      End
   End
   Begin VB.Menu MnuUtility 
      Caption         =   "Utility"
      Index           =   5000
      Begin VB.Menu MnuUtlUser 
         Caption         =   "User"
         Index           =   5100
      End
      Begin VB.Menu MnuUtlUserGroup 
         Caption         =   "User Group"
         Index           =   5200
      End
      Begin VB.Menu FrmUtlAutomaticMenuEntry 
         Caption         =   "Automatic Menu Entry"
         Index           =   5300
      End
      Begin VB.Menu MnuUtlGantiPassword 
         Caption         =   "Ganti Password"
         Index           =   5400
      End
   End
   Begin VB.Menu MnuWindow 
      Caption         =   "Window"
      Index           =   6000
      WindowList      =   -1  'True
      Begin VB.Menu MnuVertical 
         Caption         =   "Vertical"
         Index           =   6100
      End
      Begin VB.Menu MnuHorizontal 
         Caption         =   "HoriZontal"
         Index           =   6200
      End
      Begin VB.Menu MnuCascade 
         Caption         =   "Cascade"
         Index           =   6300
      End
   End
   Begin VB.Menu MnuExit 
      Caption         =   "Exit"
      Index           =   7000
   End
End
Attribute VB_Name = "MDIProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public cn As ADODB.Connection
Public UserID, NamaUser, Password, Jabatan, Wilayah, GroupUser As String
Public Periode     As Date
Public No          As Integer
Public CompanyName As String
Public CompanyID   As Integer
Dim mDeactivated   As Boolean
Dim Theme          As Integer

Private Sub MDIForm_Activate()
    If mDeactivated = False Then
        Me.MousePointer = vbHourglass
        FrmUtlLogin.Show vbModal
        Me.MousePointer = vbNormal
    End If
End Sub

Private Sub MDIForm_Deactivate()
    mDeactivated = True
End Sub

Private Sub MDIForm_Initialize()
    mDeactivated = False
End Sub

Private Sub MDIForm_Load()
    Theme = 1
    '# SET Theme
    ACPRibbon1.Theme = Theme    ' 0 - Black
    ' 1 - Blue
    ' 2 - Silver
    '# OPTIONAL - Load Background for Form.
    MDIProject.Picture = ACPRibbon1.LoadBackground
    '# OPTIONAL - Load Background for Form
    MDIProject.BackColor = ACPRibbon1.BackColor
End Sub

Private Sub TampilForm(frm As Form, Optional Modal As Boolean)

    If Modal = True Then
        frm.Show vbModal
    Else
        frm.Show
        frm.SetFocus
    End If
    '# Set Theme for new Child Form
    frm.Picture = ACPRibbon1.LoadBackground
    frm.BackColor = ACPRibbon1.BackColor
End Sub

Private Sub MnuApprovalSumurPantau_Click(Index As Integer)
    Call TampilForm(FrmApprovalLevelAir, False)
End Sub

Private Sub MnuExit_Click(Index As Integer)
    End
End Sub

Private Sub MnuHorizontal_Click(Index As Integer)
    Me.Arrange vbHorizontal
End Sub

Private Sub MnuCascade_Click(Index As Integer)
    Me.Arrange vbCascade
End Sub

Private Sub MnuRptCRHBulanan_Click(Index As Integer)
    FrmRptCRHMingguan.FormKontrol = "ALL"
    Call TampilForm(FrmRptCRHMingguan, False)
End Sub

Private Sub MnuRptCRHMingguan_Click(Index As Integer)
    FrmRptCRHMingguan.FormKontrol = "WEEK"
    Call TampilForm(FrmRptCRHMingguan, False)
End Sub

Private Sub MnuRptGrafikCRHAll_Click(Index As Integer)
    FrmRptGrafikCurahHujan.FormKontrol = "ALL"
    Call TampilForm(FrmRptGrafikCurahHujan, False)
End Sub

Private Sub MnuRptGrafikCRHWil_Click(Index As Integer)
    FrmRptGrafikCurahHujan.FormKontrol = "Wil"
    Call TampilForm(FrmRptGrafikCurahHujan, False)
End Sub

Private Sub MnuRptGrafikLAKAll_Click(Index As Integer)
    FrmRptGrafikLAK.FormKontrol = "ALL"
    Call TampilForm(FrmRptGrafikLAK, False)
End Sub

Private Sub MnuRptGrafikLAKWil_Click(Index As Integer)
    FrmRptGrafikLAK.FormKontrol = "Wil"
    Call TampilForm(FrmRptGrafikLAK, False)
End Sub

Private Sub MnuRptGrafikLASP_Click(Index As Integer)
    FrmRptGrafikPerWil.FormKontrol = "ALL"
    Call TampilForm(FrmRptGrafikPerWil, False)
End Sub

Private Sub MnuRptGrafikLASPWil_Click(Index As Integer)
    FrmRptGrafikPerWil.FormKontrol = "WIL"
    Call TampilForm(FrmRptGrafikPerWil, False)
End Sub

Private Sub MnuRptLAKBulanan_Click(Index As Integer)
    FrmRptLAKMingguan.FormKontrol = "ALL"
    Call TampilForm(FrmRptLAKMingguan, False)
End Sub

Private Sub MnuRptLAKMingguan_Click(Index As Integer)
    FrmRptLAKMingguan.FormKontrol = "WEEK"
    Call TampilForm(FrmRptLAKMingguan, False)
End Sub

Private Sub MnuRptLASPBulanan_Click(Index As Integer)
    FrmRptLASPBulanan.FormKontrol = "ALL"
    Call TampilForm(FrmRptLASPBulanan, False)
End Sub

Private Sub MnuRptLASPBulananWil_Click(Index As Integer)
    FrmRptLASPBulanan.FormKontrol = "WIL"
    Call TampilForm(FrmRptLASPBulanan, False)
End Sub

Private Sub MnuRptLASPMingguan_Click(Index As Integer)
    FrmRptLASPMingguan.FormKontrol = "ALL"
    Call TampilForm(FrmRptLASPMingguan, False)
End Sub

Private Sub MnuRptLASPMingguanPerWil_Click(Index As Integer)
    FrmRptLASPMingguan.FormKontrol = "WIL"
    Call TampilForm(FrmRptLASPMingguan, False)
End Sub

Private Sub MnuUtlGantiPassword_Click(Index As Integer)
    Call TampilForm(FrmUtlGantiPassWord, False)
End Sub

Private Sub MnuVertical_Click(Index As Integer)
    Me.Arrange vbVertical
End Sub

Private Sub FrmUtlAutomaticMenuEntry_Click(Index As Integer)
    Call TampilForm(FrnUtlAutomaticMenuEntry, False)
End Sub

Private Sub MnuMonCRHBulanan_Click(Index As Integer)
    Call TampilForm(FrmMonCurahHujanBulanan, False)
End Sub

Private Sub MnuMonCRHMingguan_Click(Index As Integer)
    Call TampilForm(FrmMonCurahHujanMingguan, False)
End Sub

Private Sub MnuMonitoringLASPMingguan_Click(Index As Integer)
    Call TampilForm(FrmMonLevelAirSumurPantauMingguan, False)
End Sub

Private Sub MnuMonitoringLASPBulanan_Click(Index As Integer)
    Call TampilForm(FrmMonLevelAirSumurPantauBulanan, False)
End Sub

Private Sub MnuMonLevelAirKanalBulanan_Click(Index As Integer)
    Call TampilForm(FrmMonLAKBulanan, False)
End Sub

Private Sub MnuMonLevelAirKanalMingguan_Click(Index As Integer)
    Call TampilForm(FrmMonLAKMingguan, False)
End Sub

Private Sub MnuMstAfdeling_Click(Index As Integer)
    Call TampilForm(FrmMstAfdeling, False)
End Sub

Private Sub MnuMstDivisi_Click(Index As Integer)
    Call TampilForm(FrmMstDivisi, False)
End Sub

Private Sub MnuMstPersil_Click(Index As Integer)
    Call TampilForm(FrmMstPersil, False)
End Sub

Private Sub MnuMstSetRekapitulasiPersil_Click(Index As Integer)
    Call TampilForm(FrmMstSetRekapitulasiPersil, False)
End Sub

Private Sub MnuMstWilayah_Click(Index As Integer)
    Call TampilForm(FrmMstWilayah, False)
End Sub

Private Sub MnuTrnCurahHujan_Click(Index As Integer)
    Call TampilForm(FrmTrnLevelCurahHujan, False)
End Sub

Private Sub MnuTrnLevelAirKanal_Click(Index As Integer)
    Call TampilForm(FrmTrnLevelAirKanal, False)
End Sub

Private Sub MnuTrnRekapitulasiSumurPantau_Click(Index As Integer)
    Call TampilForm(FrmTrnLevelAirSumurPantau, False)
End Sub

Private Sub MnuUtlUser_Click(Index As Integer)
    Call TampilForm(FrmUtlUser, False)
End Sub

Private Sub MnuUtlUserGroup_Click(Index As Integer)
    Call TampilForm(FrmUtlGroupMenu, False)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Select Case Button.Key
'
'        Case "Home"
'            MsgBox "User Home"
'        Case "Persil"
'            Call TampilForm(FrmMstPersil, False)
'        Case "Grafik"
'            FrmRptGrafikPerWil.FormKontrol = "ALL"
'            Call TampilForm(FrmRptGrafikPerWil, False)
'
'    End Select
End Sub

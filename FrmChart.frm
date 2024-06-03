VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmChart 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5535
      Left            =   240
      OleObjectBlob   =   "FrmChart.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "FrmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call LoadDataChart
End Sub

Sub LoadDataChart()
        Dim rs  As ADODB.Recordset
    Dim SQL As String
    Dim i As Integer
    Dim j As Integer
    Dim DataSebelum As String
    Set rs = New ADODB.Recordset
    SQL = "Select DetailID,Wilayah,KodePersil,KoordinatX,KoordinatY,Nilai,Keterangan from vwRekapitulasiSumurPantau where HeaderID='" & HeaderID & "'order by DetailID Asc"
    rs.Open SQL, ActiveCn, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        With MSChart1
            ' pengaturan grafik
            .AllowSelections = False
            .DoSetCursor = True
            .MousePointer = VtMousePointerArrowQuestion
            .ChartType = VtChChartType2dBar
            
            ' menampilkan legend
            .ShowLegend = True
            
            ' mengatur teks legend
            With MSChart1.Plot.SeriesCollection(1)
                .LegendText = "Suhu Maksimum - (C)"
            End With
            With MSChart1.Plot.SeriesCollection(2)
                .LegendText = "Suhu Minimum - (C)"
            End With
            With .Legend
                .Location.LocationType = VtChLocationTypeTop
                .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
                .VtFont.VtColor.Set 255, 0, 0
                .Backdrop.Fill.Style = VtFillStyleBrush
                .Backdrop.Fill.Brush.Style = VtBrushStyleSolid
                .Backdrop.Fill.Brush.FillColor.Set 219, 230, 255
            End With
            
            ' mengatur judul grafik
            MSChart1.Title = "Perbandingan Suhu Kota di Indonesia"
            With MSChart1.Title.VtFont
                .Name = "Calibri"
                .Size = 20
                .Effect = VtFontEffectUnderline
            End With
            
            ' mengatur title untuk sumbu x dan y
            With MSChart1.Plot.Axis(1, 1)
                .AxisTitle.VtFont.Size = 9
                .AxisTitle.VtFont.Name = "Calibri"
                .AxisTitle.VtFont.Effect = Bold
                .AxisTitle.Visible = True
                .AxisTitle.Text = "Suhu dalam derajat Celcius"
            End With
            
            With MSChart1.Plot.Axis(0, 1)
                .AxisTitle.VtFont.Size = 9
                .AxisTitle.VtFont.Name = "Calibri"
                .AxisTitle.VtFont.Effect = Bold
                .AxisTitle.Visible = True
                .AxisTitle.Text = "Nama Kota"
            End With
            
            ' mengatur footnote grafik
            MSChart1.Footnote = "Sumber Data: Badan Metereologi Cuaca Batam - 2014"
            
            ' mengatur warna grafik
            With MSChart1.Plot.SeriesCollection(1)
                .DataPoints(-1).Brush.FillColor.Set 45, 44, 78
            End With
            
            ' mengatur warna background grafik
            With MSChart1.Backdrop.Fill
                .Style = VtFillStyleBrush
                .Brush.FillColor.Set 255, 255, 255
            End With
        End With
    End If
End Sub

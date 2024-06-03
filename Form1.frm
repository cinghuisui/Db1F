VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OpenPDF 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PDFViewer1 
      Height          =   4815
      Left            =   1200
      ScaleHeight     =   4755
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OpenPDF_Click()
With CommonDialog1
.DefaultExt = "pdf"
.Filter = "PDF File Formats (*.pdf)|*.pdf|All Files (*.*) | *.* ||"
.FilterIndex = 1
End With
CommonDialog1.ShowOpen
PDFViewer1.LoadFile CommonDialog1.FileName
End Sub


Private Sub EDOffice_DocumentOpened()
EDOffice1.ProtectDoc 1 ' XlProtectTypeNormal
End Sub

'
'Private Sub DisableAdobeReader_Click()
'PDFViewer1.SetReadOnly
'PDFViewer1.DisableHotKeyCopy
'PDFViewer1.DisableHotKeyPrint
'PDFViewer1.DisableHotKeySave
'PDFViewer1.DisableHotKeySearch
'PDFViewer1.DisableHotKeyShowBookMarks
'PDFViewer1.DisableHotKeyShowThumnails
'PDFViewer1.DisableHotKeyShowToolbars
'End Sub

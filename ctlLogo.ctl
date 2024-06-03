VERSION 5.00
Begin VB.UserControl csLogo 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000009&
   ScaleHeight     =   4545
   ScaleWidth      =   645
End
Attribute VB_Name = "csLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** Control to create a sidebar on a form
'** Based on code from Steve McMahon at www.vbaccelerator.com

'** Created by Claudius Schultz, Freeware@necsus.dk


Private Type RECT
    left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private m_bRGBStart(1 To 3) As Integer
Private m_bRGBEnd(1 To 3) As Integer

'** Default properties...
Const m_def_FontAutoSize = True
Const m_def_Caption = "csLogo"
Const m_def_StartColor = &H800000
Const m_def_EndColor = &HFFC0C0
Const m_def_FontName = "Arial"
Const m_def_FontBold = True
Const m_def_FontItalic = False
Const m_def_FontSize = 16
Const m_def_FontUnderline = False
Const m_def_ForeColor = vbWhite

'** Property variables...
Dim m_FontAutoSize As Boolean
Dim m_Caption As String
Dim m_StartColor As OLE_COLOR
Dim m_EndColor As OLE_COLOR

'** Properties...
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Draw
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Gibt den Rahmenstil für ein Objekt zurück oder legt diesen fest."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    Draw
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Erzwingt ein vollständiges Neuzeichnen eines Objekts."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Text to be shown in the csLogo Box"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Draw
End Property

Public Property Get StartColor() As OLE_COLOR
Attribute StartColor.VB_Description = "The startcolor of the Gradient"
    StartColor = m_StartColor
End Property

Public Property Let StartColor(ByVal New_StartColor As OLE_COLOR)
    m_StartColor = New_StartColor
    PropertyChanged "StartColor"
    FillStartColor
    Draw
End Property

Public Property Get EndColor() As OLE_COLOR
Attribute EndColor.VB_Description = "The Endcolor of the Gradient"
    EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal New_EndColor As OLE_COLOR)
    m_EndColor = New_EndColor
    PropertyChanged "EndColor"
    FillEndColor
    Draw
End Property

Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_StartColor = m_def_StartColor
    m_EndColor = m_def_EndColor
    UserControl.FontName = m_def_FontName
    UserControl.FontBold = m_def_FontBold
    UserControl.FontItalic = m_def_FontItalic
    UserControl.FontSize = m_def_FontSize
    UserControl.FontUnderline = m_def_FontUnderline
    UserControl.ForeColor = m_def_ForeColor
    FillStartColor
    FillEndColor
    Draw
    m_FontAutoSize = m_def_FontAutoSize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_StartColor = PropBag.ReadProperty("StartColor", m_def_StartColor)
    m_EndColor = PropBag.ReadProperty("EndColor", m_def_EndColor)
    UserControl.FontName = PropBag.ReadProperty("FontName", "Arial")
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 10)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", True)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", False)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", False)

    FillStartColor
    FillEndColor
    Draw

    m_FontAutoSize = PropBag.ReadProperty("FontAutoSize", m_def_FontAutoSize)
End Sub

Private Sub UserControl_Resize()
    Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("StartColor", m_StartColor, m_def_StartColor)
    Call PropBag.WriteProperty("EndColor", m_EndColor, m_def_EndColor)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("FontAutoSize", m_FontAutoSize, m_def_FontAutoSize)
End Sub

Public Sub Draw()
    Dim lHeight As Long, lWidth As Long
    Dim lYStep As Long
    Dim lY As Long
    Dim bRGB(1 To 3) As Integer
    Dim tLF As LOGFONT
    Dim hFnt As Long
    Dim hFntOld As Long
    Dim lR As Long
    Dim rct As RECT
    Dim hBr As Long
    Dim hDC As Long
    Dim dR(1 To 3) As Double
    Dim tmpSize As Long
    Dim lSize As Long
    On Error GoTo DrawError

    hDC = UserControl.hDC
    lHeight = UserControl.Height \ Screen.TwipsPerPixelY
    rct.Right = UserControl.Width \ Screen.TwipsPerPixelY
    
    ' Set a graduation of 255 pixels:
    lYStep = lHeight \ 255
    If (lYStep = 0) Then
        lYStep = 1
    End If
    rct.Bottom = lHeight

    bRGB(1) = m_bRGBStart(1)
    bRGB(2) = m_bRGBStart(2)
    bRGB(3) = m_bRGBStart(3)
    dR(1) = m_bRGBEnd(1) - m_bRGBStart(1)
    dR(2) = m_bRGBEnd(2) - m_bRGBStart(2)
    dR(3) = m_bRGBEnd(3) - m_bRGBStart(3)

    For lY = lHeight To 0 Step -lYStep
        ' Draw bar:
        rct.Top = rct.Bottom - lYStep
        hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
        FillRect hDC, rct, hBr
        DeleteObject hBr
        rct.Bottom = rct.Top
        ' Adjust colour:
        bRGB(1) = m_bRGBStart(1) + dR(1) * (lHeight - lY) / lHeight
        bRGB(2) = m_bRGBStart(2) + dR(2) * (lHeight - lY) / lHeight
        bRGB(3) = m_bRGBStart(3) + dR(3) * (lHeight - lY) / lHeight
        'Debug.Print bRGB(1), (lHeight - lY) / lHeight
    Next lY

    '** Adjust the fontsize if needed
    'tmpSize = UserControl.FontSize
    If m_FontAutoSize Then
        lSize = 1

        Do
            lSize = lSize + 1
            UserControl.Font.Size = lSize
        Loop Until UserControl.TextHeight("Xg") > UserControl.Width

        lSize = lSize - 3
        UserControl.Font.Size = lSize
    End If

    pOLEFontToLogFont UserControl.Font, hDC, tLF
    tLF.lfEscapement = 900
    hFnt = CreateFontIndirect(tLF)
    If (hFnt <> 0) Then
        hFntOld = SelectObject(hDC, hFnt)
        lR = TextOut(hDC, 0, lHeight - 16, m_Caption, Len(m_Caption))
        SelectObject hDC, hFntOld
        DeleteObject hFnt
    End If

    UserControl.Refresh
    Exit Sub
    
DrawError:
    Debug.Print "Problem: " & Err.Description
End Sub

Private Sub pOLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
    Dim sFont As String
    Dim iChar As Integer

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough

    End With
End Sub

Private Sub FillStartColor()
    '** Fills an Array with the StartColor

    Dim lColor As Long
    
    OleTranslateColor m_StartColor, 0, lColor
    m_bRGBStart(1) = lColor And &HFF&
    m_bRGBStart(2) = ((lColor And &HFF00&) \ &H100)
    m_bRGBStart(3) = ((lColor And &HFF0000) \ &H10000)

End Sub

Private Sub FillEndColor()
    '** Fills an Array with the EndColor
    Dim lColor As Long
    OleTranslateColor m_EndColor, 0, lColor
    m_bRGBEnd(1) = lColor And &HFF&
    m_bRGBEnd(2) = ((lColor And &HFF00&) \ &H100)
    m_bRGBEnd(3) = ((lColor And &HFF0000) \ &H10000)

End Sub

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Gibt den Namen der Schriftart an, die in jeder Zeile für die gegebene Ebene verwendet wird."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    Draw
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Gibt die Größe der Schriftart (in Punkten) an, die in jeder Zeile für die gegebene Ebene verwendet wird."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    Draw
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Gibt Schriftstile für Fettschrift zurück oder legt diese fest."
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    Draw
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Gibt Schriftstile für Kursivschrift zurück oder legt diese fest."
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    Draw
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Gibt Schriftstile für unterstrichene Schrift zurück oder legt diese fest."
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
    Draw
End Property

Public Property Get FontAutoSize() As Boolean
Attribute FontAutoSize.VB_Description = "Adjust the Fontsize to the Controls width"
    FontAutoSize = m_FontAutoSize
End Property

Public Property Let FontAutoSize(ByVal New_FontAutoSize As Boolean)
    m_FontAutoSize = New_FontAutoSize
    PropertyChanged "FontAutoSize"
    Draw
End Property


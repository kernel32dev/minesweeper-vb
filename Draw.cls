VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function os_Polygon Lib "gdi32" Alias "Polygon" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Private Type Bitmap '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long


'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function GetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long


Private Declare Function OS_BeginPath Lib "gdi32" Alias "BeginPath" (ByVal hdc As Long) As Long
Private Declare Function OS_EndPath Lib "gdi32" Alias "EndPath" (ByVal hdc As Long) As Long
Private Declare Function OS_FillPath Lib "gdi32" Alias "FillPath" (ByVal hdc As Long) As Long
Private Declare Function OS_StrokeAndFillPath Lib "gdi32" Alias "StrokeAndFillPath" (ByVal hdc As Long) As Long
Private Declare Function OS_StrokePath Lib "gdi32" Alias "StrokePath" (ByVal hdc As Long) As Long
Private Declare Function OS_SetBkMode Lib "gdi32" Alias "SetBkMode" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function OS_DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As t_RECT, ByVal wFormat As Long) As Long
Private Declare Function OS_DrawEdge Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, qrc As t_RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OS_BitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OS_CreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" (ByVal crColor As Long) As Long
Private Declare Function OS_FillRect Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As t_RECT, ByVal hBrush As Long) As Long
Private Declare Function OS_SelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OS_DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function OS_SetTextColor Lib "gdi32" Alias "SetTextColor" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OS_LineTo Lib "gdi32.dll" Alias "LineTo" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function OS_DrawFocusRect Lib "user32" Alias "DrawFocusRect" (ByVal hdc As Long, lpRect As t_RECT) As Long
Private Declare Function OS_CreatePen Lib "gdi32" Alias "CreatePen" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OS_DrawIcon Lib "user32" Alias "DrawIcon" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function OS_DrawIconEx Lib "user32" Alias "DrawIconEx" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function OS_DrawFrameControl Lib "user32" Alias "DrawFrameControl" (ByVal hdc As Long, lpRect As t_RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function OS_InvertRect Lib "user32" Alias "InvertRect" (ByVal hdc As Long, lpRect As t_RECT) As Long
Private Declare Function OS2_DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As Long) As Long
Private Declare Function OS_CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal S As Long, ByVal c As Long, ByVal op As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function OS_StretchBlt Lib "gdi32" Alias "StretchBlt" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function OS_Ellipse Lib "gdi32" Alias "Ellipse" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OS_CreateHatchBrush Lib "gdi32" Alias "CreateHatchBrush" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function OS_CreateBrushIndirect Lib "gdi32" Alias "CreateBrushIndirect" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function OS_PlayEnhMetaFile Lib "gdi32" Alias "PlayEnhMetaFile" (ByVal hdc As Long, ByVal hemf As Long, lpRect As t_RECT) As Long
Private Declare Function OS_DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As t_RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OS_SetTextJustification Lib "gdi32" Alias "SetTextJustification" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare Function OS_SetTextCharacterExtra Lib "gdi32" Alias "SetTextCharacterExtra" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Enum EnumPenStyles
    '  Pen Styles
    PsSolid = 0
    PsDash = 1                    '  -------
    PsDot = 2                     '  .......
    PsDashDot = 3                 '  _._._._
    PsDashDotDot = 4              '  _.._.._
    PsNull = 5
    PsInsideFrame = 6
    PsUserStyle = 7
    PsAlternate = 8
    PsStyleMask = &HF
End Enum

Public Enum EnumHatchStyles       ' Hatch Styles
    HsHorizontal = 0              '  -----
    HsVertical = 1                '  |||||
    HsFDiagonal = 2               '  \\\\\
    HsBDiagonal = 3               '  /////
    HsCross = 4                   '  +++++
    HsDiagCross = 5               '  xxxxx
End Enum

Public Enum EnumSetBkMode
    BkOpaque = 2
    bkTransparent = 1
End Enum




Public Enum FrameControl_States
    DFCS_INACTIVE = &H100
    DFCS_PUSHED = &H200
    DFCS_CHECKED = &H400
    DFCS_ADJUSTRECT = &H2000
    DFCS_FLAT = &H4000
    DFCS_MONO = &H8000
End Enum

Public Enum FrameControl_Scroll
    DFCS_SCROLLUP = &H0
    DFCS_SCROLLDOWN = &H1
    DFCS_SCROLLLEFT = &H2
    DFCS_SCROLLRIGHT = &H3
    DFCS_SCROLLCOMBOBOX = &H5
    DFCS_SCROLLSIZEGRIP = &H8
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
End Enum

Public Enum FrameControl_Button
  DFCS_BUTTONCHECK = &H0
  DFCS_BUTTONRADIOIMAGE = &H1
  DFCS_BUTTONRADIOMASK = &H2
  DFCS_BUTTONRADIO = &H4
  DFCS_BUTTON3STATE = &H8
  DFCS_BUTTONPUSH = &H10
End Enum

'Const DFC_CAPTION = 1
'Const DFC_MENU = 2
Const DFC_SCROLL = 3
'Const DFC_BUTTON = 4

Public Enum EDrawIcon
    DI_MASK = &H1
    DI_IMAGE = &H2
    DI_NORMAL = &H3
    DI_COMPAT = &H4
    DI_DEFAULTSIZE = &H8
End Enum

Public Enum EBorders_Types
  dcOutRect = 2&
  dcInRect = 1&
End Enum

Public Enum EMaskAutoLegend
    AuNone
    AuByte
    AuText
End Enum

Public Enum EMask_Styles 'Não pode mudar
    FoText
    FoDate
    FoCurrency
    FoDouble
    FoLong
    FoCGC
    FoCPF
    FoCEP
    FoIE
    FoYesNo
    FoMemo
    FoPicture
    FoTime
    FoByteText03
    FoByteText02
    FoDateMMYYYY
    FoFone1
    FoTimeHHMM
    FoCurrency_3
    FoCurrency_4
    FoDDMMYY_HHMM
    FoCurrency_auto
End Enum

Const DT_TOP = &H0, DT_LEFT = &H0, DT_CENTER = &H1, DT_RIGHT = &H2
Const DT_VCENTER = &H4, DT_BOTTOM = &H8, DT_WORDBREAK = &H10, DT_SINGLELINE = &H20
Const DT_EXPANDTABS = &H40, DT_TABSTOP = &H80, DT_NOCLIP = &H100, DT_EXTERNALLEADING = &H200
Const DT_CALCRECT = &H400, DT_NOPREFIX = &H800, DT_INTERNAL = &H1000
Public Enum DrawTextFormatFlags
  dtLeft = DT_LEFT
  dtTop = DT_TOP
  dtCenter = DT_CENTER
  dtRight = DT_RIGHT
  dtVCenter = DT_VCENTER
  dtBottom = DT_BOTTOM
  dtNoPrefix = DT_NOPREFIX
  dtCalcRect = DT_CALCRECT
  dtWordBreak = DT_WORDBREAK
  dtSingleLine = DT_SINGLELINE
  dtSingleLine_VCenter = DT_SINGLELINE Or DT_VCENTER
End Enum


' Brush Styles
'Const BS_SOLID = 0
Const BS_NULL = 1
'Const BS_HOLLOW = BS_NULL
'Const BS_HATCHED = 2
'Const BS_PATTERN = 3
'Const BS_INDEXED = 4
'Const BS_DIBPATTERN = 5
'Const BS_DIBPATTERNPT = 6
'Const BS_PATTERN8X8 = 7
'Const BS_DIBPATTERN8X8 = 8
'Types //////////////==========================================
' Logical Brush (or Pattern)
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type


Dim uFont As Long

Public Function SetBkMode(ByVal hdc As Long, ByVal nBkMode As EnumSetBkMode) As Long
  SetBkMode = OS_SetBkMode(hdc, nBkMode)
End Function

Public Function TextWidth(ByVal hdc As Long, Text As String) As Long
Dim lpRect As t_RECT
    OS_DrawText hdc, Text, Len(Text), lpRect, DT_CALCRECT 'Or DT_WORDBREAK Or DT_LEFT
    TextWidth = lpRect.Right - lpRect.Left
End Function

Public Function TextHeight(ByVal hdc As Long, Text As String, Optional mTextWith As Long, Optional wFormat As DrawTextFormatFlags) As Long
Dim lpRect As t_RECT
    lpRect.Right = mTextWith
    OS_DrawText hdc, Text, Len(Text), lpRect, DT_CALCRECT Or wFormat
    TextHeight = lpRect.Bottom - lpRect.Top
End Function


Public Function BitBlt(ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
  BitBlt = OS_BitBlt(hDestDC, X, y, nWidth, nHeight, hSrcDC, xSrc, ySrc, dwRop)
End Function

Public Function BitBltCopy(ByVal hDestDC As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
  BitBltCopy = OS_BitBlt(hDestDC, 0, 0, nWidth, nHeight, hSrcDC, 0, 0, dwRop)
End Function

Public Function CreateSolidBrush(ByVal crColor As Long) As Long
  CreateSolidBrush = OS_CreateSolidBrush(crColor)
End Function

Friend Function FillRect(ByVal hdc As Long, lpRect As t_RECT, ByVal hBrush As Long) As Long
    
  FillRect = OS_FillRect(hdc, lpRect, hBrush)
End Function

Public Function SelectObject(ByVal hdc As Long, ByVal hObject As Long) As Long
  SelectObject = OS_SelectObject(hdc, hObject)
End Function

Public Function DestroyIcon(ByVal hIcon As Long) As Long
    DestroyIcon = OS_DestroyIcon(hIcon)
    If DestroyIcon = 0 Then
       Debug.Assert 0
    End If
End Function

Public Function DeleteObject(ByVal hObject As Long) As Long
    DeleteObject = OS_DeleteObject(hObject)
    If DeleteObject = 0 And hObject <> 0 Then
        Debug.Assert 0
        'frmErr5.Show
        'frmErr5.lstErr.AddItem "DeleteObjectEx " & DeleteObject
    End If
End Function

Public Function SetTextColor(ByVal hdc As Long, ByVal crColor As Long) As Long
  SetTextColor = OS_SetTextColor(hdc, crColor)
End Function

Friend Function Focus(ByVal hdc As Long, lpRect As t_RECT) As Long
    Focus = OS_DrawFocusRect(hdc, lpRect)
End Function


Friend Sub FillSolidRect(ByVal hdc As Long, lpRect As t_RECT, ByVal crColor As Long)
Dim hBrush As Long
  hBrush = OS_CreateSolidBrush(crColor)
  Call FillRect(hdc, lpRect, hBrush)
  Call DeleteObject(hBrush)
End Sub


Public Function DrawIcon(ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal hIcon As Long) As Long
    DrawIcon = OS_DrawIcon(hdc, X, y, hIcon)
End Function

Public Function DrawIconEx(ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, Optional ByVal istepIfAniCur As Long, Optional ByVal hbrFlickerFreeDraw As Long, Optional ByVal diFlags As EDrawIcon = DI_NORMAL) As Long
  DrawIconEx = OS_DrawIconEx(hdc, xLeft, yTop, hIcon, cxWidth, cyWidth, istepIfAniCur, hbrFlickerFreeDraw, diFlags)
End Function

Friend Function FrameControl_Scroll(ByVal hdc As Long, lpRect As t_RECT, ByVal un2 As FrameControl_Scroll, Optional ByVal State As FrameControl_States) As Long
  FrameControl_Scroll = OS_DrawFrameControl(hdc, lpRect, DFC_SCROLL, un2 Or State)
End Function


Friend Function InvertRect(ByVal hdc As Long, lpRect As t_RECT) As Long
  InvertRect = OS_InvertRect(hdc, lpRect)
End Function

Public Sub SetHdcFont(ByVal hdc As Long, ByVal hdcRef As Long)
Dim Font As Long
Const OBJ_FONT = 6
    Font = GetCurrentObject(hdcRef, OBJ_FONT)
    uFont = SelectObject(hdc, Font)
    'Call DeleteObject(aFont)
    'uFont = Font
End Sub

Private Sub Class_Terminate()
    If uFont <> 0 Then
        Call DeleteObject(uFont)
    End If
End Sub

Public Sub TileBltEx(ByVal hWndDest As Long, ByVal hBmpSrc As Long, ByVal bmX As Long, ByVal bmY As Long, ByVal bmWidth As Long, ByVal bmHeight As Long)
   '
   ' 32-Bit Tiling BitBlt Function
   ' Written by Karl E. Peterson, 9/22/96.
   ' Tiles a bitmap across the client area of destination window.
   '
   ' Parameters ************************************************************
   '   hWndDest:     hWnd of destination
   '   hBmpSrc:      hBitmap of source
   ' ***********************************************************************
   '
   'Dim bmp As BITMAP     ' Header info for passed bitmap handle
   ' Device context for source
   Dim hdcDest As Long   ' Device context for destination
   Dim hDCSrc As Long
   Dim hBmpTmp As Long   ' Holding space for temporary bitmap
   Dim dRect As t_RECT     ' Holds coordinates of destination rectangle
   Dim Rows As Long      ' Number of rows in destination
   Dim Cols As Long      ' Number of columns in destination
   Dim dX As Long        ' CurrentX in destination
   Dim dY As Long        ' CurrentY in destination
   Dim i As Long, j As Long
   'Minha modificação
   If hBmpSrc = 0 Then Exit Sub
   
   '
   ' Get destination rectangle and device context.
   '
   Call GetClientRect(hWndDest, dRect)
   hdcDest = GetDC(hWndDest)
   '
   ' Create source DC and select passed bitmap into it.
   '
   hDCSrc = CreateCompatibleDC(hdcDest)
   hBmpTmp = SelectObject(hDCSrc, hBmpSrc)
   '
   ' Get size information about passed bitmap, and
   ' Calc number of rows and columns to paint.
   '
   'Call GetObj(hBmpSrc, Len(bmp), bmp)
   Rows = dRect.Right \ bmWidth
   Cols = dRect.Bottom \ bmHeight
   '
   ' Spray out across destination.
   '
   For i = 0 To Rows
      dX = i * bmWidth
      For j = 0 To Cols
         dY = j * bmHeight
         Call BitBlt(hdcDest, dX, dY, bmWidth, bmHeight, hDCSrc, bmX, bmY, vbSrcCopy)
      Next j
   Next i
   '
   ' and clean up
   '
   Call SelectObject(hDCSrc, hBmpTmp)
   Call DeleteDC(hDCSrc)
   Call ReleaseDC(hWndDest, hdcDest)
End Sub

Friend Sub DrawJust(ByVal hdc As Long, ByVal Text As String, lpRect As t_RECT)
Dim Cont As Long, buffer As String ', Char As String, Liga As Boolean
'Dim lpRectCalc as t_RECT, tw As Long
Dim y As Long, TextAc As String, Cont2 As Long
Dim JustFail As Boolean, Lista As New Collection
Dim Pos13_10 As Long
 
Text = Text & " " & vbNewLine & " "

Pos13_10 = 1
Do
  Pos13_10 = InStr(Pos13_10, Text, Chr$(13))
  If Pos13_10 = 0 Then Exit Do
  Mid(Text, Pos13_10, 1) = " "
Loop

DrawJustGetP "", True

buffer = ""

Do
  TextAc = DrawJustGetP(Text)
  If TextAc = "" Then Exit Do
  
  If (DrawJustWidth(hdc, buffer & TextAc) < (lpRect.Right - lpRect.Left)) And Not Cont = Lista.Count - 1 Then
    buffer = buffer & TextAc
    If Asc(Right(TextAc, 1)) = 10 Then
      lpRect.Top = lpRect.Top + y
      y = DrawJustText(hdc, buffer, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
      buffer = ""
    End If
  Else
    If Asc(Right(TextAc, 1)) = 10 Then
      lpRect.Top = lpRect.Top + y
      y = DrawJustText(hdc, buffer, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
      buffer = ""
      lpRect.Top = lpRect.Top + y
      y = DrawJustText(hdc, TextAc, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
    Else
      If Cont = Lista.Count - 1 Then
        buffer = buffer & TextAc
      End If
      If Cont = Lista.Count - 1 Then
        lpRect.Top = lpRect.Top + y
      Else
        buffer = RTrim(buffer)
        lpRect.Top = lpRect.Top + y
        
          JustFail = True
          For Cont2 = 1 To 1000
            OS_SetTextJustification hdc, Cont2, 8
            If DrawJustWidth(hdc, buffer) = (lpRect.Right - lpRect.Left) Then
              JustFail = False
              Exit For
            ElseIf DrawJustWidth(hdc, buffer) > (lpRect.Right - lpRect.Left) Then
              OS_SetTextJustification hdc, Cont2 - 1, 8
              Exit For
            End If
          Next
          If JustFail And False Then
            For Cont2 = 1 To 100
              OS_SetTextCharacterExtra hdc, Cont2
              If DrawJustWidth(hdc, buffer) >= (lpRect.Right - lpRect.Left) Then
                Exit For
              End If
            Next
          End If
        
      End If
      y = DrawJustText(hdc, buffer, lpRect, DT_SINGLELINE Or DT_EXPANDTABS)
      OS_SetTextCharacterExtra hdc, 0
      OS_SetTextJustification hdc, 0, 2
      buffer = TextAc
    End If
  End If
  If lpRect.Top > lpRect.Bottom Then
    Exit Do
  End If
Loop
End Sub

Private Function DrawJustGetP(ByVal Text As String, Optional Reset As Boolean) As String
Static Cont2 As Long
Dim Char  As String, Liga As Boolean, buffer As String
If Reset Then
  Cont2 = 1
  Exit Function
End If
Do
'For Cont2 = 1 To Len(Text)
  Char = Mid(Text, Cont2, 1)
  If Char = " " And Not Char = Chr$(10) Then
    Liga = True
  Else
    If Liga Then
      If Char = Chr$(10) Then
        buffer = buffer & Char
        Char = ""
        Cont2 = Cont2 + 1
      End If
      DrawJustGetP = buffer
      Exit Function
    End If
  End If
  buffer = buffer & Char
  Cont2 = Cont2 + 1
  If Cont2 > Len(Text) Then
    Exit Do
  End If
Loop
End Function

Private Function DrawJustWidth(ByVal hdc As Long, ByVal Text As String) As Long
Dim lpRect As t_RECT
  DrawJustText hdc, Text, lpRect, DT_CALCRECT Or DT_EXPANDTABS
  DrawJustWidth = (lpRect.Right - lpRect.Left)
End Function

Private Function DrawJustText(ByVal hdc As Long, ByVal lpStr As String, lpRect As t_RECT, ByVal wFormat As Long) As Long
  If Right(lpStr, 1) = Chr$(10) Then
    lpStr = Left(lpStr, Len(lpStr) - 1)
  End If
  DrawJustText = OS_DrawText(hdc, lpStr, Len(lpStr), lpRect, wFormat)
End Function

Public Function StretchBlt(ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
  StretchBlt = OS_StretchBlt(hdc, X, y, nWidth, nHeight, hSrcDC, xSrc, ySrc, nSrcWidth, nSrcHeight, dwRop)
End Function

Public Function Ellipse(ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
  Ellipse = CBool(OS_Ellipse(hdc, X1, Y1, X2, Y2))
End Function

Public Function CreateHatchBrush(ByVal nIndex As EnumHatchStyles, ByVal crColor As Long) As Long
  CreateHatchBrush = OS_CreateHatchBrush(nIndex, crColor)
End Function

Friend Sub FillHatchRect(ByVal hdc As Long, lpRect As t_RECT, ByVal crColor As Long, ByVal nIndex As EnumHatchStyles)
Dim hBrush As Long
  hBrush = OS_CreateHatchBrush(nIndex, crColor)
  Call FillRect(hdc, lpRect, hBrush)
  Call DeleteObject(hBrush)
End Sub

Public Function CreatePen(ByVal nPenStyle As EnumPenStyles, ByVal nWidth As Long, ByVal crColor As Long) As Long
  CreatePen = OS_CreatePen(nPenStyle, nWidth, crColor)
End Function

Public Function CreateNullBrush() As Long
Dim lpLogBrush As LOGBRUSH
  lpLogBrush.lbStyle = BS_NULL
  CreateNullBrush = OS_CreateBrushIndirect(lpLogBrush)
End Function

Friend Function PlayEnhMetaFile(ByVal hdc As Long, ByVal hemf As Long, lpRect As t_RECT) As Long
    PlayEnhMetaFile = OS_PlayEnhMetaFile(hdc, hemf, lpRect)
End Function

Public Function BeginPath(ByVal hdc As Long) As Long
    BeginPath = OS_BeginPath(hdc)
End Function

Public Function EndPath(ByVal hdc As Long) As Long
    EndPath = OS_EndPath(hdc)
End Function

Public Function FillPath(ByVal hdc As Long) As Long
    FillPath = OS_FillPath(hdc)
End Function

Public Function StrokeAndFillPath(ByVal hdc As Long) As Long
    StrokeAndFillPath = OS_StrokeAndFillPath(hdc)
End Function

Public Function StrokePath(ByVal hdc As Long) As Long
    StrokePath = OS_StrokePath(hdc)
End Function

Public Function SetTextCharacterExtra(ByVal hdc As Long, ByVal nCharExtra As Long) As Long
    SetTextCharacterExtra = OS_SetTextCharacterExtra(hdc, nCharExtra)
End Function


Public Sub DrawBitmapEx(ByVal HdcOut As Long, ByVal hBitmap As Long, ByVal X As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, Optional ByVal Stretch As Boolean = False, Optional ByVal StretchMode As Long = 4)
Dim HdcTemp As Long, sBitMap As Bitmap
    HdcTemp = CreateCompatibleDC(HdcOut)
    OS_SelectObject HdcTemp, hBitmap
    If Stretch Then
        GetObject hBitmap, Len(sBitMap), sBitMap
        SetStretchBltMode HdcOut, StretchMode
        OS_StretchBlt HdcOut, X, y, w, h, HdcTemp, 0, 0, sBitMap.bmWidth, sBitMap.bmHeight, vbSrcCopy     'GetDeviceCaps(HdcTemp, 8), GetDeviceCaps(HdcTemp, 10)
    Else
        OS_BitBlt HdcOut, X, y, w, h, HdcTemp, 0, 0, vbSrcCopy
    End If
    OS2_DeleteDC HdcTemp
End Sub

Public Function StretchBltMode(ByVal hdc As Long, ByVal nStretchMode As Long) As Long
    StretchBltMode = SetStretchBltMode(hdc, nStretchMode)
End Function

Public Function DeleteDC(hdc As Long) As Long
    DeleteDC = OS2_DeleteDC(hdc)
    If DeleteDC = 0 And hdc <> 0 Then
        Debug.Assert 0
        'frmErr5.Show
        'frmErr5.lstErr.AddItem "DeleteDC " & DeleteDC
    End If
End Function


Function GetCursorHand() As StdPicture
    Set GetCursorHand = LoadResPicture(103, vbResCursor)
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hGDIBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OS_CreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" (ByVal crColor As Long) As Long
Private Declare Function OS_FillRect Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As t_RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type Bitmap '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Dim Draw As New Draw

Private m_hDC As Long
Private m_hBmpOld As Long
Private m_hBmp As Long
Private m_lWidth As Long
Private m_lheight As Long
Private m_FontHandle As Long

Public Property Get hdc() As Long
   hdc = m_hDC
End Property

Public Property Let Width(ByVal lW As Long)
   If lW > m_lWidth Then
      pCreate lW, m_lheight
   End If
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Height(ByVal lH As Long)
    If lH > m_lheight Then
        pCreate m_lWidth, lH
    End If
End Property

Public Property Get Height() As Long
    Height = m_lheight
End Property

Private Sub pCreate(ByVal lW As Long, ByVal lH As Long, Optional mCompatibleDC As Long)
Dim lhDC As Long
    pDestroy
    If mCompatibleDC = 0 Then
        lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        m_hDC = CreateCompatibleDC(lhDC)
        m_hBmp = CreateCompatibleBitmap(lhDC, lW, lH)
    Else
        m_hDC = CreateCompatibleDC(mCompatibleDC)
        m_hBmp = CreateCompatibleBitmap(mCompatibleDC, lW, lH)
    End If
    Draw.SetBkMode m_hDC, bkTransparent
    
    m_hBmpOld = SelectObject(m_hDC, m_hBmp)
    If m_hBmpOld = 0 Then
        pDestroy
    Else
        m_lWidth = lW
        m_lheight = lH
    End If
    If mCompatibleDC = 0 Then
        Call Draw.DeleteDC(lhDC)
    End If
    If m_FontHandle <> 0 Then
        SelectObject m_hDC, m_FontHandle
    End If
    Clear vbWhite
End Sub

Private Sub pDestroy()
   If Not m_hBmpOld = 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      Draw.DeleteObject m_hBmp
      m_hBmp = 0
   End If
   m_lWidth = 0
   m_lheight = 0
   If Not m_hDC = 0 Then
      Draw.DeleteDC m_hDC
      m_hDC = 0
   End If
End Sub

Private Sub Class_Initialize()
'    ContaClassDrawBuffer = ContaClassDrawBuffer + 1
'    ClassAdd "GDIBuffer"
End Sub

Private Sub Class_Terminate()
    'ContaClassDrawBuffer = ContaClassDrawBuffer - 1
    pDestroy
'    ClassClose "GDIBuffer"
End Sub

Public Sub Clear(Optional ByVal color As Long = vbWhite)
Dim hBrush As Long, lpRect As RECT
    lpRect.Right = m_lWidth
    lpRect.Bottom = m_lheight
    hBrush = OS_CreateSolidBrush(color)
    Call OS_FillRect(m_hDC, lpRect, hBrush)
    Call Draw.DeleteObject(hBrush)
End Sub

Public Sub SetSize(ByVal Width As Long, ByVal Height As Long, Optional mCompatibleDC As Long)
    pCreate Width, Height, mCompatibleDC
End Sub

Public Function SetFontHandle(ByVal FontHandle As Long) As Long
    m_FontHandle = FontHandle
    If m_FontHandle <> 0 And m_hDC <> 0 Then
        SetFontHandle = SelectObject(m_hDC, m_FontHandle)
    End If
End Function

Public Function GetPicture() As StdPicture
Dim mDib As New hJPEG_DIB
Dim hDib As Long, mBitMap As Long
Dim mPic As New hGDIFuncs

    'mPic.Create m_lWidth, m_lheight
    'If mPic.CreateDIB(m_hDC, m_lWidth, m_lheight, hDib) Then
        'Set GetPicture = hDib
    'End If
   ' Clear vbYellow
    If mDib.CreateDIB(m_hDC, m_lWidth, m_lheight, hDib) Then
        mDib.LoadPictureBlt m_hDC
    
        'MsgBox hDib
      '  Draw.StretchBlt hDib, 0, 0, m_lWidth, m_lheight, m_hDC, 0, 0, m_lWidth, m_lheight
      '  Draw.FillSolidRect hDib, CRect(0, 0, 100, 100), vbRed
        Set GetPicture = mPic.BitmapToPicture(hDib)
        'MsgBox GetPicture.handle
        
    End If
  '  df
    hDib = mPic.HBitmapFromDC(m_hDC, m_lWidth, m_lheight)
    Set GetPicture = mPic.BitmapToPicture(hDib)
    
    
    
    'If mDib.CreateFromPicture(GetPicture) Then
            
        'MsgBox hDib
      '  Draw.StretchBlt hDib, 0, 0, m_lWidth, m_lheight, m_hDC, 0, 0, m_lWidth, m_lheight
      '  Draw.FillSolidRect hDib, CRect(0, 0, 100, 100), vbRed
    '    Set GetPicture = mPic.BitmapToPicture(mDib.m_hDIb)
        'MsgBox GetPicture.handle
        
    'End If
    
    'mBitMap = mPic.HBitmapFromDC(m_hDC, m_lWidth, m_lheight)
    
    'Set GetPicture = LoadPicture("t:\aa.jpg")
End Function

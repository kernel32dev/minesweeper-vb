VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Campo Minado"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   1080
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1800
   End
   Begin VB.PictureBox Buff2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   615
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Buff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   15
      Left            =   2280
      Picture         =   "Form1.frx":014A
      Top             =   1440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   14
      Left            =   2280
      Picture         =   "Form1.frx":048C
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   13
      Left            =   2280
      Picture         =   "Form1.frx":07CE
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   12
      Left            =   1920
      Picture         =   "Form1.frx":0B10
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   11
      Left            =   1560
      Picture         =   "Form1.frx":0E52
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   10
      Left            =   1920
      Picture         =   "Form1.frx":1194
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   7
      Left            =   840
      Picture         =   "Form1.frx":14D6
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   8
      Left            =   1200
      Picture         =   "Form1.frx":1818
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   9
      Left            =   1560
      Picture         =   "Form1.frx":1B5A
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   6
      Left            =   2280
      Picture         =   "Form1.frx":1E9C
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   5
      Left            =   1920
      Picture         =   "Form1.frx":21DE
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   4
      Left            =   1560
      Picture         =   "Form1.frx":2520
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":2862
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   2
      Left            =   840
      Picture         =   "Form1.frx":2BA4
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   1
      Left            =   1200
      Picture         =   "Form1.frx":2EE6
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Bitmaps 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "Form1.frx":3228
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu Menu 
      Caption         =   "Novo Jogo"
      Index           =   0
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu OMenu 
         Caption         =   "Iniciante"
         Index           =   0
      End
      Begin VB.Menu OMenu 
         Caption         =   "Intermediário"
         Index           =   1
      End
      Begin VB.Menu OMenu 
         Caption         =   "Expert"
         Index           =   2
      End
      Begin VB.Menu OMenu 
         Caption         =   "Customizado..."
         Index           =   3
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Outros"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu AMenu 
         Caption         =   "Cheats"
         Index           =   0
         Begin VB.Menu CMenu 
            Caption         =   "Undo"
            Index           =   0
         End
         Begin VB.Menu CMenu 
            Caption         =   "Raio X"
            Index           =   1
         End
         Begin VB.Menu CMenu 
            Caption         =   "Pincel de Bombas"
            Index           =   2
         End
         Begin VB.Menu CMenu 
            Caption         =   "Recalculadora"
            Index           =   3
         End
      End
      Begin VB.Menu AMenu 
         Caption         =   "Pontuações"
         Index           =   1
      End
      Begin VB.Menu AMenu 
         Caption         =   "Redimensionar"
         Index           =   2
         Begin VB.Menu RMenu 
            Caption         =   "x1"
            Index           =   0
         End
         Begin VB.Menu RMenu 
            Caption         =   "x2"
            Index           =   1
         End
         Begin VB.Menu RMenu 
            Caption         =   "x3"
            Index           =   2
         End
         Begin VB.Menu RMenu 
            Caption         =   "x4"
            Index           =   3
         End
         Begin VB.Menu RMenu 
            Caption         =   "x5"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function AdjustWindowRect Lib "user32" (Rectangle As Long, ByVal Style As Long, ByVal hasMenu As Boolean) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal Index As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, Rectangle As Long) As Long
Private Declare Function SetWindowRect Lib "user32" (ByVal hwnd As Long, Rectangle As Long) As Long

Dim Draw As New Draw

Dim BDraw As Boolean

Dim ActuallyGenerated As Boolean

Private Type Move
X As Long
Y As Long
Action As Boolean
End Type

Private Type Estatistics
Clicks As Long
Mistakes As Long
SizeX As Long
SizeY As Long
BombCount As Long
Time As Double
Cheater As Boolean
End Type

Dim Moves() As Move
Dim MaxM As Long

Dim PlayerName As String

'Dim Buff As New hGDIBuffer
'Dim Buff2 As New hGDIBuffer

'Const DefalutSizeX As Long = 16
'Const DefalutSizeY As Long = 16

Dim JustDoubleClicked As Boolean

Dim ESS As Estatistics

Dim SizeX As Long
Dim SizeY As Long
Dim BombCount As Long

Dim h As Long
Dim w As Long

Const Spacing As Long = 2
Const AppName As String = "Campo Minado GJ"

Private Type Tile
Marked As Long
Revealed As Boolean
Value As Long
Flag As Boolean
End Type

Dim LastGrid() As Tile
Dim UndoFlag As Boolean

Dim Grid() As Tile
Dim Tiles(-4 To 11) As Long

Dim ForceB As Boolean
Dim CX As Long
Dim CY As Long

Dim Off As Long
Dim OffB As Boolean

Dim EndGame As Boolean
Dim BombX As Long
Dim BombY As Long

Private Sub AMenu_Click(Index As Integer)
If Index = 1 Then
Form2.Show
End If
End Sub

Private Sub CMenu_Click(Index As Integer)
CMenu(Index).Checked = Not CMenu(Index).Checked
If Index = 1 Then FlagAll
If CMenu(Index).Checked Then ESS.Cheater = True
End Sub

Private Sub Form_DblClick()
If Not (CX > 0 And CY > 0 And CX < SizeX And CY < SizeY) Or EndGame Then
JustDoubleClicked = True
GenerateMap False
FlagAll
EndGame = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 8 And CMenu(0).Checked Then UndoMove
If KeyCode = 82 Then FlagAll
End Sub

Private Sub Form_Load()
'OMenu_Click 0
Dim X As Long
Dim Y As Long
Dim w As Long
X = GetSetting(AppName, "Data", "SizeX", 8)
Y = GetSetting(AppName, "Data", "SizeY", 8)
w = GetSetting(AppName, "Data", "BombCount", 10)

If X * Y < w Then w = X * Y
SizeX = X
SizeY = Y
BombCount = w
For w = 0 To CMenu.Count - 1
CMenu(w).Checked = GetSetting(AppName, "Data", "Cheats" & CStr(w), "0")
Next
EndGame = (GetSetting(AppName, "Data", "GameState", 0) = 1)
BombX = GetSetting(AppName, "Data", "BoomX", 0)
BombY = GetSetting(AppName, "Data", "BoomY", 0)
ReDim Grid(SizeX - 1, SizeY - 1)
ReDim LastGrid(SizeX - 1, SizeY - 1)
Dim S As String
S = GetSetting(AppName, "Data", "Map", "R")
If S = "R" Then
GenerateMap False
Else
SetMapString S
FlagAll
End If
StringToESS GetSetting(AppName, "Data", "Score", "")
S = GetSetting(AppName, "Data", "LastMap", "R")
If Not S = "R" Then
UndoFlag = True
SetMapString S, True
FlagAll
Else
UndoFlag = False
End If
'Tiles(-1) = RGB(255, 255, 255)
'Tiles(0) = RGB(128, 128, 128)
'Tiles(1) = RGB(0, 0, 255)
'Tiles(2) = RGB(0, 255, 0)
'Tiles(3) = RGB(255, 0, 0)
'Tiles(4) = RGB(128, 0, 0)
'Tiles(5) = RGB(200, 0, 200)
Tiles(-4) = Bitmaps(14)
Tiles(-3) = Bitmaps(12)
Tiles(-2) = Bitmaps(11)
Tiles(-1) = Bitmaps(0)
Tiles(0) = Bitmaps(1)
Tiles(1) = Bitmaps(2)
Tiles(2) = Bitmaps(3)
Tiles(3) = Bitmaps(4)
Tiles(4) = Bitmaps(5)
Tiles(5) = Bitmaps(6)
Tiles(6) = Bitmaps(7)
Tiles(7) = Bitmaps(8)
Tiles(8) = Bitmaps(9)
Tiles(9) = Bitmaps(10)
Tiles(10) = Bitmaps(13)
Tiles(11) = Bitmaps(15)
UpdateFlagCount
End Sub

Function ESSToString() As String
Dim Ls(6) As Long
Dim S As String
Ls(0) = ESS.BombCount
Ls(1) = ESS.Cheater
Ls(2) = ESS.Clicks
Ls(3) = ESS.Mistakes
Ls(4) = ESS.SizeX
Ls(5) = ESS.SizeY
Ls(6) = ESS.Time
S = Ls(0) & vbNewLine
S = S & Ls(1) & vbNewLine
S = S & Ls(2) & vbNewLine
S = S & Ls(3) & vbNewLine
S = S & Ls(4) & vbNewLine
S = S & Ls(5) & vbNewLine
S = S & Ls(6) & vbNewLine
S = S & PlayerName
ESSToString = S
End Function

Function StringToESS(S As String)
Dim C() As String
If S = "" Then Exit Function
C = Split(S, vbNewLine)
ESS.BombCount = Val(C(0))
ESS.Cheater = Val(C(1))
ESS.Clicks = Val(C(2))
ESS.Mistakes = Val(C(3))
ESS.SizeX = Val(C(4))
ESS.SizeY = Val(C(5))
ESS.Time = Val(C(6))
PlayerName = C(7)
End Function


Function GetMapString(Optional Last As Boolean) As String
Dim X As Long
Dim Y As Long
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
    'A = Revelado
    'B = Não Marcado
    'C = Marcado 1
    'D = Marcado 2
    'E = Não Marcado c/Bomba
    'F = Marcado 1 c/Bomba
    'G = Marcado 2 c/Bomba
        If Last Then
            On Error GoTo Err
            With LastGrid(X, Y)
                If .Revealed = True Then
                GetMapString = GetMapString & "A"
                ElseIf .Marked = 0 And .Value <> 9 Then
                GetMapString = GetMapString & "B"
                ElseIf .Marked = 1 And .Value <> 9 Then
                GetMapString = GetMapString & "C"
                ElseIf .Marked = 2 And .Value <> 9 Then
                GetMapString = GetMapString & "D"
                ElseIf .Marked = 0 And .Value = 9 Then
                GetMapString = GetMapString & "E"
                ElseIf .Marked = 1 And .Value = 9 Then
                GetMapString = GetMapString & "F"
                ElseIf .Marked = 2 And .Value = 9 Then
                GetMapString = GetMapString & "G"
                End If
            End With
        Else
            With Grid(X, Y)
                If .Revealed = True Then
                GetMapString = GetMapString & "A"
                ElseIf .Marked = 0 And .Value <> 9 Then
                GetMapString = GetMapString & "B"
                ElseIf .Marked = 1 And .Value <> 9 Then
                GetMapString = GetMapString & "C"
                ElseIf .Marked = 2 And .Value <> 9 Then
                GetMapString = GetMapString & "D"
                ElseIf .Marked = 0 And .Value = 9 Then
                GetMapString = GetMapString & "E"
                ElseIf .Marked = 1 And .Value = 9 Then
                GetMapString = GetMapString & "F"
                ElseIf .Marked = 2 And .Value = 9 Then
                GetMapString = GetMapString & "G"
                End If
            End With
        End If
    Next
Next
Err:
End Function

Function SetMapString(S As String, Optional Last As Boolean)
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim SS As String
For Z = 1 To Len(S)
    SS = Mid$(S, Z, 1)
    If Last Then
        With LastGrid(X, Y)
            .Flag = True
            .Marked = 0
            .Revealed = False
            .Value = 0
            If SS = "A" Then
            .Revealed = True
            ElseIf SS = "B" Then
            ElseIf SS = "C" Then
            .Marked = 1
            ElseIf SS = "D" Then
            .Marked = 2
            ElseIf SS = "E" Then
            .Value = 9
            ElseIf SS = "F" Then
            .Marked = 1
            .Value = 9
            ElseIf SS = "G" Then
            .Marked = 2
            .Value = 9
            End If
        End With
    Else
        With Grid(X, Y)
            .Flag = True
            .Marked = 0
            .Revealed = False
            .Value = 0
            If SS = "A" Then
            .Revealed = True
            ElseIf SS = "B" Then
            ElseIf SS = "C" Then
            .Marked = 1
            ElseIf SS = "D" Then
            .Marked = 2
            ElseIf SS = "E" Then
            .Value = 9
            ElseIf SS = "F" Then
            .Marked = 1
            .Value = 9
            ElseIf SS = "G" Then
            .Marked = 2
            .Value = 9
            End If
        End With
    End If
    X = X + 1
    If X = SizeX Then
    X = 0
    Y = Y + 1
    If Y = SizeY Then Exit For
    End If
Next
Values Last
End Function

Function Values(Optional Last As Boolean)
On Error Resume Next
Dim X As Long
Dim Y As Long
Dim Z As Long
    If Not Last Then
        For Y = 0 To SizeY - 1
            For X = 0 To SizeX - 1
                If Grid(X, Y).Value <> 9 Then
                Z = 0
                Z = Z - (Grid(X, Y + 1).Value = 9)
                Z = Z - (Grid(X, Y - 1).Value = 9)
                Z = Z - (Grid(X + 1, Y).Value = 9)
                Z = Z - (Grid(X - 1, Y).Value = 9)
                Z = Z - (Grid(X + 1, Y + 1).Value = 9)
                Z = Z - (Grid(X - 1, Y - 1).Value = 9)
                Z = Z - (Grid(X + 1, Y - 1).Value = 9)
                Z = Z - (Grid(X - 1, Y + 1).Value = 9)
                Z = Z + (Grid(X, Y + 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X, Y - 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X + 1, Y).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X - 1, Y).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X + 1, Y + 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X - 1, Y - 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X + 1, Y - 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (Grid(X - 1, Y + 1).Marked = 1 And CMenu(3).Checked)
                If Z < 0 Then Z = 0
                Grid(X, Y).Flag = (Grid(X, Y).Value <> Z)
                Grid(X, Y).Value = Z
                End If
            Next
        Next
    Else
        For Y = 0 To SizeY - 1
            For X = 0 To SizeX - 1
                If LastGrid(X, Y).Value <> 9 Then
                Z = 0
                Z = Z - (LastGrid(X, Y + 1).Value = 9)
                Z = Z - (LastGrid(X, Y - 1).Value = 9)
                Z = Z - (LastGrid(X + 1, Y).Value = 9)
                Z = Z - (LastGrid(X - 1, Y).Value = 9)
                Z = Z - (LastGrid(X + 1, Y + 1).Value = 9)
                Z = Z - (LastGrid(X - 1, Y - 1).Value = 9)
                Z = Z - (LastGrid(X + 1, Y - 1).Value = 9)
                Z = Z - (LastGrid(X - 1, Y + 1).Value = 9)
                Z = Z + (LastGrid(X, Y + 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X, Y - 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X + 1, Y).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X - 1, Y).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X + 1, Y + 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X - 1, Y - 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X + 1, Y - 1).Marked = 1 And CMenu(3).Checked)
                Z = Z + (LastGrid(X - 1, Y + 1).Marked = 1 And CMenu(3).Checked)
                If Z < 0 Then Z = 0
                LastGrid(X, Y).Value = Z
                End If
            Next
        Next
    End If
End Function

Function CheckWin() 'Also Counts Flags
Dim X As Long
Dim Y As Long
Dim FlagCount As Long
If EndGame Then Exit Function
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
        If Grid(X, Y).Value <> 9 And Not Grid(X, Y).Revealed Then Exit Function
    Next
Next
EndGame = True
BombX = -1
BombY = -1
FlagAll
Timer2.Enabled = True

End Function

Function UpdateFlagCount()
Dim NewCaption As String
Dim X As Long, Y As Long
Dim FlagCount As Long
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
        If Grid(X, Y).Marked = 1 Then
        FlagCount = FlagCount + 1
        End If
    Next
Next
If FlagCount <= BombCount Then
NewCaption = AppName & " - " & BombCount - FlagCount & " Bombas"
Else
NewCaption = AppName & " - " & FlagCount - BombCount & " Bandeiras"
End If
If NewCaption <> Caption Then
Caption = NewCaption
End If

End Function

Function FlagAll()
Dim X As Long
Dim Y As Long
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
    Grid(X, Y).Flag = True
    Next
Next
End Function

Function GenerateMap(ByVal ActuallyDoIt As Boolean, Optional EPX As Long = -3, Optional EPY As Long = -3)
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim B As Boolean
If Not ActuallyDoIt Then
    For Y = 0 To SizeY - 1
        For X = 0 To SizeX - 1
        Grid(X, Y).Marked = 0
        Grid(X, Y).Flag = True
        Grid(X, Y).Value = 0
        Grid(X, Y).Revealed = False
        Next
    Next
    ResetESS
    UpdateFlagCount
    ActuallyGenerated = False
Exit Function
End If
ActuallyGenerated = True
'If BombCount_ = -1 Then BombCount_ = BombCount
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
    Grid(X, Y).Marked = 0
    Grid(X, Y).Flag = True
    Grid(X, Y).Value = 0
    Grid(X, Y).Revealed = False
    Next
Next
Randomize Timer
For Z = 1 To BombCount '_
    Do
    X = Rnd * (SizeX - 1)
    Y = Rnd * (SizeY - 1)
    B = Not (X >= EPX - 1 And X <= EPX + 1 And Y >= EPY - 1 And Y <= EPY + 1)
    B = B And (Grid(X, Y).Value <> 9)
    Loop Until B
    Grid(X, Y).Value = 9
Next
Values
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If JustDoubleClicked Then Exit Sub
Dim XX As Long
Dim YY As Long
If OffB Then
XX = X
YY = Y - Off
Else
XX = X - Off
YY = Y
End If
On Error Resume Next
Grid(CX, CY).Flag = True
CX = (XX * SizeX / w) - 0.5
CY = (YY * SizeY / h) - 0.5
Grid(CX, CY).Flag = True
If CX >= 0 And CY >= 0 And CX < SizeX And CY < SizeY Then
MouseDown (Button), (CX), (CY)
ElseIf Button = 2 Then
PopupMenu Menu
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If JustDoubleClicked Then Exit Sub
Dim XX As Long
Dim YY As Long
If OffB Then
XX = X
YY = Y - Off
Else
XX = X - Off
YY = Y
End If
On Error Resume Next
Grid(CX, CY).Flag = True
CX = (XX * SizeX / w) - 0.5
CY = (YY * SizeY / h) - 0.5
Grid(CX, CY).Flag = True
Menu(0).Visible = Y < 8
Menu(1).Visible = Menu(0).Visible
If CX >= 0 And CY >= 0 And CX < SizeX And CY < SizeY Then MouseMove (Button), (CX), (CY)
UpdateFlagCount
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If JustDoubleClicked Then
JustDoubleClicked = False
Exit Sub
End If
Dim XX As Long
Dim YY As Long
If OffB Then
XX = X
YY = Y - Off
Else
XX = X - Off
YY = Y
End If
On Error Resume Next
Grid(CX, CY).Flag = True
CX = (XX * SizeX / w) - 0.5
CY = (YY * SizeY / h) - 0.5
Grid(CX, CY).Flag = True
If CX >= 0 And CY >= 0 And CX < SizeX And CY < SizeY Then MouseUp (Button), (CX), (CY)
End Sub

Private Sub Form_Resize()
UpdateSize
End Sub

Function UpdateSize()
If Form1.ScaleHeight = 0 Or Form1.ScaleWidth = 0 Then Exit Function
    Buff2.Width = Form1.ScaleWidth
    Buff2.Height = Form1.ScaleHeight
    If Form1.ScaleWidth / Form1.ScaleHeight > SizeX / SizeY Then
    Buff.Width = Form1.ScaleHeight * (SizeX / SizeY)
    Buff.Height = Form1.ScaleHeight
    OffB = False
    Else
    Buff.Width = Form1.ScaleWidth
    Buff.Height = Form1.ScaleWidth * (SizeY / SizeX)
    OffB = True
    End If
    If OffB Then
    Off = Abs(Buff2.Height - Buff.Height) / 2
    Else
    Off = Abs(Buff2.Width - Buff.Width) / 2
    End If
    w = Buff.Width ' + Spacing * SizeX
    h = Buff.Height ' + Spacing * SizeY
    FlagAll
End Function

'Private Sub Form_Resize()
'Buff2.Width = Form1.ScaleWidth
'Buff2.Height = Form1.ScaleHeight
'If Form1.ScaleWidth > Form1.ScaleHeight Then
'Buff.Width = Form1.ScaleHeight
'Buff.Height = Form1.ScaleHeight
'OffB = False
'Else
'Buff.Width = Form1.ScaleWidth
'Buff.Height = Form1.ScaleWidth
'OffB = True
'End If
'Off = Abs(Form1.ScaleWidth - Form1.ScaleHeight) / 2
'W = Buff.Width ' + Spacing * SizeX
'h = Buff.Height ' + Spacing * SizeY
'FlagAll
'End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting AppName, "Data", "SizeX", CStr(SizeX)
SaveSetting AppName, "Data", "SizeY", CStr(SizeY)
SaveSetting AppName, "Data", "BombCount", CStr(BombCount)
SaveSetting AppName, "Data", "GameState", CStr(-EndGame + 0)
SaveSetting AppName, "Data", "BoomX", CStr(BombX)
Dim L As Long
For L = 0 To CMenu.Count - 1
If CMenu(L).Checked Then
SaveSetting AppName, "Data", "Cheats" & CStr(L), "-1"
Else
SaveSetting AppName, "Data", "Cheats" & CStr(L), "0"
End If
Next
SaveSetting AppName, "Data", "Map", GetMapString
SaveSetting AppName, "Data", "Map", GetMapString
If UndoFlag Then
SaveSetting AppName, "Data", "LastMap", GetMapString(True)
Else
SaveSetting AppName, "Data", "LastMap", "R"
End If
SaveSetting AppName, "Data", "Score", ESSToString
On Error Resume Next
Unload Form2
End Sub

Private Sub OMenu_Click(Index As Integer)
If Index = 3 Then
NewGameForm.Show
NewGameForm.Text1(0) = SizeY
NewGameForm.Text1(1) = SizeX
NewGameForm.Text1(2) = BombCount
NewGameForm.Text1(0).Tag = SizeY
NewGameForm.Text1(1).Tag = SizeX
NewGameForm.Text1(2).Tag = BombCount
Me.Enabled = False
Else
    Select Case Index
    Case 0
    NewGame 8, 8, 10
    Case 1
    NewGame 16, 16, 40
    Case 2
    NewGame 24, 24, 100
    End Select
    
End If
End Sub

Private Sub RMenu_Click(Index As Integer)
Dim Zoom As Long
Zoom = Index + 1
Dim R(3) As Long

GetWindowRect hwnd, R(0)

R(2) = R(0) + Zoom * SizeX * 16
R(3) = R(1) + Zoom * SizeY * 16

AdjustWindowRect R(0), GetWindowLongA(hwnd, -16), False

If R(0) < 0 Then
R(0) = 0
R(2) = R(2) - R(0)
End If
If R(1) < 0 Then
R(1) = 0
R(3) = R(3) - R(1)
End If


MoveWindow hwnd, R(0), R(1), R(2) - R(0), R(3) - R(1), True

End Sub

Private Sub Timer1_Timer()
'Dim X As Long
'Dim Y As Long
'X = (SizeX - 1) * Rnd
'Y = (SizeY - 1) * Rnd
'Grid(X, Y).Flag = True
Redraw
End Sub

Function SaveMove()
UndoFlag = True
LastGrid = Grid
End Function

Function UndoMove()
If UndoFlag Then
UndoFlag = False
Grid = LastGrid
FlagAll
EndGame = False
End If
End Function

'Function AddMove(X As Long, Y As Long, Action As Boolean)
'MaxM = MaxM + 1
'Moves(MaxM).X = X
'Moves(MaxM).Y = Y
'Moves(MaxM).Action = Action
'ReDim Preserve Moves(MaxM)
'End Function

'Function UndoMove()
'If MaxM >= 0 Then
'EndGame = False
'    With Moves(MaxM)
'        If .Action Then
'        Grid(.X, .Y).Marked = Grid(.X, .Y).Marked - 1
'            If Grid(.X, .Y).Marked = -1 Then
'            Grid(.X, .Y).Marked = 2
'            End If
'        Else
'        Grid(.X, .Y).Revealed = False
'        End If
'    End With
'MaxM = MaxM - 1
'End If
'End Function

Function NewGame(SizeX_ As Long, SizeY_ As Long, BombCount_ As Long)
MaxM = -1
SizeX = SizeX_
SizeY = SizeY_
ReDim Grid(SizeX - 1, SizeY - 1)
ReDim LastGrid(SizeX - 1, SizeY - 1)
BombCount = BombCount_
ForceB = False
UndoFlag = False
EndGame = False
Dim L As Long
ESS.Cheater = False
For L = 0 To CMenu.Count - 1
ESS.Cheater = CMenu(L).Checked Or ESS.Cheater
Next
GenerateMap False
UpdateSize
ResetESS
FlagAll
End Function

Function ResetESS()
ESS.BombCount = BombCount
ESS.Clicks = 0
ESS.Mistakes = 0
ESS.SizeX = SizeX
ESS.SizeY = SizeY
Timer3.Enabled = False
Timer3.Enabled = True
ESS.Time = Timer
End Function

Function Reveal(X As Long, Y As Long)
On Error Resume Next
Dim B As Boolean
B = (Not Grid(X, Y).Revealed And Grid(X, Y).Marked = 0)
If B Then
    Grid(X, Y).Revealed = True
    Grid(X, Y).Flag = True
    B = (Grid(X, Y).Value = 0)
    If B Then
    Reveal X, Y + 1
    Reveal X, Y - 1
    Reveal X + 1, Y
    Reveal X - 1, Y
    Reveal X + 1, Y + 1
    Reveal X - 1, Y - 1
    Reveal X + 1, Y - 1
    Reveal X - 1, Y + 1
    End If
End If
End Function

Function GameOver()
EndGame = True
ForceB = False
Dim X As Long
Dim Y As Long
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
    Grid(X, Y).Flag = True
    Next
Next
End Function

Function MouseDown(Button As Long, X As Long, Y As Long)
ForceB = Not EndGame
If CMenu(2).Checked And Button = 4 And ActuallyGenerated Then
BDraw = (Grid(X, Y).Value <> 9)
MouseMove 4, X, Y
End If
End Function

Function MouseMove(Button As Long, X As Long, Y As Long)
'Form1.Caption = CX & " " & CY
If CMenu(2).Checked And Button = 4 And ActuallyGenerated Then
    If BDraw Then
    Grid(X, Y).Value = 9
    Else
    Grid(X, Y).Value = 0
    End If
    Values
End If
End Function

Function MouseUp(Button As Long, X As Long, Y As Long)
ForceB = False
If Not ActuallyGenerated Then GenerateMap True, X, Y
If CMenu(2).Checked Then Exit Function
If Not EndGame Then ESS.Clicks = ESS.Clicks + 1
If Button = 1 And Not EndGame Then
    SaveMove
    If Grid(X, Y).Value = 9 Then
    GameOver
    ESS.Mistakes = ESS.Mistakes + 1
    BombX = X
    BombY = Y
    Else
    Reveal X, Y
        If CMenu(3).Checked Then
        Values
        FlagAll
        End If
    CheckWin
    End If
    ''AddMove X, Y, False
ElseIf Button = 2 And Not Grid(X, Y).Revealed And Not EndGame Then
SaveMove
Grid(X, Y).Marked = Grid(X, Y).Marked + 1
If Grid(X, Y).Marked = 3 Then Grid(X, Y).Marked = 0
If CMenu(3).Checked Then Values
'AddMove X, Y, True
End If
End Function

Function Redraw()
Dim X As Long
Dim Y As Long
Dim Z As Long
Dim R As t_RECT
For Y = 0 To SizeY - 1
    For X = 0 To SizeX - 1
        With Grid(X, Y)
            If .Flag Then
                .Flag = False
                
                If ForceB And CX = X And CY = Y And Not .Revealed Then
                Z = 0
                ElseIf X = BombX And Y = BombY And EndGame Then
                Z = 10
                Else
                    If .Revealed Or (EndGame And Grid(X, Y).Value = 9 And Not Grid(X, Y).Marked = 1) Or (CMenu(1).Checked) Then
                    Z = .Value
                    Else
                    Z = -1 - .Marked
                    If EndGame And Z = -2 And Grid(X, Y).Value <> 9 Then Z = -4
                    If EndGame And Z = -2 And Grid(X, Y).Value = 9 Then Z = 11
                    End If
                    If EndGame And BombX = -1 And Grid(X, Y).Value = 9 Then Z = 11
                End If
                
                R.Top = Y * h / SizeY ' + (Y + 1) * Spacing
                R.Left = X * w / SizeX ' + (X + 1) * Spacing
                R.Right = (X + 1) * w / SizeX ' + (X + 1) * Spacing
                R.Bottom = (Y + 1) * h / SizeY ' + (Y + 1) * Spacing

                Draw.DrawBitmapEx Buff.hdc, Tiles(Z), R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, True
                'Draw.FillSolidRect Buff.hDC, R, Tiles(Z)
            End If
        End With
    Next
Next
R.Top = 0
R.Left = 0
R.Right = Buff2.Width
R.Bottom = Buff2.Height
Draw.FillSolidRect Buff2.hdc, R, Form1.BackColor
If OffB Then
Draw.BitBlt Buff2.hdc, 0, Off, w, h, Buff.hdc, 0, 0
Else
Draw.BitBlt Buff2.hdc, Off, 0, w, h, Buff.hdc, 0, 0
End If
Draw.BitBlt Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, Buff2.hdc, 0, 0
End Function

Private Sub Timer2_Timer()
ESS.Time = Round(Timer - ESS.Time, 2)
Timer2.Enabled = False
Dim S As String
Dim L As Long
S = InputBox("Registrar Vitória?" & vbNewLine & "Qual seu nome?", "Campo Minado", PlayerName)
If Not S = "" Then
PlayerName = S
    Do
    L = L + 1
    S = GetSetting(AppName, "Scores", "Score" & CStr(L), "Free")
    Loop Until S = "Free"
SaveSetting AppName, "Scores", "Score" & CStr(L), ESSToString
End If
End Sub

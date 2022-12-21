VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Pontuações Arquivadas"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   739
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   1560
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const AppName As String = "Campo Minado GJ"

Dim Kill As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
If KeyCode = 46 Then
Kill = List1.ListIndex + 1
Reload
End If
End Sub

Private Sub Form_Load()
Kill = -1
Reload
End Sub

Private Sub Form_Resize()
List1.Width = Form1.ScaleWidth
List1.Height = Form1.ScaleHeight
On Error Resume Next
Form1.Width = List1.Width * Screen.TwipsPerPixelX
Form1.Height = List1.Height * Screen.TwipsPerPixelY
End Sub

'Ls(0) = ESS.BombCount
'Ls(1) = ESS.Cheater
'Ls(2) = ESS.Clicks
'Ls(3) = ESS.Mistakes
'Ls(4) = ESS.SizeX
'Ls(5) = ESS.SizeY
'Ls(6) = ESS.Time

Function Reload()
Dim S As String
Dim T As String
Dim B As Boolean
Dim L As Long
Dim X As Long
Dim C() As String
List1.Clear
Do
L = L + 1
S = GetSetting(AppName, "Scores", "Score" & CStr(L), "FreeD")
    
    If X = Kill Then
    SaveSetting AppName, "Scores", "Score" & CStr(L - 1), "Free"
    Kill = -1
    ElseIf Not S = "Free" And Not S = "FreeD" Then
    X = X + 1
    C = Split(S, vbNewLine)
    T = ""
    AddToString T, C(7) & ":", 10
    AddToString T, "Tempo: " & C(6) \ 60 & ":" & C(6) Mod 60
    AddToString T, " Tamanho: " & C(4) & "x" & C(5)
    AddToString T, " Bombs: " & C(0)
    AddToString T, " Clicks: " & C(2)
    If C(1) Then AddToString T, " Erros:" & C(3) & " [C]"
        List1.AddItem T
        'If C(1) Then
        'List1.AddItem C(7) & ":  Tempo:" & C(6) & " Tamanho:" & C(4) & "x" & C(5) & " Bombs:" & C(0) & " Clicks:" & C(2)
        'Else
        'List1.AddItem C(7) & ":  Tempo:" & C(6) & " Tamanho:" & C(4) & "x" & C(5) & " Bombs:" & C(0) & " Clicks:" & C(2) & " Erros:" & C(3) & " [U]"
        'End If
    End If
Loop Until S = "FreeD"
Timer1.Enabled = Not Timer1.Enabled
End Function

Function AddToString(Str As String, ToAdd As String, Optional RequiredSize As Long = 15)
On Error GoTo Err
Str = Str & ToAdd & Space$(RequiredSize - Len(ToAdd))
Exit Function
Err:
Str = Str & ToAdd
End Function

Private Sub Timer1_Timer()
Reload
End Sub

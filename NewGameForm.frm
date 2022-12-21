VERSION 5.00
Begin VB.Form NewGameForm 
   Caption         =   "Novo Jogo"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NewGameForm.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Novo Jogo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Text            =   "30"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Text            =   "16"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Text            =   "16"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "                      "
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Bombas:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Largura:"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Altura:"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "NewGameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim L As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
L = L + 1
If L >= 10 Then
Label2 = "(Versão NoDep)"
Label2.Visible = True
End If
End Sub

Private Sub Command1_Click()
Form1.NewGame Val(Text1(1)), Val(Text1(0)), Val(Text1(2))
Form1.UpdateSize
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub

Private Sub Text1_Change(Index As Integer)
If IsNumeric(Text1(Index)) Or Text1(Index) = "" Then
Text1(Index).Tag = Text1(Index)
Else
Text1(Index) = Text1(Index).Tag
End If
Command1.Enabled = (Val(Text1(0)) > 0 And Val(Text1(1)) > 0 And Val(Text1(2)) > 0 And Val(Text1(0)) * Val(Text1(1)) >= Val(Text1(2)))
'Command1.Enabled = (Text1(1) > 0 And Text1(2) > 0 And Text1(1) * Text1(1) >= Text1(2))
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
End Sub

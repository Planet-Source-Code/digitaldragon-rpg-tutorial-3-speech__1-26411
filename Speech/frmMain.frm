VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RPG Basics by DigitalDragon"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Ekran 
      AutoRedraw      =   -1  'True
      Height          =   7200
      Left            =   120
      MousePointer    =   1  'Arrow
      ScaleHeight     =   7140
      ScaleWidth      =   7140
      TabIndex        =   3
      Top             =   360
      Width           =   7200
      Begin VB.PictureBox PicTXT 
         BackColor       =   &H00404040&
         Height          =   2895
         Left            =   60
         ScaleHeight     =   2835
         ScaleWidth      =   6975
         TabIndex        =   4
         Top             =   4200
         Visible         =   0   'False
         Width           =   7030
         Begin VB.PictureBox Holder 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   6135
            Left            =   0
            ScaleHeight     =   6135
            ScaleWidth      =   6735
            TabIndex        =   6
            Top             =   0
            Width           =   6735
            Begin VB.Label Pitanje 
               AutoSize        =   -1  'True
               BackColor       =   &H00404040&
               Caption         =   ":::"
               ForeColor       =   &H000080FF&
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   0
               Width           =   6420
               WordWrap        =   -1  'True
            End
            Begin VB.Label Choice 
               AutoSize        =   -1  'True
               BackColor       =   &H00404040&
               Caption         =   ">>>"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   360
               Width           =   6450
               WordWrap        =   -1  'True
            End
         End
         Begin VB.VScrollBar TextScroll 
            Height          =   2830
            LargeChange     =   3
            Left            =   6720
            Max             =   0
            TabIndex        =   5
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      LargeChange     =   3
      Left            =   120
      Max             =   0
      TabIndex        =   2
      Top             =   7560
      Width           =   7215
   End
   Begin VB.VScrollBar VS 
      Height          =   7215
      LargeChange     =   3
      Left            =   7320
      Max             =   0
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Choice_Click(Index As Integer)
RedirectSpeech Speech(CSpeech).Redirect(Index), TalkingTo
End Sub

Private Sub Choice_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
For i = 0 To Choice.Count - 1
    Choice(i).ForeColor = vbWhite
Next
Choice(Index).ForeColor = vbYellow
End Sub

Private Sub Ekran_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Ekran.MousePointer = 1
If IsCursorOverNPC(x / 15, y / 15) Then Ekran.MousePointer = 14
End Sub

Private Sub Form_Load()
Running = True
Me.Show
DoEvents
Init
Soldier.Init
FPSString = "FPS: ..."
Map.Init
Map.LoadMap App.Path & "\Collision.map"
HS.Max = Map.LenX - Int((frmMain.Ekran.Width / 15) / Map.TileWidth)
VS.Max = Map.LenY - Int((frmMain.Ekran.Height / 15) / Map.TileHeight)
Ekran.SetFocus
MainLoop
End Sub

Private Sub Ekran_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If PicTXT.Visible = True Then Exit Sub
If IsCursorOverNPC(x / 15, y / 15) Then
    Dim t1 As Integer
    t1 = GetNPCArrayIndexEx(x / 15, y / 15)
    Soldier.SetDestination NPC(t1).PosX - 1, NPC(t1).PosY - 1
    Soldier.SpeechDest = True
Else
    Soldier.SetDestination x / 15 + HS.Value * Map.TileWidth, y / 15 + VS.Value * Map.TileHeight
End If
End Sub

Sub MainLoop()
Dim MyFont As New StdFont
Map.DrawMap GetSrcRect
Do While Running
    Map.BltToBackBuffer 0, 0, GetSrcRect
    Soldier.UpdateAnimation
    Soldier.UpdateMove
    Soldier.Draw BackBuffer
    DrawNPC
    RX.getFPSString
    Label1.Caption = RX.FPSString
    FlipToDC
    DoEvents
Loop
Destroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
Running = False
Destroy
End Sub

Private Sub Holder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
For i = 0 To Choice.Count - 1
Choice(i).ForeColor = vbWhite
Next
Ekran.MousePointer = 1
End Sub

Private Sub HS_Change()
Map.ChangeView HS.Value, VS.Value
End Sub

Private Sub PicTXT_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
For i = 0 To Choice.Count - 1
Choice(i).ForeColor = vbWhite
Next
Ekran.MousePointer = 1
End Sub

Private Sub TextScroll_Change()
Holder.Top = (TextScroll.Value * -1) * 400
End Sub

Private Sub VS_Change()
Map.ChangeView HS.Value, VS.Value
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoldier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//Sprite stuff
Private PosX As Single, PosY As Single, Speed As Single
Private destX As Integer, destY As Integer, DestSet As Boolean
Private SpriteWidth As Integer, SpriteHeight As Integer
Public MoveX As Integer, MoveY As Single
'//Display stuff
Private CurrentKeyFrame As Integer
Private dds As DirectDrawSurface7
Private CRect As RECT
Private StartTick As Long, dx As New DirectX7
Private UnitType As Integer, CTick As Long
Private ViseManjeX As Integer, ViseManjeY As Integer
Public SpeechDest As Boolean

Public Sub Init()
SpriteWidth = 32
SpriteHeight = 32
CurrentKeyFrame = 1
UnitType = 1
Speed = 40
CRect.Bottom = 32
CRect.Right = 32
End Sub

Public Sub UpdateAI()
'//Tu se kasnije doda kod
End Sub

Public Sub UpdateAnimation()
If DestSet Then
    If dx.TickCount >= StartTick + 1000 / 5 Then
        StartTick = dx.TickCount
        If Not CurrentKeyFrame = 5 Then CurrentKeyFrame = CurrentKeyFrame + 1 Else CurrentKeyFrame = 2
        CRect.Left = CurrentKeyFrame * SpriteWidth - SpriteWidth
        CRect.Right = CRect.Left + SpriteWidth
    End If
Else: CurrentKeyFrame = 1
End If
End Sub

Public Sub Turn(Side As Integer)
CurrentKeyFrame = 1
CRect.Left = 0
CRect.Right = CRect.Left + SpriteWidth
CRect.Top = Side * SpriteHeight
CRect.Bottom = CRect.Top + SpriteHeight
End Sub

Public Sub Draw(destSurface As DirectDrawSurface7)
If DestSet Then destSurface.DrawCircle destX - frmMain.HS.Value * Map.TileWidth, destY - frmMain.VS.Value * Map.TileHeight, 3
destSurface.BltFast Int(PosX) - SpriteWidth / 2 - frmMain.HS.Value * Map.TileWidth, Int(PosY) - (SpriteHeight - 5) - frmMain.VS.Value * Map.TileHeight, ddsSoldier, CRect, DDBLTFAST_SRCCOLORKEY
End Sub

Private Function GetKoef() As Single
GetKoef = (dx.TickCount - CTick) / 1000
End Function

Public Sub UpdateMove()
If DestSet And CTick < dx.TickCount Then
    If VerifyPosition = True Then Exit Sub
    Dim Koef As Single
    Koef = GetKoef
    If Not Collision(Koef) = True Then
        PosX = PosX + MoveX * Koef
        PosY = PosY + MoveY * Koef
    End If
    CTick = dx.TickCount
End If
End Sub

Public Sub SetDestination(x As Integer, y As Integer)
destX = x
destY = y
Dim KoefX As Single, KoefY As Single, SizeY As Double, SizeX As Double, SizeC As Double
SizeX = Abs(destX - PosX)
SizeY = Abs(destY - PosY)
SizeC = Sqr(SizeX * SizeX + SizeY * SizeY) '//c2=a2+b2
If SizeX = 0 Or SizeY = 0 Then Exit Sub
KoefX = SizeX / SizeC
KoefY = SizeY / SizeC
If destX > PosX Then
    MoveX = Speed * KoefX
    Turn 0
    ViseManjeX = 1
ElseIf destX < PosX Then
    MoveX = (Speed * KoefX) * -1
    Turn 1
    ViseManjeX = 0
End If
If destY > PosY Then
    MoveY = Speed * KoefY
    ViseManjeY = 1
ElseIf destY < PosY Then
    MoveY = (Speed * KoefY) * -1
    ViseManjeY = 0
End If
DestSet = True
CTick = dx.TickCount
End Sub

Public Sub SetPosition(x As Integer, y As Integer)
PosX = x
PosY = y
CurrentKeyFrame = 1
StopMovement
End Sub

Public Function GetDestinationX() As Integer
GetDestinationX = destX
End Function

Public Function GetDestinationY() As Integer
GetDestinationY = destY
End Function

Public Function GetPositionX() As Integer
GetPositionX = PosX
End Function

Public Function GetPositionY() As Integer
GetPositionY = PosY
End Function

Private Function VerifyPosition() As Boolean
If ViseManjeX = 1 And PosX >= destX And ViseManjeY = 1 And PosY >= destY Then
    EmergencyStop
    VerifyPosition = True
ElseIf ViseManjeX = 1 And PosX >= destX And ViseManjeY = 0 And PosY <= destY Then
    EmergencyStop
    VerifyPosition = True
ElseIf ViseManjeX = 0 And PosX <= destX And ViseManjeY = 1 And PosY >= destY Then
    EmergencyStop
    VerifyPosition = True
ElseIf ViseManjeX = 0 And PosX <= destX And ViseManjeY = 0 And PosY <= destY Then
    EmergencyStop
    VerifyPosition = True
End If
End Function

Public Function GetTurn() As Integer
GetTurn = CRect.Top / SpriteHeight
End Function

Public Sub StopMovement()
DestSet = False
CurrentKeyFrame = 1
CRect.Left = 0
CRect.Right = CRect.Left + SpriteWidth
End Sub

Private Sub EmergencyStop()
DestSet = False
CurrentKeyFrame = 1
CRect.Left = 0
CRect.Right = CRect.Left + SpriteWidth
PosX = destX
PosY = destY
SpeechDest = False
End Sub

Private Function Collision(Koef As Single) As Boolean
Dim Result As String
Result = Map.GetTile(Int((PosX + MoveX * Koef) / Map.TileWidth), Int((PosY + MoveY * Koef) / Map.TileHeight))
If Not Result = ">" Then
    Collision = True
    StopMovement
ElseIf CollisionWithNPC(Int((PosX + MoveX * Koef)), Int((PosY + MoveY * Koef))) = True And SpeechDest = True Then
    Collision = True
    StopMovement
    InitSpeech GetNPCArrayIndex(Int((PosX + MoveX * Koef)), Int((PosY + MoveY * Koef)))
ElseIf CollisionWithNPC(Int((PosX + MoveX * Koef)), Int((PosY + MoveY * Koef))) = True Then
    Collision = True
    StopMovement
Else: Collision = False
End If
End Function

Private Function CollisionWithNPC(x As Integer, y As Integer) As Boolean
For i = 0 To 5
    If Not NPC(i).PosX = 0 And Not NPC(i).PosY = 0 And x >= NPC(i).PosX - 16 And _
    x <= NPC(i).PosX + 16 And y >= NPC(i).PosY - 8 And y <= NPC(i).PosY _
    Then CollisionWithNPC = True
Next
End Function

Public Sub Destroy()
Set dx = Nothing
End Sub


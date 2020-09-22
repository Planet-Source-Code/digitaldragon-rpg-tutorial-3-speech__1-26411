VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Map(100, 100) As Integer, dds As DirectDrawSurface7, MapSurface As DirectDrawSurface7
Private TilesCount As Integer
Public TileWidth As Integer, TileHeight As Integer
Public LenX As Integer, LenY As Integer

Public Sub Init()

TilesCount = 3
TileWidth = 32
TileHeight = 32
Dim ddsd As DDSURFACEDESC2
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd.lWidth = 384
ddsd.lHeight = 32
ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
Set dds = DD.CreateSurfaceFromFile(App.Path & "\Tiles.bmp", ddsd)
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
ddsd.lWidth = frmMain.Ekran.Width / 15 + 2 * TileWidth
ddsd.lHeight = frmMain.Ekran.Height / 15 + 2 * TileHeight
Set MapSurface = DD.CreateSurface(ddsd)
End Sub

Public Sub ChangeView(TileX As Integer, TileY As Integer)
DrawMap CreateRect(TileX * TileWidth, TileX * TileWidth + frmMain.Ekran.Width / 15, TileY * TileHeight, TileY * TileHeight + frmMain.Ekran.Height / 15)
End Sub

Public Sub LoadMap(Path As String)
Dim tStr As String, t1 As String, t2 As String
Dim Tag As String
Open Path For Input As 1
Do Until EOF(1)
    Line Input #1, Tag
    VerifyTag Tag
Loop
Close
End Sub

Private Sub VerifyTag(Tag As String)
If Tag = "<PPOS>" Then
    InputPPos
ElseIf Tag = "<MAP>" Then
    InputMap
ElseIf Tag = "<NPC>" Then
    InputNPC
ElseIf Tag = "<SPEECH>" Then
    InputSpeech
End If
End Sub

Private Sub InputSpeech()
Dim tStr As String, i As Integer, j As Integer
Line Input #1, tStr
    Do Until EOF(1)
    i = Val(Mid(tStr, 2, Len(tStr) - 2))
    Line Input #1, Speech(i).Pitanje
    Do
        Line Input #1, tStr
        If Left(tStr, 1) = "<" Then
            Speech(i).nOdgovor = j
            Exit Do
        End If
        Speech(i).Redirect(j) = Val(Left(tStr, 4))
        Speech(i).Odgovor(j) = Mid(tStr, 6)
        j = j + 1
    Loop
    j = 0
Loop
End Sub

Private Sub InputNPC()
Dim tStr As String, i As Integer
Do
    Line Input #1, tStr
    If tStr = "</NPC>" Then Exit Do
    NPC(i).NPCName = tStr
    Line Input #1, tStr
    NPC(i).PosX = Val(tStr)
    Line Input #1, tStr
    NPC(i).PosY = Val(tStr)
    Line Input #1, tStr
    NPC(i).SIndex = Val(tStr)
    Line Input #1, tStr
    NPC(i).ATIndex = Val(tStr)
    i = i + 1
Loop
End Sub

Private Sub InputPPos()
Dim pX As Integer, pY As Integer
Input #1, pX
Input #1, pY
Soldier.SetPosition pX, pY
End Sub

Private Sub InputMap()
Dim CounterX As Integer, CounterY As Integer, tStr As String
Line Input #1, tStr
Do Until tStr = "</MAP>"
    Do Until CounterX = Len(tStr)
        CounterX = CounterX + 1
        LenX = CounterX
        Map(CounterY, CounterX - 1) = Asc(Mid(tStr, CounterX, 1))
    Loop
    CounterX = 0
    CounterY = CounterY + 1
    Line Input #1, tStr
Loop
LenY = CounterY
End Sub

Public Sub DrawMap(Position As RECT)
MapSurface.BltColorFill Position, &HFFFFFF
Dim GetRect As RECT
For y = 0 To Int(Position.Bottom / TileHeight) - Int(Position.Top / TileHeight)
    For x = 0 To Int(Position.Right / TileWidth) - Int(Position.Left / TileWidth)
        GetRect.Bottom = TileHeight
        GetRect.Left = (Map(y + Position.Top / TileHeight, x + Position.Left / TileWidth) - 60) * TileWidth
        GetRect.Right = GetRect.Left + TileWidth
        MapSurface.BltFast x * TileWidth, y * TileHeight, dds, GetRect, DDBLTFAST_WAIT
    Next
Next
End Sub

Public Sub BltToBackBuffer(OffsetX As Integer, OffsetY As Integer, srcRect As RECT)
BackBuffer.BltFast OffsetX, OffsetY, MapSurface, srcRect, DDBLTFAST_WAIT
End Sub

Public Function GetTile(ArrayX As Integer, ArrayY As Integer) As String
GetTile = Chr(Map(ArrayY, ArrayX))
End Function

Sub Destroy()
Set dds = Nothing
Set MapSurface = Nothing
End Sub

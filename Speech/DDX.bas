Attribute VB_Name = "DDX"
Public Sub Flip()
    Primary.Blt GetWndRect, BackBuffer, GetSrcRect, DDBLT_WAIT
End Sub

Public Sub FlipToDC()
    BackBuffer.BltToDC frmMain.Ekran.hDC, GetSrcRect, GetSrcRect
    frmMain.Ekran.Refresh
End Sub

Public Function GetSrcRect() As RECT
GetSrcRect.Bottom = frmMain.Ekran.Height / 15
GetSrcRect.Right = frmMain.Ekran.Width / 15
End Function

Public Function GetWndRect() As RECT
dx.GetWindowRect frmMain.Ekran.hWnd, GetWndRect
End Function

Public Function CreateSurface(Optional Width As Integer, Optional Height As Integer, Optional ColorKey As Long) As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2, ckey As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS
If Width > 0 And Height > 0 Then ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Set CreateSurface = DD.CreateSurface(ddsd)
If ColorKey > 0 Then
    ckey.high = ColorKey
    ckey.low = ColorKey
    CreateSurface.SetColorKey DDCKEY_SRCBLT, ckey
End If
End Function

Public Function CreateSurfaceFromFile(Path As String, Optional ColorKey As Long, Optional Width As Integer, Optional Height As Integer) As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2, ckey As DDCOLORKEY
ddsd.lFlags = DDSD_CAPS
If Width > 0 And Height > 0 Then ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
Set CreateSurfaceFromFile = DD.CreateSurfaceFromFile(Path, ddsd)
If ColorKey > 0 Then
    ckey.high = ColorKey
    ckey.low = ColorKey
    CreateSurfaceFromFile.SetColorKey DDCKEY_SRCBLT, ckey
End If
End Function

Public Function CursorIsOverRect(x As Integer, y As Integer, r As RECT) As Boolean
If y >= r.Top And y <= r.Bottom And x >= r.Left And x <= r.Right Then CursorIsOverRect = True Else CursorIsOverRect = False
End Function

Public Function CreateRect(x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer) As RECT
CreateRect.Left = x1
CreateRect.Right = x2
CreateRect.Top = y1
CreateRect.Bottom = y2
End Function

Public Sub DrawNPC()
For i = 0 To 5
    If Not NPC(i).PosX = 0 Or Not NPC(i).PosY = 0 Then NPC(i).Draw BackBuffer
Next
End Sub

Public Function IsCursorOverNPC(x As Integer, y As Integer) As Boolean
For i = 0 To 5
    If Not NPC(i).PosX = 0 And Not NPC(i).PosY = 0 And x + _
    frmMain.HS.Value * 32 >= NPC(i).PosX - 16 And x + frmMain.HS.Value * 32 <= NPC(i).PosX + 16 And _
    y + frmMain.VS.Value * 32 >= NPC(i).PosY - 32 And y + frmMain.VS.Value * 32 <= NPC(i).PosY _
    Then IsCursorOverNPC = True
Next
End Function

Public Function GetNPCArrayIndexEx(x As Integer, y As Integer) As Integer
For i = 0 To 5
    If Not NPC(i).PosX = 0 And Not NPC(i).PosY = 0 And x + _
    frmMain.HS.Value * 32 >= NPC(i).PosX - 16 And x + frmMain.HS.Value * 32 <= NPC(i).PosX + 16 And _
    y + frmMain.VS.Value * 32 >= NPC(i).PosY - 32 And y + frmMain.VS.Value * 32 <= NPC(i).PosY _
    Then Exit For
Next
GetNPCArrayIndexEx = i
End Function

Public Function GetNPCArrayIndex(x As Integer, y As Integer) As Integer
For i = 0 To 5
    If Not NPC(i).PosX = 0 And Not NPC(i).PosY = 0 And x >= NPC(i).PosX - 16 And _
    x <= NPC(i).PosX + 16 And y >= NPC(i).PosY - 8 And y <= NPC(i).PosY _
    Then Exit For
Next
GetNPCArrayIndex = i
End Function
Public Sub InitSpeech(ArrayIndex As Integer)
TalkingTo = ArrayIndex
Dim tInt As Integer
If Not NPC(ArrayIndex).AlreadyTalkedTo = True Then
    NPC(ArrayIndex).AlreadyTalkedTo = True
    tInt = NPC(ArrayIndex).SIndex
Else: tInt = NPC(ArrayIndex).ATIndex
End If
CSpeech = tInt
frmMain.Pitanje.Caption = NPC(ArrayIndex).NPCName & " : " & vbCrLf & "::: " & Speech(tInt).Pitanje
frmMain.Choice(0).Top = frmMain.Pitanje.Top + frmMain.Pitanje.Height + 60
frmMain.Choice(0).Caption = ">>> " & Speech(tInt).Odgovor(0)
For i = 1 To Speech(tInt).nOdgovor - 1
    Load frmMain.Choice(frmMain.Choice.Count)
    frmMain.Choice(frmMain.Choice.Count - 1).Left = 120
    frmMain.Choice(frmMain.Choice.Count - 1).Top = frmMain.Choice(frmMain.Choice.Count - 2).Top + frmMain.Choice(frmMain.Choice.Count - 2).Height + 60
    frmMain.Choice(frmMain.Choice.Count - 1).Caption = ">>> " & Speech(tInt).Odgovor(i)
    frmMain.Choice(frmMain.Choice.Count - 1).AutoSize = True
    frmMain.Choice(frmMain.Choice.Count - 1).Visible = True
Next
frmMain.PicTXT.Visible = True
UpdateScrollBar
End Sub

Public Sub CloseSpeech()
frmMain.PicTXT.Visible = False
For i = 0 To frmMain.Choice.Count - 2
    Unload frmMain.Choice(frmMain.Choice.Count - 1)
Next
Soldier.SpeechDest = False
End Sub

Public Sub RedirectSpeech(Index As Integer, NPCIndex As Integer)
frmMain.TextScroll.Value = 0
If Index = 0 Then
    CloseSpeech
    Exit Sub
End If
For i = 0 To frmMain.Choice.Count - 2
    Unload frmMain.Choice(frmMain.Choice.Count - 1)
Next
tInt = Index
CSpeech = tInt
frmMain.Pitanje.Caption = NPC(NPCIndex).NPCName & " : " & vbCrLf & "::: " & Speech(tInt).Pitanje
frmMain.Choice(0).Top = frmMain.Pitanje.Top + frmMain.Pitanje.Height + 60
frmMain.Choice(0).Caption = ">>> " & Speech(tInt).Odgovor(0)
For i = 1 To Speech(tInt).nOdgovor - 1
    Load frmMain.Choice(frmMain.Choice.Count)
    frmMain.Choice(frmMain.Choice.Count - 1).Left = 120
    frmMain.Choice(frmMain.Choice.Count - 1).Top = frmMain.Choice(frmMain.Choice.Count - 2).Top + frmMain.Choice(frmMain.Choice.Count - 2).Height + 60
    frmMain.Choice(frmMain.Choice.Count - 1).Caption = ">>> " & Speech(tInt).Odgovor(i)
    frmMain.Choice(frmMain.Choice.Count - 1).AutoSize = True
    frmMain.Choice(frmMain.Choice.Count - 1).Visible = True
Next
UpdateScrollBar
End Sub

Public Sub UpdateScrollBar()
Dim tInt As Integer
tInt = frmMain.Choice(frmMain.Choice.Count - 1).Top - frmMain.Choice(frmMain.Choice.Count - 1).Height
frmMain.TextScroll.Max = tInt / (frmMain.PicTXT.Height / 2)
frmMain.PicTXT.SetFocus
End Sub

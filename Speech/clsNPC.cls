VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SIndex As Integer, ATIndex As Integer
Public PosX As Integer, PosY As Integer
Public NPCName As String, AlreadyTalkedTo As Boolean

Public Sub Draw(destSurface As DirectDrawSurface7)
destSurface.BltFast Int(PosX) - 32 / 2 - frmMain.HS.Value * Map.TileWidth, Int(PosY) - (32 - 5) - frmMain.VS.Value * Map.TileHeight, ddsSoldier, CreateRect(0, 32, 0, 32), DDBLTFAST_SRCCOLORKEY
End Sub

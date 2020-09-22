Attribute VB_Name = "Global"
'//DirectX Stuff
Public dx As New DirectX7
Public DD As DirectDraw7
Public Primary As DirectDrawSurface7
Public BackBuffer As DirectDrawSurface7
Public Clipper As DirectDrawClipper
'//DirectDraw Surfaces
Public ddsSoldier As DirectDrawSurface7
Public Tiles As DirectDrawSurface7
'//Classes declarations
Public Map As New clsMap
Public Soldier As New clsSoldier
Public RX As New RTSX
Public NPC(5) As New clsNPC
'//Other Stuff
Public Running As Boolean
Public DRect As RECT
Public NPCcount As Integer

Public Type SpeechType
    Pitanje As String
    Odgovor(10) As String
    Redirect(10) As Integer
    nOdgovor As Integer
End Type
Public Speech(30) As SpeechType
Public TalkingTo As Integer, CSpeech As Integer

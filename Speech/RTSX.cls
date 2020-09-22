VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTSX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FPSString As String, dx As New DirectX7

Public Function getFPSString()
Static FramesDone As Integer, LastTimeCount As Long
    FramesDone = FramesDone + 1
    If dx.TickCount >= LastTimeCount + 1000 Then
        FPSString = "FPS: " & CStr(FramesDone)
        LastTimeCount = dx.TickCount
    FramesDone = 0
    End If
End Function

Attribute VB_Name = "Initialization"
Public Sub Init()
Dim ddsd As DDSURFACEDESC2, MyFont As New StdFont
Set DD = dx.DirectDrawCreate("")
DD.SetCooperativeLevel frmMain.Ekran.hWnd, DDSCL_NORMAL
ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
ddsd.lWidth = frmMain.Ekran.Width / 15
ddsd.lHeight = frmMain.Ekran.Height / 15
Set BackBuffer = DD.CreateSurface(ddsd)
ddsd.lFlags = DDSD_CAPS
ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
Set Primary = DD.CreateSurface(ddsd)
Set Clipper = DD.CreateClipper(0)
Clipper.SetHWnd frmMain.Ekran.hWnd
Primary.SetClipper Clipper
Set ddsSoldier = CreateSurfaceFromFile(App.Path & "\Soldier.bmp", &HFFFFFF)
MyFont.Name = "Arial"
MyFont.Size = 9
BackBuffer.SetFont MyFont
End Sub

Public Sub Destroy()
Set dx = Nothing
Set DD = Nothing
Set Primary = Nothing
Set BackBuffer = Nothing
Set ddsSoldier = Nothing
Soldier.Destroy
Map.Destroy
End
End Sub

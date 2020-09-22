VERSION 5.00
Begin VB.UserControl AniText 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MousePointer    =   1  'Arrow
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox PicBox 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   960
      ScaleHeight     =   675
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "AniText1"
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3255
      End
   End
End
Attribute VB_Name = "AniText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private pCaption As String, pOriginColor As Long, pDestColor As Long
Private bStretch As Boolean
Event Click()
Event DblClick()

Public Property Get BackStyle() As Byte
BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(Value As Byte)
UserControl.BackColor = Value
PropertyChanged "BackStyle"
End Property

Public Property Get Caption() As String
Caption = pCaption
End Property

Public Property Let Caption(Value As String)
pCaption = Value
DrawText
PropertyChanged "Caption"
End Property

Public Property Get BackColor() As Long
BackColor = PicBox.BackColor
End Property
Public Property Let BackColor(Value As Long)
PicBox.BackColor = Value
PropertyChanged "BackColor"
End Property

Public Property Get OriginalColor() As Long
OriginalColor = pOriginColor
End Property
Public Property Let OriginalColor(Value As Long)
pOriginColor = Value
PropertyChanged "OriginalColor"
End Property

Public Property Get MouseOverColor() As Long
MouseOverColor = pDestColor
End Property
Public Property Let MouseOverColor(Value As Long)
pDestColor = Value
PropertyChanged "MouseOverColor"
End Property

Public Property Get Stretch() As Boolean
Stretch = bStretch
End Property
Public Property Let Stretch(Value As Boolean)
bStretch = Value
PropertyChanged "Stretch"
End Property

Private Sub UserControl_InitProperties()
Lbl.Caption = UserControl.Name
DrawText
UserControl.BackStyle = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
pCaption = PropBag.ReadProperty("Caption", 0)
PicBox.BackColor = PropBag.ReadProperty("BackColor", 0)
pOriginColor = PropBag.ReadProperty("OriginalColor", 0)
pDestColor = PropBag.ReadProperty("MouseOverColor", 0)
bStretch = PropBag.ReadProperty("Stretch", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BackStyle", UserControl.BackStyle, 1
PropBag.WriteProperty "Caption", pCaption, UserControl.Name
PropBag.WriteProperty "BackColor", PicBox.BackColor, 0
PropBag.WriteProperty "OriginalColor", pOriginColor, 0
PropBag.WriteProperty "MouseOverColor", pDestColor, 0
PropBag.WriteProperty "Stretch", bStretch, False
End Sub

Private Sub DrawText()
If Stretch = True Then
    If UserControl.BackStyle = 0 Then MsgBox TransparentBlt(UserControl.hdc, 0, 0, Lbl.Width / 15, Lbl.Height / 15, PicBox.hdc, 0, 0, UserControl.Width / 15, UserControl.Height / 15, PicBox.BackColor)
    If UserControl.BackStyle = 1 Then MsgBox StretchBlt(UserControl.hdc, 0, 0, Lbl.Width / 15, Lbl.Height / 15, PicBox.hdc, 0, 0, UserControl.Width / 15, UserControl.Height / 15, vbSrcCopy)
Else
    If UserControl.BackStyle = 0 Then MsgBox TransparentBlt(UserControl.hdc, 0, 0, Lbl.Width / 15, Lbl.Height / 15, PicBox.hdc, 0, 0, Lbl.Width / 15, Lbl.Height / 15, PicBox.BackColor)
    If UserControl.BackStyle = 1 Then MsgBox StretchBlt(UserControl.hdc, 0, 0, Lbl.Width / 15, Lbl.Height / 15, PicBox.hdc, 0, 0, Lbl.Width / 15, Lbl.Height / 15, vbSrcCopy)
End If
UserControl.Refresh
End Sub
'Public Property Get p() As p
'p = pp
'End Property
'Public Property Let p(Value As p)
'pp = Value
'PropertyChanged "p"
'End Property

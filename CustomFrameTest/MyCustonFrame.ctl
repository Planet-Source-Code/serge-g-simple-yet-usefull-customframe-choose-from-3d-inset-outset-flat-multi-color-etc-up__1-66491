VERSION 5.00
Begin VB.UserControl CustomFrame 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ControlContainer=   -1  'True
   ScaleHeight     =   90
   ScaleWidth      =   90
   ToolboxBitmap   =   "MyCustonFrame.ctx":0000
   Begin VB.Label lblEvent 
      BackStyle       =   0  'Transparent
      Height          =   15
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   15
   End
   Begin VB.Line line3DBot 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   600
      X2              =   615
      Y1              =   600
      Y2              =   615
   End
   Begin VB.Line line3DLeft 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   720
      X2              =   735
      Y1              =   720
      Y2              =   735
   End
   Begin VB.Line line3DRight 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   480
      X2              =   495
      Y1              =   480
      Y2              =   495
   End
   Begin VB.Line line3DTop 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   360
      X2              =   375
      Y1              =   240
      Y2              =   255
   End
   Begin VB.Line LineBottom 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   375
      Y1              =   255
      Y2              =   270
   End
   Begin VB.Line LineRight 
      BorderColor     =   &H00808080&
      X1              =   4320
      X2              =   4335
      Y1              =   840
      Y2              =   855
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00808080&
      X1              =   720
      X2              =   735
      Y1              =   720
      Y2              =   735
   End
   Begin VB.Line LineLeft 
      BorderColor     =   &H00808080&
      X1              =   720
      X2              =   735
      Y1              =   1200
      Y2              =   1215
   End
End
Attribute VB_Name = "CustomFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Public Enum FrameStyle
    cfFrameSunken = 1
    cfFrameRaised = 2
    cfFrameFlat = 4
End Enum

Public Enum FrameWidth
    cf1 = 1
    cf2 = 2
    cf3 = 3
End Enum

Public Enum FrameBackStyle
    Transparent = 0
    Opaque = 1
End Enum

Private theBackStyle As FrameBackStyle
Private theCaption As String
Private is3D As Boolean
Private theFrameStyle As FrameStyle
Private theFrameColor As OLE_COLOR
Private newFrameWidth As FrameWidth
Private theTopColor As OLE_COLOR
Private theRightColor As OLE_COLOR
Private theLeftColor As OLE_COLOR
Private theBottomColor As OLE_COLOR
Private meBackColor As OLE_COLOR
Private lTop As Integer
Private lBot As Integer
Private lLft As Integer
Private lRig As Integer

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Private Function ConvertSystemColor(ByVal theColor As Long) As Long
'
'    Call OleTranslateColor(theColor, 0, ConvertSystemColor)
'
'End Function

Private Sub lblEvent_Click()

    RaiseEvent Click
    
End Sub

Private Sub lblEvent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub lblEvent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub lblEvent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Initialize()

    is3D = False
    UserControl.Height = 450
    UserControl.Width = 1200
    theFrameStyle = cfFrameFlat
    newFrameWidth = cf1
    theTopColor = &H808080
    theBottomColor = &H808080
    theLeftColor = &H808080
    theRightColor = &H808080
    theFrameColor = &H808080
    theBackStyle = Opaque
    
    
    meBackColor = &H8000000F
    lTop = 0
    lBot = 0
    lLft = 0
    lRig = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        theTopColor = .ReadProperty("LineTopColor", &H808080)
        theBottomColor = .ReadProperty("LineBottomColor", &H808080)
        theLeftColor = .ReadProperty("LineLeftColor", &H808080)
        theRightColor = .ReadProperty("LineRightColor", &H808080)
        meBackColor = .ReadProperty("BackColor", &H8000000F)
        newFrameWidth = .ReadProperty("CustomWidth", 1)
        is3D = .ReadProperty("Frame3D", False)
        theCaption = .ReadProperty("FrameCaption", "")
        theBackStyle = .ReadProperty("FrameBackType", 1)
    End With

    setValues

End Sub

Private Sub UserControl_Resize()

    LineLeft.X1 = lLft
    LineLeft.X2 = lLft
    LineLeft.Y1 = 0
    LineLeft.Y2 = UserControl.ScaleHeight - 15
    LineLeft.ZOrder 0

    LineBottom.X1 = lBot
    LineBottom.X2 = UserControl.ScaleWidth
    LineBottom.Y1 = UserControl.Height - 15 - lBot
    LineBottom.Y2 = UserControl.Height - 15 - lBot
    LineBottom.ZOrder 0
    
    LineRight.X1 = UserControl.ScaleWidth - 15 - lRig
    LineRight.X2 = UserControl.ScaleWidth - 15 - lRig
    LineRight.Y1 = lRig
    LineRight.Y2 = UserControl.ScaleHeight
    LineRight.ZOrder 0
    
    LineTop.X1 = 0
    LineTop.X2 = UserControl.ScaleWidth
    LineTop.Y1 = lTop
    LineTop.Y2 = lTop
    LineTop.ZOrder 0
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    line3DTop.X1 = LineTop.X1
    line3DTop.X2 = LineTop.X2
    line3DTop.Y1 = LineTop.Y1 + 15
    line3DTop.Y2 = LineTop.Y2 + 15

    line3DBot.X1 = LineBottom.X1
    line3DBot.X2 = LineBottom.X2
    line3DBot.Y1 = LineBottom.Y1 - 15
    line3DBot.Y2 = LineBottom.Y2 - 15

    line3DLeft.X1 = LineLeft.X1 + 15
    line3DLeft.X2 = LineLeft + 15
    line3DLeft.Y1 = LineLeft.Y1
    line3DLeft.Y2 = LineLeft.Y2

    line3DRight.X1 = LineRight.X1 - 15
    line3DRight.X2 = LineRight.X2 - 15
    line3DRight.Y1 = LineRight.Y1
    line3DRight.Y2 = LineRight.Y2
    
    With lblEvent
        .Top = 0
        .Left = 0
        .Height = UserControl.Height
        .Width = UserControl.Width
    End With
    
    UserControl.Cls
    UserControl.CurrentY = (UserControl.ScaleHeight / 2) - 130
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (Len(theCaption) * 32) - 25
    UserControl.Print theCaption
    
    'Disable Tab-Stop
    'UserControl.Enabled = False
    
End Sub

Public Property Get FrameCaption() As String

    FrameCaption = theCaption

End Property

Public Property Let FrameCaption(newCaption As String)

    UserControl.Cls
    
    UserControl.CurrentY = (UserControl.ScaleHeight / 2) - 130
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (Len(newCaption) * 32) - 25
    UserControl.Print newCaption
    theCaption = newCaption
    PropertyChanged ("FrameCaption")

End Property

Public Property Get FrameBackType() As FrameBackStyle

    FrameBackType = theBackStyle

End Property

Public Property Let FrameBackType(newBackStyle As FrameBackStyle)

    UserControl.BackStyle = newBackStyle
    theBackStyle = newBackStyle

End Property


Public Property Get LineTopColor() As OLE_COLOR

    LineTopColor = theTopColor
    'ConvertSystemColor (LineTop.BorderColor)

End Property

Public Property Let LineTopColor(newColor As OLE_COLOR)

    LineTop.BorderColor = newColor
    theTopColor = newColor
    
    If theTopColor = theBottomColor And theTopColor = theRightColor And theTopColor = theLeftColor Then
        theFrameColor = theTopColor
    Else
        theFrameColor = 0
    End If

End Property

Public Property Get LineBottomColor() As OLE_COLOR
 
    LineBottomColor = theBottomColor
    'ConvertSystemColor (LineBottom.BorderColor)

End Property

Public Property Let LineBottomColor(newColor As OLE_COLOR)

    LineBottom.BorderColor = newColor
    theBottomColor = newColor
    
    If theTopColor = theBottomColor And theTopColor = theRightColor And theTopColor = theLeftColor Then
        theFrameColor = theBottomColor
    Else
        theFrameColor = 0
    End If

End Property

Public Property Get LineLeftColor() As OLE_COLOR

    LineLeftColor = theLeftColor
    'ConvertSystemColor (LineLeft.BorderColor)

End Property

Public Property Let LineLeftColor(newColor As OLE_COLOR)

    LineLeft.BorderColor = newColor
    theLeftColor = newColor
    
    If theTopColor = theBottomColor And theTopColor = theRightColor And theTopColor = theLeftColor Then
        theFrameColor = theLeftColor
    Else
        theFrameColor = 0
    End If

End Property

Public Property Get LineRightColor() As OLE_COLOR

    LineRightColor = theRightColor
    'ConvertSystemColor (LineRight.BorderColor)

End Property

Public Property Let LineRightColor(newColor As OLE_COLOR)

    LineRight.BorderColor = newColor
    theRightColor = newColor
    
    If theTopColor = theBottomColor And theTopColor = theRightColor And theTopColor = theLeftColor Then
        theFrameColor = theRightColor
    Else
        theFrameColor = 0
    End If

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = meBackColor
    'ConvertSystemColor (UserControl.BackColor)

End Property

Public Property Let BackColor(newColor As OLE_COLOR)

    UserControl.BackColor = newColor
    meBackColor = newColor

End Property

Public Property Get FrameColor() As OLE_COLOR

    FrameColor = theFrameColor
    'ConvertSystemColor (LineTop.BorderColor)
    'ConvertSystemColor (LineRight.BorderColor)
    'ConvertSystemColor (LineLeft.BorderColor)
    'ConvertSystemColor (LineBottom.BorderColor)

End Property

Public Property Let FrameColor(newColor As OLE_COLOR)

    setLineColor newColor
    theFrameColor = newColor

End Property

Public Property Get FrameType() As FrameStyle

    FrameType = theFrameStyle

End Property

Public Property Let FrameType(newStyle As FrameStyle)

    Select Case newStyle
    Case 1
        setFrameSunk
    Case 2:
        setFrameRaised
    Case 4:
        setAllBlack
    End Select
    theFrameStyle = newStyle
    
End Property

Public Property Get CustomWidth() As FrameWidth

    CustomWidth = newFrameWidth

End Property

Public Property Let CustomWidth(newWidth As FrameWidth)

    set3DFalse
    setFrameWidth newWidth
    newFrameWidth = newWidth
    
End Property

Public Property Get Frame3D() As Boolean

    Frame3D = is3D

End Property

Public Property Let Frame3D(bool3D As Boolean)

    set3D bool3D
    is3D = bool3D

End Property

Private Sub setAllBlack()

    LineTop.BorderColor = &H808080
    LineRight.BorderColor = &H808080
    LineBottom.BorderColor = &H808080
    LineLeft.BorderColor = &H808080
    theTopColor = &H808080
    theBottomColor = &H808080
    theLeftColor = &H808080
    theRightColor = &H808080

End Sub

Private Sub setWidth(ByRef w As FrameWidth)

    Dim Y As Integer
    
    Select Case w
    Case 2
        Y = 5
    Case Else
        Y = 0
    End Select
    
    Dim X As Integer
    X = (w * 5) - 5
    
    lTop = X + Y
    lBot = X + 0
    lLft = X + Y
    lRig = X + 0
    UserControl_Resize

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("LineTopColor", theTopColor, &H808080)
        Call .WriteProperty("LineBottomColor", theBottomColor, &H808080)
        Call .WriteProperty("LineRightColor", theRightColor, &H808080)
        Call .WriteProperty("LineLeftColor", theLeftColor, &H808080)
        Call .WriteProperty("BackColor", meBackColor, &H8000000F)
        Call .WriteProperty("CustomWidth", newFrameWidth, 1)
        Call .WriteProperty("Frame3D", is3D, False)
        Call .WriteProperty("FrameCaption", theCaption, "")
        Call .WriteProperty("FrameBackType", theBackStyle, 1)
    End With

End Sub

Private Sub setValues()

    LineTop.BorderColor = theTopColor
    LineBottom.BorderColor = theBottomColor
    LineLeft.BorderColor = theLeftColor
    LineRight.BorderColor = theRightColor
    
    If theTopColor = theBottomColor And theTopColor = theRightColor And theTopColor = theLeftColor Then
        theFrameColor = theTopColor
    Else
        theFrameColor = 0
    End If
    
    If theBackStyle > 1 Then theBackStyle = Opaque
    
    UserControl.BackStyle = theBackStyle
    UserControl.BackColor = meBackColor
    setFrameWidth newFrameWidth
    set3D is3D
    UserControl.Cls
    UserControl.CurrentY = (UserControl.ScaleHeight / 2) - 130
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (Len(theCaption) * 32) - 25
    UserControl.Print theCaption
    'lblEvent.Caption = theCaption

End Sub

Private Sub setLineColor(newColor As OLE_COLOR)

    LineTop.BorderColor = newColor
    LineRight.BorderColor = newColor
    LineLeft.BorderColor = newColor
    LineBottom.BorderColor = newColor
    theTopColor = newColor
    theBottomColor = newColor
    theLeftColor = newColor
    theRightColor = newColor
    
End Sub

Private Sub setFrameSunk()

    setAllBlack
    LineBottom.BorderColor = vbWhite
    LineRight.BorderColor = vbWhite
    theBottomColor = vbWhite
    theRightColor = vbWhite

End Sub

Private Sub setFrameRaised()

    setAllBlack
    LineTop.BorderColor = vbWhite
    LineLeft.BorderColor = vbWhite
    theTopColor = vbWhite
    theLeftColor = vbWhite

End Sub

Private Sub setFrameWidth(w As FrameWidth)

    LineTop.BorderWidth = w
    LineRight.BorderWidth = w
    LineLeft.BorderWidth = w
    LineBottom.BorderWidth = w
    setWidth w
    newFrameWidth = w
    
End Sub

Private Sub set3D(bool As Boolean)

    If bool = True Then
        set3DTrue
    Else
        set3DFalse
    End If

End Sub

Private Sub set3DTrue()

    newFrameWidth = LineTop.BorderWidth

    LineTop.BorderWidth = 1
    LineRight.BorderWidth = 1
    LineLeft.BorderWidth = 1
    LineBottom.BorderWidth = 1
    
    setWidth cf1
    
    LineTop.BorderColor = &H808080
    LineRight.BorderColor = vbWhite
    LineLeft.BorderColor = &H808080
    LineBottom.BorderColor = vbWhite
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    line3DTop.BorderColor = vbWhite
    line3DTop.Visible = True
    
    line3DRight.BorderColor = &H808080
    line3DRight.Visible = True
    
    line3DLeft.BorderColor = vbWhite
    line3DLeft.Visible = True
    
    line3DBot.BorderColor = &H808080
    line3DBot.Visible = True
    
    LineRight.ZOrder 0
    UserControl.Refresh

End Sub

Private Sub set3DFalse()

    LineTop.BorderWidth = newFrameWidth
    LineRight.BorderWidth = newFrameWidth
    LineLeft.BorderWidth = newFrameWidth
    LineBottom.BorderWidth = newFrameWidth
    
    LineTop.BorderColor = theTopColor
    LineRight.BorderColor = theRightColor
    LineLeft.BorderColor = theLeftColor
    LineBottom.BorderColor = theBottomColor
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    line3DTop.BorderColor = vbWhite
    line3DTop.Visible = False
    
    line3DRight.BorderColor = vbWhite
    line3DRight.Visible = False
    
    line3DLeft.BorderColor = &H808080
    line3DLeft.Visible = False
    
    line3DBot.BorderColor = &H808080
    line3DBot.Visible = False

    setWidth newFrameWidth
    is3D = False
    UserControl.Refresh

End Sub

VERSION 5.00
Begin VB.UserControl UISizeBar 
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ControlContainer=   -1  'True
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
   ToolboxBitmap   =   "UISizeBar.ctx":0000
End
Attribute VB_Name = "UISizeBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DEFAULT_HEIGHT As Integer = 2000
Private Const DEFAULT_WIDTH As Integer = 105

'' use sentinel values for max and min defaults
Private Const DEFAULT_MIN_VALUE As Long = 0
Private Const DEFAULT_MAX_VALUE As Long = 99999

Enum sbOrientationEnum
    Vertical = 0
    Horizontal = 1
End Enum

Enum sbBorderStyleEnum
    Flat = 0
    Raised = EDGE_RAISED
    Etched = EDGE_ETCHED
    Sunken = EDGE_SUNKEN
    Bump = EDGE_BUMP
End Enum

Enum sbSpeedEnum
    Slow = 0
    Medium = 1
    fast = 2
End Enum

'' local property holders
Private mvarOrientation As sbOrientationEnum
Private mvarBackColor As OLE_COLOR
Private mvarHighlightColor As OLE_COLOR
Private mvarMinLeft As Long
Private mvarMinTop As Long
Private mvarMaxLeft As Long
Private mvarMaxTop As Long
Private mvarShowBorder As Boolean
Private mvarBorderStyle As sbBorderStyleEnum
    
Private mvarTopEdge As Boolean
Private mvarLeftEdge As Boolean
Private mvarRightEdge As Boolean
Private mvarBottomEdge As Boolean

'' are we dragging it or not
Private blnDragging As Boolean

'' our current x and y values
Private lCurrentX As Long
Private lCurrentY As Long

'' events
Public Event MoveBegin(Left As Single, Top As Single, Right As Single, Bottom As Single)
Attribute MoveBegin.VB_Description = "The MoveBegin event is raised before a drag operation begins"
Public Event Move(Left As Single, Top As Single, Right As Single, Bottom As Single)
Attribute Move.VB_Description = "The Move event is raised multiple times during a drag operation. Every time a MouseMove event is received"
Public Event MoveComplete(Left As Single, Top As Single, Right As Single, Bottom As Single)

Public Function Animate(ToX As Single, ToY As Single, Speed As sbSpeedEnum) As Long
    Dim stepSize As Long, currStep As Long
    Dim toValue As Long, currValue As Long
    Dim strProp As String
    
    '' make sure that these values are within the bounds and if not
    '' correct them
    ToX = IIf(ToX > mvarMaxLeft, mvarMaxLeft, ToX)
    ToX = IIf(ToX < mvarMinLeft, mvarMinLeft, ToX)
    ToY = IIf(ToY > mvarMaxTop, mvarMaxTop, ToY)
    ToY = IIf(ToY < mvarMinTop, mvarMinTop, ToY)
     
    '' if the Sizebar is vertical, we are going left and right
    '' else we are going up and down, so select our appropriate
    '' destination coordinate, and current coordinate
    toValue = IIf(mvarOrientation = Vertical, ToX, ToY)
    currValue = IIf(mvarOrientation = Vertical, _
        UserControl.Extender.Left, UserControl.Extender.Top)
        
    '' if the SizeBar is vertical we will be adjusting the Left
    '' property, else we will be adjusting the Top Property
    strProp = IIf(mvarOrientation = Vertical, "Left", "Top")
    
    '' next we want to determine our stepSize, that is
    '' the distance we move the size bar per iteration
    '' we will do this by using the Speed parameter and the total
    '' distance we have to move
    Select Case Speed
        Case Slow
            stepSize = (toValue - currValue) / 250
        Case Medium
            stepSize = (toValue - currValue) / 100
        Case fast
            stepSize = (toValue - currValue) / 50
    End Select
    
    
    If stepSize <> 0 Then
        With UserControl
            ''raise a MoveBegin event
            RaiseEvent MoveBegin(.Extender.Left, .Extender.Top, .Extender.Left + .Width, .Extender.Top + .Height)
            
            Do
                '' calculate the new value (top or left)
                currValue = currValue + stepSize
                
                '' check bounds again to make sure we don't overshoot
                currValue = IIf(stepSize < 0 And currValue < toValue, toValue, currValue)
                currValue = IIf(stepSize > 0 And currValue > toValue, toValue, currValue)
                
                '' and set the value using call by name
                '' this eliminates loops
                CallByName .Extender, strProp, VbLet, currValue
                '' raise a move event
                RaiseEvent Move(.Extender.Left, .Extender.Top, .Extender.Left + .Width, .Extender.Top + .Height)
                DoEvents
            Loop While (currValue < toValue And stepSize > 0) Or _
                        (currValue > toValue And stepSize < 0)
            '' and raise our move complete event
            RaiseEvent MoveComplete(.Extender.Left, .Extender.Top, .Extender.Left + .Width, .Extender.Top + .Height)
        End With
        
    End If
    
End Function


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "The BackColor of the SizeBar"
    BackColor = mvarBackColor
    
End Property

Public Property Let BackColor(vData As OLE_COLOR)
    mvarBackColor = vData
    UserControl.BackColor = mvarBackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
    
End Property

Public Property Get BorderStyle() As sbBorderStyleEnum
Attribute BorderStyle.VB_Description = "Determines what type of border the SizeBar will have"
    BorderStyle = mvarBorderStyle
    
End Property

Public Property Let BorderStyle(vData As sbBorderStyleEnum)
    mvarBorderStyle = vData
    UserControl.Refresh
    PropertyChanged "BorderStyle"
    
End Property

Public Property Get Bottom() As Single
Attribute Bottom.VB_Description = "Returns the Bottom coordinate of the SizeBar. Provided for ease of use."
    Bottom = UserControl.Extender.Top + UserControl.Height
    
End Property

Public Property Let Edge_Bottom(vData As Boolean)
Attribute Edge_Bottom.VB_Description = "Determines whether or not the bottom edge is drawn. Used with BorderStyle."
    mvarBottomEdge = vData
    UserControl.Refresh
    PropertyChanged "Edge_Bottom"
    
End Property

Public Property Get Edge_Bottom() As Boolean
    Edge_Bottom = mvarBottomEdge
    
End Property

Public Property Let Edge_Left(vData As Boolean)
Attribute Edge_Left.VB_Description = "Determines whether or not the left edge is drawn. Used with BorderStyle."
    mvarLeftEdge = vData
    UserControl.Refresh
    PropertyChanged "Edge_Left"
    
End Property

Public Property Get Edge_Left() As Boolean
    Edge_Left = mvarLeftEdge
    
End Property


Public Property Let Edge_Right(vData As Boolean)
Attribute Edge_Right.VB_Description = "Determines whether or not the right edge is drawn. Used with BorderStyle."
    mvarRightEdge = vData
    UserControl.Refresh
    PropertyChanged "Edge_Right"
    
End Property

Public Property Get Edge_Right() As Boolean
    Edge_Right = mvarRightEdge
    
End Property

Public Property Let Edge_Top(vData As Boolean)
Attribute Edge_Top.VB_Description = "Determines whether or not the top edge is drawn. Used with BorderStyle."
    mvarTopEdge = vData
    UserControl.Refresh
    PropertyChanged "Edge_Top"
    
End Property

Public Property Get Edge_Top() As Boolean
    Edge_Top = mvarTopEdge
    
End Property

Public Property Let HighlightColor(vData As OLE_COLOR)
Attribute HighlightColor.VB_Description = "Determines the color used when dragging the SizeBar"
    mvarHighlightColor = vData
    PropertyChanged "HighlightColor"
    UserControl.Refresh
    
End Property

Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = mvarHighlightColor
    
End Property

Public Property Get MaxLeft() As Long
Attribute MaxLeft.VB_Description = "The maximum left value the SizeBar will be moved during dragging"
    MaxLeft = mvarMaxLeft
    
End Property

Public Property Let MaxLeft(vData As Long)
    mvarMaxLeft = vData
    PropertyChanged "MaxLeft"
    
End Property

Public Property Get MaxTop() As Long
Attribute MaxTop.VB_Description = "The maximum top value the SizeBar will be moved during dragging"
    MaxTop = mvarMaxTop
    
End Property

Public Property Let MaxTop(vData As Long)
    mvarMaxTop = vData
    PropertyChanged "MaxTop"
    
End Property

Public Property Get MinLeft() As Long
Attribute MinLeft.VB_Description = "The minimum left value the SizeBar will be moved during dragging"
    MinLeft = mvarMinLeft
    
End Property

Public Property Let MinLeft(vData As Long)
    mvarMinLeft = vData
    PropertyChanged "MinLeft"
    
End Property

Public Property Get MinTop() As Long
Attribute MinTop.VB_Description = "The minimum top value the SizeBar will be moved during dragging"
    MinTop = mvarMinTop
    
End Property

Public Property Let MinTop(vData As Long)
    mvarMinTop = vData
    PropertyChanged "MinTop"
    
End Property

Public Property Get Orientation() As sbOrientationEnum
Attribute Orientation.VB_Description = "Determines the orientation of the SizeBar, which determines the MousePointer which is displayed. "
    Orientation = mvarOrientation
    
End Property

Public Property Let Orientation(vData As sbOrientationEnum)
    Dim lNewHeight As Long, lNewWidth As Long
    
    lNewHeight = IIf(vData = Vertical, DEFAULT_HEIGHT, DEFAULT_WIDTH)
    lNewWidth = IIf(vData = Vertical, DEFAULT_WIDTH, DEFAULT_HEIGHT)
    UserControl.Height = lNewHeight
    UserControl.Width = lNewWidth
    mvarOrientation = vData
    UserControl.MousePointer = IIf(mvarOrientation = Horizontal, vbSizeNS, vbSizeWE)
    UserControl.Refresh
    
    PropertyChanged "Orientation"
End Property

Public Property Get Right() As Single
Attribute Right.VB_Description = "Returns the Right coordinate of the SizeBar. Provided for ease of use."
    Right = UserControl.Extender.Left + UserControl.Width
    
End Property

Public Property Get ShowBorderWhileMoving() As Boolean
Attribute ShowBorderWhileMoving.VB_Description = "Determines whether or not a border is drawn while the SizeBar is being dragged."
    ShowBorderWhileMoving = mvarShowBorder
    
End Property

Public Property Let ShowBorderWhileMoving(vData As Boolean)
    mvarShowBorder = vData
    PropertyChanged "ShowBorderWhileMoving"
    
End Property

Private Sub UserControl_Initialize()
    UserControl.Height = DEFAULT_HEIGHT
    UserControl.Width = DEFAULT_WIDTH
    
End Sub

Private Sub UserControl_InitProperties()
    mvarBackColor = vbButtonFace
    mvarHighlightColor = vbButtonShadow
    mvarMinLeft = DEFAULT_MIN_VALUE
    mvarMaxLeft = DEFAULT_MAX_VALUE
    mvarMinTop = DEFAULT_MIN_VALUE
    mvarMaxTop = DEFAULT_MAX_VALUE
    mvarOrientation = Vertical
    mvarShowBorder = False
    mvarBorderStyle = Raised
    mvarTopEdge = True
    mvarBottomEdge = True
    mvarLeftEdge = True
    mvarRightEdge = True
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '' if the mouse button pressed is the left one, start dragging
    blnDragging = (Button = vbLeftButton) Or blnDragging
    
    With UserControl
    '' highlight the window
        .BackColor = mvarHighlightColor
        .Refresh
        
        '' get our current coordinates
        lCurrentX = x
        lCurrentY = y
        
        'Call SetCapture(UserControl.hwnd)
        
        RaiseEvent MoveBegin( _
                .Extender.Left, _
                .Extender.Top, _
                .Extender.Left + .Width, _
                .Extender.Top + .Height)
    End With
                            
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lNewTop As Long
    Dim lNewLeft As Long
    Dim lPt As POINTAPI
    
    If blnDragging Then
     
        With UserControl
            
            ' if this is a vertical size bar, then move the left
            If mvarOrientation = Vertical Then
            
                '' determine the new left value by checking our endpoint
                '' conditions
                Select Case True
                    Case .Extender.Left + (x - lCurrentX) > mvarMaxLeft
                        lNewLeft = mvarMaxLeft '.ScaleX(mvarMaxLeft, .ScaleMode, .Parent.ScaleMode)
                    Case .Extender.Left + (x - lCurrentX) < mvarMinLeft
                        lNewLeft = mvarMinLeft '.ScaleX(mvarMinLeft, .ScaleMode, .Parent.ScaleMode)
                    Case Else
                        lNewLeft = .Extender.Left + .ScaleX((x - lCurrentX), .ScaleMode, .Parent.ScaleMode)
                End Select
                
                '' and set the property, we have to use the extender property
                '' therefore this object can only be used in containers which
                '' support the left / top properties
                .Extender.Left = lNewLeft
                
            '' else since this is a horizontal size bar, move the top
            Else
                '' deterimine the new top value by checking our endpoint
                '' conditions
                Select Case True
                    Case .Extender.Top + (y - lCurrentY) > mvarMaxTop
                        lNewTop = mvarMaxTop
                    Case .Extender.Top + (y - lCurrentY) < mvarMinTop
                        lNewTop = mvarMinTop
                    Case Else
                        lNewTop = .Extender.Top + .ScaleY((y - lCurrentY), .ScaleMode, .Parent.ScaleMode)
                End Select
                
                .Extender.Top = lNewTop
                
            End If
            
            '' refresh the control
            .Refresh
                        
            '' raise a move event
            RaiseEvent Move( _
                .Extender.Left, _
                .Extender.Top, _
                .Extender.Left + .Width, _
                .Extender.Top + .Height)
                
        End With
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '' we aren't dragging anymore
    blnDragging = False
    '' and we want to repaint the border
    With UserControl
        'Call ReleaseCapture
        .BackColor = mvarBackColor
        .Refresh
        RaiseEvent MoveComplete( _
                .Extender.Left, _
                .Extender.Top, _
                .Extender.Left + .Width, _
                .Extender.Top + .Height)
    End With
    
End Sub

Private Sub UserControl_Paint()
    Dim rc As RECT
    Dim lRetValue As Long
    Dim lFlags As Long
    Dim lEdge As Long
    Dim vControl As Variant
    
    '' just in case the control doesn't support the MousePointer property
    '' we will resume next
    On Local Error Resume Next
    For Each vControl In UserControl.ContainedControls
        vControl.MousePointer = vbArrow
    Next
    
    '' clear our previous error handling
    On Local Error GoTo 0
    
    '' clear the window
    UserControl.Cls
    
    '' determine which edges to draw
    lFlags = lFlags Or IIf(mvarTopEdge, BF_TOP, 0)
    lFlags = lFlags Or IIf(mvarLeftEdge, BF_LEFT, 0)
    lFlags = lFlags Or IIf(mvarBottomEdge, BF_BOTTOM, 0)
    lFlags = lFlags Or IIf(mvarRightEdge, BF_RIGHT, 0)
    
    '' determine what type of edge to draw
    lEdge = mvarBorderStyle
    
    '' if they want a raised border and they aren't dragging OR if
    '' they want the border shown while dragging, draw an edge
    
    If lEdge <> Flat Then
        If (Not blnDragging) Or _
            (blnDragging And mvarShowBorder) Then
        
            GetClientRect UserControl.hwnd, rc
            lRetValue = DrawEdge( _
                        UserControl.hdc, _
                        rc, _
                        lEdge, _
                        lFlags)
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        mvarBackColor = .ReadProperty( _
                                "BackColor", vbButtonFace)
        UserControl.BackColor = mvarBackColor
        mvarHighlightColor = .ReadProperty("HighlightColor", vbApplicationWorkspace)
        mvarOrientation = .ReadProperty("Orientation", Vertical)
        
        mvarMinLeft = .ReadProperty("MinLeft", DEFAULT_MIN_VALUE)
        mvarMinTop = .ReadProperty("MinTop", DEFAULT_MIN_VALUE)
        
        mvarMaxLeft = .ReadProperty("MaxLeft", DEFAULT_MAX_VALUE)
        mvarMaxTop = .ReadProperty("MaxTop", DEFAULT_MAX_VALUE)
        mvarShowBorder = .ReadProperty("ShowBorderWhileMoving", False)
        mvarBorderStyle = .ReadProperty("BorderStyle", Raised)
        
        mvarTopEdge = .ReadProperty("Edge_Top", True)
        mvarLeftEdge = .ReadProperty("Edge_Left", True)
        mvarBottomEdge = .ReadProperty("Edge_Bottom", True)
        mvarRightEdge = .ReadProperty("Edge_Right", True)
        
        UserControl.MousePointer = IIf(mvarOrientation = Horizontal, _
                                    vbSizeNS, vbSizeWE)
    End With
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", mvarBackColor, vbButtonFace
        .WriteProperty "HighlightColor", mvarHighlightColor, vbButtonShadow
        .WriteProperty "Orientation", mvarOrientation, Vertical
        .WriteProperty "ShowBorderWhileMoving", mvarShowBorder, False
        .WriteProperty "MinLeft", mvarMinLeft, DEFAULT_MIN_VALUE
        .WriteProperty "MaxLeft", mvarMaxLeft, DEFAULT_MAX_VALUE
        .WriteProperty "MinTop", mvarMinTop, DEFAULT_MIN_VALUE
        .WriteProperty "MaxTop", mvarMaxTop, DEFAULT_MAX_VALUE
        .WriteProperty "BorderStyle", mvarBorderStyle, Raised
        .WriteProperty "Edge_Left", mvarLeftEdge, True
        .WriteProperty "Edge_Right", mvarRightEdge, True
        .WriteProperty "Edge_Top", mvarTopEdge, True
        .WriteProperty "Edge_Bottom", mvarBottomEdge, True
    End With
    
End Sub



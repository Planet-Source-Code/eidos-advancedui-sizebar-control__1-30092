VERSION 5.00
Object = "*\ASizeBar Control\AdvancedUI.vbp"
Begin VB.Form frmMain 
   Caption         =   "SizeBar Demo - by eidos"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin AdvancedUI.UISizeBar UISizeBar2 
      Height          =   135
      Left            =   2280
      TabIndex        =   10
      Top             =   2130
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   238
      BackColor       =   16761024
      HighlightColor  =   16761024
      Orientation     =   1
      ShowBorderWhileMoving=   -1  'True
      Edge_Left       =   0   'False
      Edge_Right      =   0   'False
   End
   Begin AdvancedUI.UISizeBar UISizeBar1 
      Height          =   4035
      Left            =   30
      TabIndex        =   2
      Top             =   150
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   7117
      BackColor       =   16761024
      HighlightColor  =   16761024
      ShowBorderWhileMoving=   -1  'True
      Begin VB.CommandButton cmdAnimate 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   1470
         Width           =   315
      End
      Begin VB.OptionButton optSpeed 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Slow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   1200
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optSpeed 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   840
         Width           =   915
      End
      Begin VB.OptionButton optSpeed 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fast"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   210
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SizeBar Demo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Click to animate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   1770
         Width           =   1185
      End
      Begin VB.Label lblDes 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This is a SizeBar as well, it contains other controls!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   2970
         Width           =   1185
      End
   End
   Begin VB.PictureBox picSel 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1890
      ScaleHeight     =   1695
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   2310
      Width           =   5415
   End
   Begin VB.TextBox txtMain 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   1740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   -30
      Width           =   5475
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSpeed As sbSpeedEnum

Private Sub ResizeForm()
On Local Error Resume Next
    UISizeBar1.MaxLeft = Me.ScaleWidth - 500
    UISizeBar1.MinLeft = 0
    
    UISizeBar2.MaxTop = Me.ScaleHeight - 500
    UISizeBar2.MinTop = 200
    
    With UISizeBar1
        .Top = Me.ScaleTop
        .Height = Me.ScaleHeight
    End With
    
    With UISizeBar2
        .Left = UISizeBar1.Right
        .Width = Me.ScaleWidth - .Left
    End With
    
    With txtMain
        .Left = UISizeBar1.Right
        .Width = Me.ScaleWidth - .Left
        .Top = Me.ScaleTop
        .Height = UISizeBar2.Top
    End With
    
    With picSel
        .Left = UISizeBar1.Right
        .Width = Me.ScaleWidth - .Left
        .Top = UISizeBar2.Bottom
        .Height = Me.ScaleHeight - .Top
    End With
    
    With lblDes
        .Top = UISizeBar1.Height - .Height - 100
    End With
    
End Sub

Private Sub cmdAnimate_Click()
    If cmdAnimate.Caption = ">" Then
        UISizeBar1.Animate UISizeBar1.MaxLeft, 0, iSpeed
        cmdAnimate.Caption = "<"
    ElseIf cmdAnimate.Caption = "<" Then
        UISizeBar1.Animate UISizeBar1.MinLeft, 0, iSpeed
        cmdAnimate.Caption = ">"
    End If
    
    
    
End Sub

Private Sub Form_Load()
    Call ResizeForm
    
End Sub


Private Sub Form_Resize()
    Call ResizeForm
    
End Sub


Private Sub optSpeed_Click(Index As Integer)
    Select Case Index
        Case 0
            iSpeed = fast
        Case 1
            iSpeed = Medium
        Case 2
            iSpeed = Slow
            
    End Select
    
End Sub

Private Sub txtMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '' a little eye candy
    
    With picSel
        .Cls
        .AutoRedraw = True
        .Font.Name = "Tahoma"
        .Font.Size = 8
        .Font.Bold = True
        .CurrentX = 10
        .CurrentY = 10
        picSel.Print "Current Selection: "
        
        .Font.Bold = False
        .Font.Size = 8

        picSel.Print vbCrLf & txtMain.SelText
        
    End With
    
End Sub


Private Sub UISizeBar1_Move(Left As Single, Top As Single, Right As Single, Bottom As Single)
    ResizeForm
    
End Sub


Private Sub UISizeBar1_MoveComplete(Left As Single, Top As Single, Right As Single, Bottom As Single)
    ResizeForm
    
End Sub


Private Sub UISizeBar2_Move(Left As Single, Top As Single, Right As Single, Bottom As Single)
    ResizeForm
    
End Sub


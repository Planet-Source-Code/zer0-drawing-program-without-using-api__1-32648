VERSION 5.00
Begin VB.Form frmDraw 
   Caption         =   "Draw - By Zer0"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Screen"
      Height          =   735
      Left            =   11400
      TabIndex        =   57
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tool Modifiers"
      Height          =   5655
      Left            =   120
      TabIndex        =   37
      Top             =   4800
      Width           =   1695
      Begin VB.Frame Frame9 
         Caption         =   "Circle Style"
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   3120
         Width           =   1455
         Begin VB.OptionButton optSolid 
            Caption         =   "Solid"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optHollow 
            Caption         =   "Hollow"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "CircleSize"
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   2400
         Width           =   1455
         Begin VB.HScrollBar hsbCircle 
            Height          =   255
            Left            =   120
            Max             =   1000
            Min             =   100
            TabIndex        =   55
            Top             =   360
            Value           =   100
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Spray Size"
         Height          =   1095
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
         Begin VB.OptionButton optLargeSpray 
            Caption         =   "Large"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optMediumSpray 
            Caption         =   "Medium"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optSmallSpray 
            Caption         =   "Small"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Thichkness"
         Height          =   975
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1455
         Begin VB.HScrollBar hsbDrawWidth 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   2
            TabIndex        =   39
            Top             =   600
            Value           =   2
            Width           =   1215
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tools"
      Height          =   4695
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton optFill 
         Caption         =   "Fill"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton optErase 
         Caption         =   "Eraser"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2880
         Width           =   975
      End
      Begin VB.OptionButton optSquare 
         Caption         =   "Square"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton optSpray 
         Caption         =   "Spray"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optLine 
         Caption         =   "Line"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optSpiral 
         Caption         =   "Spiral"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optPencil 
         Caption         =   "Pencil"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   12720
      TabIndex        =   35
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Selected Colours"
      Height          =   1095
      Left            =   1920
      TabIndex        =   33
      Top             =   9360
      Width           =   1815
      Begin VB.PictureBox lblColour 
         BackColor       =   &H00000000&
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   1395
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colours"
      Height          =   1095
      Left            =   3840
      TabIndex        =   1
      Top             =   9360
      Width           =   7455
      Begin VB.Frame Frame2 
         Caption         =   "Custom Colour"
         Height          =   855
         Left            =   5280
         TabIndex        =   30
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton cmdCustom 
            Caption         =   "Change Custom Colour"
            Height          =   495
            Left            =   600
            TabIndex        =   32
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblCustomColour 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   27
         Left            =   4800
         TabIndex        =   29
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   26
         Left            =   4440
         TabIndex        =   28
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   25
         Left            =   4080
         TabIndex        =   27
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   24
         Left            =   3720
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   23
         Left            =   3360
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   22
         Left            =   3000
         TabIndex        =   24
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   21
         Left            =   2640
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   20
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   19
         Left            =   1920
         TabIndex        =   21
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   18
         Left            =   1560
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   17
         Left            =   1200
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   16
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   15
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   13
         Left            =   4800
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   12
         Left            =   4440
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   11
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   10
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   9
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   8
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   6
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblChooseColour 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   9360
      Left            =   1920
      ScaleHeight     =   9300
      ScaleWidth      =   13275
      TabIndex        =   0
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************************
'"Draw" Version 2 (C) Robbie Leggett AKA Zer0                                                  *
'**********************************************************************************************************************
Option Explicit
Dim LastX As Integer                    'LastX is one of the varibles that defines where to draw
Dim LastY As Integer                    'LastY is the other
Dim RandomX As Integer                  'RandomX determines one of the random factors of the spray tool
Dim RandomY As Integer                  'RandomY is the other
Dim SmallSpray As Integer               'SmallSpray is the number of dots drawn in the smallspray size
Dim MediumSpray As Integer              'MediumSpray is the number of dots drawn in the medium spray size
Dim LargeSpray As Integer               'LargeSPray is the number of dots drawn in the Largespray size
Dim Password As String                  'Password Stores the password
Dim CircleSize As Integer               'Circle size determines the size of the circle tool
Private Sub cmdClear_Click()            'Clears the Screen
    picDraw.Cls
    picDraw.BackColor = vbWhite
End Sub
Private Sub cmdExit_Click()             'Asks for the Exit Password then Exits
        Unload frmDraw                  'Exits the program
        Unload frmCustom
End Sub
Private Sub imgCircle_Click()           'These Subs just allow the user to click the pictures of tools instead of
    optCircle.Value = True              'the option buttons making the program simpler for a child.
End Sub                                 '/\
Private Sub imgErase_Click()            '/\
    optErase.Value = True               '/\
End Sub                                 '/\
Private Sub imgFill_Click()             '/\
    optFill.Value = True                '/\
End Sub                                 '/\
Private Sub imgFilled_Click()           '/\
    optFilled.Value = True              '/\
End Sub                                 '/\
Private Sub imgHollow_Click()           '/\
    optHollow.Value = True              '/\
End Sub                                 '/\
Private Sub imgLargeSpray_Click()       '/\
    optLargeSpray.Value = True          '/\
End Sub                                 '/\
Private Sub imgLine_Click()             '/\
    optLine.Value = True                '/\
End Sub                                 '/\
Private Sub imgMediumSpray_Click()      '/\
    optMediumSpray.Value = True         '/\
End Sub                                 '/\
Private Sub imgPencil_Click()           '/\
    optPencil.Value = True              '/\
End Sub                                 '/\
Private Sub imgSmallSpray_Click()       '/\
    optSmallSpray.Value = True          '/\
End Sub                                 '/\
Private Sub imgSpiral_Click()           '/\
    optSpiral.Value = True              '/\
End Sub                                 '/\
Private Sub imgSpray_Click()            '/\
    optSpray.Value = True               '/\
End Sub                                 '/\
Private Sub imgSquare_Click()           '/\
    optSquare.Value = True              '/\
End Sub                                 '/\

Private Sub cmdCustom_Click()
    frmCustom.Show
End Sub

Private Sub lblchoosecolour_Click(Index As Integer)         'Chooses the colour using a control array
    lblColour.BackColor = lblChooseColour(Index).BackColor
End Sub
Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
       LastX = X                'Sets up the X and Y variables
       LastY = Y
    If Button = 1 And optCircle.Value = True Then   'this is the Circle Tool
        If optHollow.Value = True Then      'Draws the Hollow Circle
         picDraw.Circle (LastX, LastY), hsbCircle.Value, lblColour.BackColor
        ElseIf optSolid = True Then        'Draws the Filled Circle
         picDraw.FillStyle = 0
         picDraw.FillColor = lblColour.BackColor
            picDraw.Circle (LastX, LastY), hsbCircle, vbBlack
        End If
            picDraw.FillStyle = 1           'Resets the fillstyle
    ElseIf Button = 1 And optFill.Value = True Then 'This is the fill tool(not a proper fill)
        picDraw.BackColor = lblColour.BackColor
    End If
End Sub
Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picDraw.DrawWidth = hsbDrawWidth.Value           'Sets the Property Drawwidth of picdraw to a horizontal scroll bar
   If Button = 1 And optPencil.Value = True Then
        picDraw.Line (LastX, LastY)-(X, Y), lblColour.BackColor     'Draws a line like a pencil
        LastX = X
        LastY = Y
    ElseIf Button = 1 And optSpiral.Value = True Then               'Draws a Spiral
        picDraw.Line (X, Y)-(LastX, LastY), lblColour.BackColor
    ElseIf Button = 1 And optLine.Value = True Then
        picDraw.Cls
        picDraw.Line (LastX, LastY)-(X, Y), lblColour.BackColor
    ElseIf Button = 1 And optSpray.Value = True Then                'This is the spray paint tool
         SmallSpray = 1
         MediumSpray = 1
         LargeSpray = 1
        For SmallSpray = 1 To 20                                   'This loop is for the small spray size
          If optSmallSpray.Value = True Then
            RandomX = 150 - Rnd * 300 + X                           'Sets the Random elment of the spray tool
            RandomY = 150 - Rnd * 300 + Y
            picDraw.Line (RandomX, RandomY)-(RandomX, RandomY), lblColour.BackColor
            SmallSpray = SmallSpray + 1
          End If
        Next
        For MediumSpray = 1 To 30                                  'This loop is for the Medium Spray size
          If optMediumSpray.Value = True Then
            RandomX = 300 - Rnd * 600 + X
            RandomY = 300 - Rnd * 600 + Y
            picDraw.Line (RandomX, RandomY)-(RandomX, RandomY), lblColour.BackColor
            MediumSpray = MediumSpray + 1
          End If
        Next
        For LargeSpray = 1 To 50                                    'This loop is for the large spray size
         If optLargeSpray.Value = True Then
            RandomX = 600 - Rnd * 1200 + X
            RandomY = 600 - Rnd * 1200 + Y
            picDraw.Line (RandomX, RandomY)-(RandomX, RandomY), lblColour.BackColor
            LargeSpray = LargeSpray + 1
          End If
        Next
    ElseIf Button = 1 And optSquare.Value = True Then
       picDraw.FillColor = lblColour.BackColor
       picDraw.FillStyle = 0
       picDraw.Line (LastX, LastY)-(X, Y), vbBlack, B               'Draws a rectangle (bit buggy)
       picDraw.FillStyle = 1
    ElseIf Button = 1 And optErase.Value = True Then
        picDraw.DrawWidth = 40                                                           'This is the eraser tool
        picDraw.Line (LastX, LastY)-(X, Y), vbWhite
        LastX = X
        LastY = Y
    End If
End Sub


VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   "Choose Custom Colour"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Colour"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   255
      Left            =   600
      Max             =   255
      Min             =   1
      TabIndex        =   2
      Top             =   720
      Value           =   1
      Width           =   2655
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   255
      Left            =   600
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   1080
      Value           =   1
      Width           =   2655
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   255
      Left            =   600
      Max             =   255
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   1
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Blue"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Green"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCustomColour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    frmDraw.lblColour.BackColor = lblCustomColour.BackColor
    frmDraw.lblCustomColour.BackColor = lblCustomColour.BackColor
    Unload Me
End Sub

Private Sub Form_Load()
    lblCustomColour.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub

Private Sub hsbBlue_Change()
    lblCustomColour.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub

Private Sub hsbGreen_Change()
    lblCustomColour.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub

Private Sub hsbRed_Change()
    lblCustomColour.BackColor = RGB(hsbRed.Value, hsbGreen.Value, hsbBlue.Value)
End Sub

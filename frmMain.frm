VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "OBS Window Size Helper"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInvertColor 
      Caption         =   "InvertColors"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSize1920x1080 
      Caption         =   "1920x1080"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSize1280x720 
      Caption         =   "1280x720"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSize1024x768 
      Caption         =   "1024x768"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSize800x600 
      Caption         =   "800x600"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Custom Size"
      Height          =   1455
      Left            =   1800
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
      Begin VB.TextBox txtCustomW 
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCustomH 
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdApplyCustomSize 
         Caption         =   "Apply Custom"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "W:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "H:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Timer updateCurrentPosition 
      Interval        =   200
      Left            =   3480
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custom Position"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
      Begin VB.CommandButton cmdApplyCustomPosition 
         Caption         =   "Apply Custom"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCustomLeft 
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCustomTop 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&X"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCenter 
      Caption         =   "Center Main"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdLeftTop 
      Caption         =   "Left Top Main"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to drag"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "by ^[GS]^"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   23
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "OBS Window Size Helper"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblCurrentSize 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Size: 0,0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   765
   End
   Begin VB.Label lblCurrentPosition 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pos: 0,0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape border 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   615
      Left            =   5640
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngX As Long
Private mlngY As Long
Private Const Horizontal = 8 * 1.875
Private Const Vertical = 10 * 1.5

Private Sub cmdApplyCustomPosition_Click()
    Me.Left = Int(txtCustomLeft.Text) * Horizontal
    Me.Top = Int(txtCustomTop.Text) * Vertical
End Sub

Private Sub cmdApplyCustomSize_Click()
    If Int(txtCustomH.Text) < 150 And Int(txtCustomH.Text) < 150 Then
        MsgBox "Too small"
        Exit Sub
    End If
    Me.Width = Int(txtCustomW.Text) * Horizontal
    Me.Height = Int(txtCustomH.Text) * Vertical
End Sub


Private Sub cmdInvertColor_Click()
    If Me.BackColor = vbWhite Then
        Me.BackColor = vbRed
        border.BorderColor = vbWhite
        border.FillColor = vbWhite
    Else
        Me.BackColor = vbWhite
        border.BorderColor = vbRed
        border.FillColor = vbRed
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If (Button And vbLeftButton) > 0 Then
        lngLeft = Me.Left + X - mlngX
        lngTop = Me.Top + Y - mlngY
        Me.Move lngLeft, lngTop
    End If
End Sub

Private Sub cmdLeftTop_Click()
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub cmdCenter_Click()
    Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
    Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
End Sub

Private Sub cmdSize1920x1080_Click()
    Me.Width = 1920 * Horizontal
    Me.Height = 1080 * Vertical
End Sub

Private Sub cmdSize1280x720_Click()
    Me.Width = 1280 * Horizontal
    Me.Height = 720 * Vertical
End Sub

Private Sub cmdSize1024x768_Click()
    Me.Width = 1024 * Horizontal
    Me.Height = 768 * Vertical
End Sub

Private Sub cmdSize800x600_Click()
    Me.Width = 800 * Horizontal
    Me.Height = 600 * Vertical
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub Form_Resize()
    border.Top = 0
    border.Left = 0
    border.Height = Me.Height
    border.Width = Me.Width
    
    cmdClose.Left = Me.Width - (cmdClose.Width + 100)
    cmdClose.Refresh
End Sub

Private Sub updateCurrentPosition_Timer()
    lblCurrentPosition.Caption = "Pos: " & (Me.Top / Vertical) & ", " & (Me.Left / Horizontal)
    lblCurrentSize.Caption = "Size: " & (Me.Width / Vertical) & ", " & (Me.Height / Horizontal)
End Sub

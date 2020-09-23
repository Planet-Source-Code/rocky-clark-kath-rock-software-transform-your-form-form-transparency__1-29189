VERSION 5.00
Object = "{EA994F5F-DC8F-45AF-B855-39FD0DD476CC}#1.1#0"; "TransForm.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kath-Rock - TransForm Demo"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Demo.frx":1042
   ScaleHeight     =   4680
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSys 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   870
      Width           =   285
   End
   Begin VB.CheckBox chkSys 
      BackColor       =   &H000000FF&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   870
      Width           =   285
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   75
      Picture         =   "Demo.frx":58EC4
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   90
      Width           =   300
   End
   Begin vbpTransForm.TransForm TransForm1 
      Left            =   240
      Top             =   3795
      _ExtentX        =   847
      _ExtentY        =   847
      AutoDrag        =   0
      MaskColor       =   16777215
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click the Icon above to activate the system menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1980
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   645
      Width           =   870
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TransForm Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   2
      Left            =   1860
      TabIndex        =   3
      Top             =   2145
      Width           =   3540
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kath-Rock's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   600
      Index           =   1
      Left            =   2085
      TabIndex        =   2
      Top             =   1365
      Width           =   3015
   End
   Begin VB.Label lblCap 
      BackStyle       =   0  'Transparent
      Caption         =   $"Demo.frx":5924E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1230
      Index           =   0
      Left            =   1605
      TabIndex        =   1
      Top             =   3045
      Width           =   4020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HTCAPTION         As Long = &H2&
Private Const WM_NCLBUTTONDOWN  As Long = &HA1&

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub chkSys_Click(Index As Integer)

    Select Case Index
        Case 0  'Minimize
            Me.WindowState = vbMinimized
        Case 1  'Close
            Unload Me
    End Select

End Sub

Private Sub chkSys_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    chkSys(Index).Value = vbUnchecked
    picIcon.SetFocus
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        Call ReleaseCapture
        Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, &H0&)
    End If
    
End Sub


Private Sub lblCap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Form_MouseDown(Button, Shift, X, Y)
    
End Sub


Private Sub picIcon_Click()

    Call TransForm1.PopupSysMenu(vbPopupMenuLeftAlign Or _
        vbPopupMenuRightButton, picIcon.Left, picIcon.Top + picIcon.Height)

End Sub



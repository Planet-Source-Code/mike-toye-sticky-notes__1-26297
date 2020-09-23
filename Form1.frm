VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   180
      ScaleHeight     =   1605
      ScaleWidth      =   2505
      TabIndex        =   0
      ToolTipText     =   "Left click and drag to move this sticky note"
      Top             =   60
      Width           =   2535
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   2440
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   2100
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   120
         ScaleWidth      =   150
         TabIndex        =   4
         Top             =   30
         Width           =   150
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   1920
         Picture         =   "Form1.frx":0142
         ScaleHeight     =   120
         ScaleWidth      =   150
         TabIndex        =   3
         Top             =   30
         Width           =   150
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   0
         Picture         =   "Form1.frx":0284
         ScaleHeight     =   180
         ScaleWidth      =   810
         TabIndex        =   2
         Top             =   0
         Width           =   810
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2320
         Picture         =   "Form1.frx":0A76
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   1
         Top             =   20
         Width           =   165
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "User32" (ByVal _
    hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()


Sub SetTopmostWindow(ByVal hWnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hWnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Load()
    Picture2.Left = 0
    Picture2.Top = 0
    Me.Width = Picture2.Width
    Me.Height = Picture2.Height
    SetTopmostWindow Me.hWnd
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = 60
Dim sCIN As String
    If sCIN > "" Then
        Text1 = Replace(sCIN, "|", vbCrLf)
    Else
        Text1 = ""
    End If
End Sub

Private Sub Picture1_Click()
    Unload Me
    End
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Private Sub Picture4_Click()
    Picture2.Height = 460
    Me.Height = Picture2.Height
End Sub

Private Sub Picture5_Click()
    Picture2.Height = 1635
    Me.Height = Picture2.Height
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame frmSize 
      Caption         =   "Distance To Resize From Bottom (In Pixels)"
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   4215
      Begin VB.TextBox txtDistance 
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblResize 
         Caption         =   "Resize Distance:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frmSDA 
      Caption         =   "Set Desktop Area:"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   4335
      Begin VB.OptionButton optCurrent 
         Caption         =   "From Current Screen "
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optFull 
         Caption         =   "From Full Screen"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdResizeDesktop 
      Caption         =   "Resize"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReset_Click()
        If Not SetDesktopArea(RF_FROMFULL, , , , 30) Then MsgBox "Cannot Resize!"
        frmMain.WindowState = 1 'Minimize Window
        frmMain.WindowState = 2 'Maximize Window
End Sub

Private Sub cmdResizeDesktop_Click()
    If Not IsNumeric(txtDistance.Text) Then
        MsgBox "Must Enter a Numeric Value"
        Exit Sub
    End If
    
    If Val(txtDistance.Text) > (Screen.Height / Screen.TwipsPerPixelY) Then
        MsgBox "Distance number must be LESS than current screen height (in pixels)"
        Exit Sub
    End If
    If optFull.Value = True Then
        If Not SetDesktopArea(RF_FROMFULL, , , , Val(txtDistance.Text)) Then MsgBox "Cannot Resize!"
    Else
        If Not SetDesktopArea(RF_FROMCURRENT, , , , Val(txtDistance.Text)) Then MsgBox "Cannot Resize!"
    End If
    
    frmMain.WindowState = 1 'Minimize Window
    frmMain.WindowState = 2 'Maximize Window
End Sub

Private Sub Form_Load()
    frmMain.WindowState = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not SetDesktopArea(RF_FROMFULL, , , , 30) Then MsgBox "Cannot Resize!"
End Sub

Private Sub optCurrent_Click()
    If optCurrent.Value = False Then
        optCurrent.Value = True
        optFull.Value = False
    End If
End Sub

Private Sub optFull_Click()
    If optFull.Value = False Then
        optFull.Value = True
        optCurrent.Value = False
    End If
End Sub

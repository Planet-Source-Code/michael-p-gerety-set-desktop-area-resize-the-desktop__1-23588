Attribute VB_Name = "modSetDesktopArea"
'**************************************************
'* modSetDesktopArea by Michael P. Gerety *********
'**************************************************
'
'This module contains function to set the desktop area,
'or resize the desktop.
'
'This could be useful for creating a shell replacement or a
'dockable form, so that your applications don't maximize over it
'

'USAGE:
'
'
' ***To set desktop area to allow 10 pixels on every side of the screen....
' If Not SetDesktopArea(RF_FROMFULL, 10, 10, 10, 10) Then MsgBox "Cannot Resize!"
'
' ***To set desktop area to take up 30 less pixels from bottom of screen than currently....
' If Not SetDesktopArea(RF_FROMCURRENT, , , , 30) Then MsgBox "Cannot Resize!"
'
' ***To RESET DESKTOP AREA to FULL screen size
' If Not SetDesktopArea(RF_FROMFULL) Then MsgBox "Cannot Resize!"
'
' *** To RESET DESKTOP size to Windows Default:
' If Not SetDesktopArea(RF_FROMFULL, , , , 30) Then MsgBox "Cannot Resize!"
'
'
'
'*********** NOTE ***********
'It is usually a good idea to throw the LAST example into your Form_Close() Sub.
'Your Desktop will STAY AFFECTED after you close your program.  You have to reset it manually.
'****************************

'API Functions
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

'Constants
Public Const SPI_GETWORKAREA = 48
Public Const SPI_SETWORKAREA = 47
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_SENDCHANGE = SPIF_SENDWININICHANGE

'Type Declarations
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum E_RESIZEFROM
    RF_FROMCURRENT = &H4
    RF_FROMFULL = &H5
End Enum


'Public Function SetDesktopArea(ByVal lTop As Long, ByVal lRight As Long, ByVal lLeft As Long, ByVal lBottom As Long)
'Author: Michael P. Gerety
'Description: Set the desktop area, i.e. area of screen Applications maximize to.
'
'Parameters: lTop - Distance from Top of Screen (in Pixels)
'            lRight - Distance from Right of Screen (in Pixels)
'            lLeft - Distance from Left of Screen (in Pixels)
'            lBottom - Distance from Bottom of Screen (in Pixels)

Public Function SetDesktopArea(ByVal RF_FROM As E_RESIZEFROM, Optional lTop As Long = 0, Optional lRight As Long = 0, Optional lLeft As Long = 0, Optional lBottom As Long = 0) As Boolean
    'Dimension variables for screen, screen height, and screen width
    Dim rctScreen As RECT, intScreenHeight As Integer, intScreenWidth As Integer, lResult As Long
    
    
    'Get Screen Rectangle
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, rctScreen, SPIF_SENDCHANGE)
    
    Select Case RF_FROM
        Case RF_FROMCURRENT
            'Set Screen Size from CURRENT size
            rctScreen.Top = rctScreen.Top + lTop
            rctScreen.Bottom = rctScreen.Bottom - lBottom
            rctScreen.Left = rctScreen.Left + lLeft
            rctScreen.Right = rctScreen.Right - iRight
            
            'Attempt to submit changes
            lResult = SystemParametersInfo(SPI_SETWORKAREA, 0, rctScreen, SPIF_SENDCHANGE)
            
            'Check and see if successful.. If not, return false
            If lResult = 0 Then SetDesktopArea = False Else SetDesktopArea = True
            
            'Exit Function
            Exit Function
        Case RF_FROMFULL
            'Get Screen Height & Width
            iScreenHeight = Screen.Height / Screen.TwipsPerPixelY
            iScreenWidth = Screen.Width / Screen.TwipsPerPixelX
            
            'Set values from FULL SCREEN
            rctScreen.Top = 0 + lTop
            rctScreen.Bottom = iScreenHeight - lBottom
            rctScreen.Left = 0 + lLeft
            rctScreen.Right = iScreenWidth - lRight
            
            'Attempt to submit changes
            lResult = SystemParametersInfo(SPI_SETWORKAREA, 0, rctScreen, SPIF_SENDCHANGE)
            
            'Check and see if successful.. If not, return false
            If lResult = 0 Then SetDesktopArea = False Else SetDesktopArea = True
            
            'Exit Function
            Exit Function
        Case Else
            'If incorrect E_RESIZEFROM parameter set, then return false
            SetDesktopArea = False
            Exit Function
        End Select
    
End Function
    

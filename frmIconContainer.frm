VERSION 5.00
Begin VB.Form frmIconContainer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISPN Client [Not Logged On]"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1500
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmIconContainer.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   1500
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Icon container for offline status. Also includes tray menu handler."
      Height          =   945
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   1470
   End
End
Attribute VB_Name = "frmIconContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com


'Tray menu handler (unused)

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next

'Static lngMsg As Long
'Static blnFlag As Boolean
'Dim result As Long

'lngMsg = X / Screen.TwipsPerPixelX

'If blnFlag = False Then
    
'    blnFlag = True
        
'    Select Case lngMsg
        
'        Case WM_LBUTTONDBLCLICK 'Double Click
        
'            frmClient.Show
        
'        Case WM_RBUTTONUP 'Right Button
        
'            PopupMenu frmClient.mnuTrayMenu, , , , frmClient.mnuToggleVisible
'            result = SetForegroundWindow(frmClient.hWnd)

'    End Select

'    blnFlag = False

'End If

'End Sub

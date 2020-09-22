VERSION 5.00
Begin VB.Form frmAlerter 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1365
   ClientLeft      =   16695
   ClientTop       =   13500
   ClientWidth     =   1800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmAlerter"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAlerter.frx":0000
   ScaleHeight     =   1365
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   1470
      Top             =   -45
   End
   Begin VB.Image imgLogoff 
      Height          =   240
      Left            =   795
      Picture         =   "frmAlerter.frx":093A
      Top             =   1050
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgStandard 
      Height          =   240
      Left            =   1050
      Picture         =   "frmAlerter.frx":0CC4
      Top             =   1035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIM 
      Height          =   240
      Left            =   585
      Picture         =   "frmAlerter.frx":104E
      Top             =   1065
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgiMail 
      Height          =   240
      Left            =   300
      Picture         =   "frmAlerter.frx":371E
      Top             =   1035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUsers 
      Height          =   240
      Left            =   45
      Picture         =   "frmAlerter.frx":3AA8
      Top             =   1035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAlert 
      Height          =   240
      Left            =   30
      Picture         =   "frmAlerter.frx":3E32
      Top             =   15
      Width           =   240
   End
   Begin VB.Label hlnk_options 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "options"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1155
      MouseIcon       =   "frmAlerter.frx":41BC
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label hlnk_alertlnk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "{Unknown} {unspecified alert}"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   810
      Left            =   60
      MouseIcon       =   "frmAlerter.frx":44C6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "This"
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: i s p n ::. "
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   435
      TabIndex        =   0
      Top             =   15
      Width           =   930
   End
End
Attribute VB_Name = "frmAlerter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Dim Alert_Action As String
Dim AlertType_Local As Integer

Public Sub ShowAlert(AlertText As String, AlertType As Integer, Optional ActionShellString As String)
AlertType_Local = AlertType

'Set window position on screen
frmAlerter.Top = Screen.Height - (frmAlerter.Height + 400) '('X' Co-ordinate)
frmAlerter.Left = Screen.Width - frmAlerter.Width '('Y' Co-ordinate)

'Set the alert text
hlnk_alertlnk = AlertText

'Set a relavent icon for the window
Select Case AlertType_Local

    Case ISPN_ALERT_UNSPECIFIED
    
        imgAlert.Picture = imgStandard
    
    Case ISPN_ALERT_USERLOGON
    
        imgAlert.Picture = imgUsers
        Alert_Action = ActionShellString
    
    Case ISPN_ALERT_USERLOGOFF
    
        imgAlert.Picture = imgLogoff
        Alert_Action = ActionShellString
        
    Case ISPN_ALERT_NEWIMAIL
        
        imgAlert.Picture = imgiMail
    
    Case ISPN_ALERT_IM '[Internal Alert]
        
        imgAlert.Picture = imgIM
        Alert_Action = ActionShellString
      
End Select

'Show the alerter (with 'ping' visual effect)
frmAlerter.Show
frmAlerter.WindowState = 0

'Put the alerter window on top of other windows..
PROBas.FormsOnTop frmAlerter, True

'Close the alerter window in the alloted time & reset the timer
tmrUnload.Enabled = False
tmrUnload.Interval = 4400 '4.4 Second timeout
tmrUnload.Enabled = True
End Sub

Private Sub hlnk_alertlnk_Click()
Select Case AlertType_Local

    Case ISPN_ALERT_UNSPECIFIED
    
        MsgBox Alert_Action, 64, "Standard Alert"
    
    Case ISPN_ALERT_USERLOGON
    
        'perform default user action
        modISPN.PerfomDefaultAction Alert_Action

    Case ISPN_ALERT_USERLOGOFF
    
    Case ISPN_ALERT_NEWIMAIL
        
        MsgBox "iMail has not been setup correctly on your client. Please update your client software", 16

    Case ISPN_ALERT_IM
    
        'Give focus to the IM window
        Call modIMHandler.DisplayIM_VirtualWindow(TrimHandle(Alert_Action) & GetAttachedServer(True), , True)

End Select
End Sub

Private Sub hlnk_options_Click()
frmAlertSettings.Show
End Sub

Private Sub tmrUnload_Timer()
tmrUnload.Enabled = False
Unload frmAlerter
End Sub

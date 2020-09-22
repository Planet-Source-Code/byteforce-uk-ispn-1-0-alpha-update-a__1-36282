VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon to ISPN provider"
   ClientHeight    =   3330
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1967.474
   ScaleMode       =   0  'User
   ScaleWidth      =   4084.415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAutoLogon 
      Appearance      =   0  'Flat
      Caption         =   "Log me in automatically"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   2940
      Width           =   2160
   End
   Begin VB.CommandButton cmdLogon 
      Caption         =   "&Logon"
      Default         =   -1  'True
      Height          =   390
      Left            =   3240
      TabIndex        =   7
      Top             =   2850
      Width           =   1035
   End
   Begin VB.TextBox txtServerIP 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1395
      TabIndex        =   3
      Top             =   1650
      Width           =   1590
   End
   Begin VB.TextBox txtPort 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1395
      TabIndex        =   5
      Top             =   2340
      Width           =   915
   End
   Begin VB.TextBox txtPIN 
      ForeColor       =   &H000000C0&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "-"
      TabIndex        =   4
      Top             =   1995
      Width           =   915
   End
   Begin VB.TextBox txtUserHandle 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Top             =   885
      Width           =   2895
   End
   Begin VB.TextBox txtPassword 
      ForeColor       =   &H000000C0&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1230
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   3585
      TabIndex        =   13
      Top             =   405
      Width           =   285
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5633.676
      Y1              =   1630.699
      Y2              =   1630.699
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Port:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   11
      Top             =   2370
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Server PIN:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   2025
      Width           =   900
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Password:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "User Handle:"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   8
      Top             =   915
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Logon"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   780
      TabIndex        =   0
      Top             =   165
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   195
      Picture         =   "frmLogin.frx":0CCA
      Top             =   135
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   -14.084
      X2              =   5619.592
      Y1              =   443.125
      Y2              =   443.125
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -1620
      Picture         =   "frmLogin.frx":1994
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Private Sub cmdLogon_Click()
'Check for null input
If CheckNull(txtUserHandle, True) = True Then MsgBox "Please enter a valid user handle to log on with.", 48, "Logon to ISPN": Exit Sub
If CheckNull(txtPassword, True) = True Then MsgBox "Please enter a valid password to log on with.", 48, "Logon to ISPN": Exit Sub
If CheckNull(txtPIN, True) = True Then MsgBox "Please enter a valid server PIN to log on with.", 48, "Logon to ISPN": Exit Sub
If CheckNull(txtServerIP, True) = True Then MsgBox "Please enter a valid server IP address or domain name to log on to.", 48, "Logon to ISPN": Exit Sub
If CheckNull(txtPort, True) = True Then MsgBox "Please enter a valid server port to log on to.", 48, "Logon to ISPN": Exit Sub

'Disable controls
txtPassword.Enabled = False
cmdLogon.Enabled = False
chkAutoLogon.Enabled = False
txtPIN.Enabled = False
txtServerIP.Enabled = False
txtUserHandle.Enabled = False
txtPort.Enabled = False

'If error is encountered goto the error handler
On Error GoTo errhandle

'Save variables
modISPN.ISPN_LocalHandle = txtUserHandle
modISPN.ISPN_GivenPIN = txtPIN
modISPN.ISPN_LocalPassword = txtPassword
modISPN.ISPN_AddressAttachedTo = txtServerIP
modISPN.ISPN_PortAttachedTo = txtPort

'Show\Hide Controls on the main client window
frmClient.fraLogon.Visible = False
frmClient.imgWait.Visible = True
frmClient.lblWait.Visible = True

'Enable\Disable menu items
frmClient.mnuLogon.Enabled = False
frmClient.mnuLogoff2.Enabled = False


'Show the logon indicator form
frmLogonStatus.Show
frmLogonStatus.lblStatus = "Finding host: ispn://" & ISPN_AddressAttachedTo & ":" & ISPN_PortAttachedTo & "..."

'Log on to the ISPN
'-Please see protocol.txt for more information on how to communicate with the server using TCP\IP

If frmClient.Client.State <> 0 Then frmClient.Client.Close
frmClient.Client.Connect ISPN_AddressAttachedTo, ISPN_PortAttachedTo

Exit Sub

errhandle: 'If any errors are encountered, the focus is sent to this label
           '(as defined in On Error Goto [linelabelornumber])

Select Case Err.Number
    
    Case 13 'Type Mismatch (Wrong variable type assigned, eg; Tried to set a Long variable to a string value)
    
    MsgBox "One or more logon details you entered were the wrong format. Please enure that the server PIN and server port are numeric values. The server IP may also be numeric, but a domain name can be specified in place of an IP address to ensure a static address.", 16, "Logon Error"
    
    Case Else 'Unhandled error.. I havent forseen this error.
    
    MsgBox "Unhandled Exception " & Err & " in module ISPNCLIENT.FRMLOGIN.CMDLOGON_CLICK. Please check your logon credentials and re-start ISPN client, or; contact your vendor.", 16, "Unhandled Exception"

End Select

'Enable controls
txtPassword.Enabled = True
cmdLogon.Enabled = True
chkAutoLogon.Enabled = True
txtPIN.Enabled = True
txtServerIP.Enabled = True
txtUserHandle.Enabled = True
txtPort.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next ' No errors that we need to deal with here

'Load default logon data from user specific registry key
txtUserHandle = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_UserHandle")
txtPassword = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_LocalPassword")
txtServerIP = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_ServerIP")
txtPIN = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_ServerPIN")
txtPort = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_ServerPort")
frmLogin.chkAutoLogon.Value = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_AutoLogon", frmLogin.chkAutoLogon.Value)

'Set the focus on the User Handle text box
txtUserHandle.SetFocus
End Sub

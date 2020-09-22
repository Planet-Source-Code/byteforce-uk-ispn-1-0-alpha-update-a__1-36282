VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISPN Client - [Not Logged On]"
   ClientHeight    =   5220
   ClientLeft      =   7560
   ClientTop       =   4755
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "ISPNClient"
   MaxButton       =   0   'False
   Picture         =   "frmContacts.frx":0A3A
   ScaleHeight     =   5220
   ScaleWidth      =   4590
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   3915
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":46F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":472EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":47688
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":47A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContacts.frx":47DBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstContacts 
      Height          =   3840
      Left            =   -6000
      TabIndex        =   1
      ToolTipText     =   "Displays contacts that you have added. Right click a user to perform an action."
      Top             =   705
      Visible         =   0   'False
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   6773
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imlLarge"
      SmallIcons      =   "imlSmall"
      ForeColor       =   4210752
      BackColor       =   0
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
      Picture         =   "frmContacts.frx":48156
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   -15
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraLogon 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   60
      TabIndex        =   6
      Top             =   1335
      Width           =   4470
      Begin VB.CheckBox chkAutoLogon 
         Appearance      =   0  'Flat
         Caption         =   "Log me in automatically"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   660
         Width           =   2160
      End
      Begin VB.Label hlnk_Logon 
         BackStyle       =   0  'Transparent
         Caption         =   "Click here to logon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   240
         Left            =   375
         MouseIcon       =   "frmContacts.frx":4A4F9
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label lblLogonCaption 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "You are not logged on!"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   90
         Width           =   1725
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   90
         Picture         =   "frmContacts.frx":4A803
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.Image imgWait 
      Height          =   480
      Left            =   1538
      Picture         =   "frmContacts.frx":4AB8D
      Top             =   2055
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblWait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait.."
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   2093
      TabIndex        =   11
      Top             =   2265
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3675
      Picture         =   "frmContacts.frx":4B857
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label lblISPNVerDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha Source Code Release"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   75
      TabIndex        =   10
      Top             =   4290
      Width           =   4395
   End
   Begin VB.Label lblISPNTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ISPN Client 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   4125
      Width           =   4395
   End
   Begin VB.Label lblHandleHighlight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(not logged on)"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   1710
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lblAttachedToHighlight 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(not logged on)"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   1710
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lblHandle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Handle: (not logged on)"
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   1710
      TabIndex        =   3
      Top             =   315
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lblAttachedTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Attached To: (not logged on)"
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   1710
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Image imgContactsTitle 
      Height          =   705
      Left            =   0
      Picture         =   "frmContacts.frx":57C95
      Top             =   0
      Visible         =   0   'False
      Width           =   4590
   End
   Begin VB.Image imgAdvertisement 
      Height          =   675
      Left            =   0
      Picture         =   "frmContacts.frx":625BF
      Top             =   4560
      Width           =   4590
   End
   Begin VB.Line lnAdSeperator 
      X1              =   0
      X2              =   4650
      Y1              =   4545
      Y2              =   4545
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   4545
      Width           =   4590
   End
   Begin VB.Menu mnuISPN 
      Caption         =   "&Client"
      Begin VB.Menu mnuOptions 
         Caption         =   "ISPN Client &Options.."
      End
      Begin VB.Menu mnuServiceOptions 
         Caption         =   "Setup ISPN &Services.."
      End
      Begin VB.Menu mnuOrganiseContacts 
         Caption         =   "Organise &Contacts.."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogon 
         Caption         =   "Logon"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Logoff"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Visible         =   0   'False
      Begin VB.Menu mnuViewProfile 
         Caption         =   "View User &Profile and Info"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIM2 
         Caption         =   "Send Instant &Message.."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuiMail2 
         Caption         =   "Use as &iMail recipient.."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrganiseContacts2 
         Caption         =   "Organise &Contacts.."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Services"
      Begin VB.Menu mnuBrowse 
         Caption         =   "&Browse ISPN community"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuiM 
         Caption         =   "Send Instant &Message"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Edit my &Profile"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuiMail 
         Caption         =   "Open &iMail Console"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAlerts 
         Caption         =   "Customise &Alerts"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDebugRefresher 
         Caption         =   "Refresh Contacts"
         Visible         =   0   'False
      End
      Begin VB.Menu Dash10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrganiseContacts3 
         Caption         =   "&Organise Contacts"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuWWW 
         Caption         =   "Help on the &Web"
         Begin VB.Menu mnuWebLinkDeveloper 
            Caption         =   "ISPN &developer web page..."
         End
         Begin VB.Menu mnuWebLinkDarkSide 
            Caption         =   "NST &home - [DS]Dark_Side..."
         End
         Begin VB.Menu mnuWebLinkAuthor 
            Caption         =   "About the &author..."
         End
      End
      Begin VB.Menu Dash9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Help &Topics and Tutorials"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About ISPN Client"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuTrayMenu 
      Caption         =   "&TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckInbox 
         Caption         =   "Check iMail &Inbox (0 New)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIM3 
         Caption         =   "Send Instant &Message"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBrowseISPN 
         Caption         =   "&Browse ISPN community"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogon2 
         Caption         =   "&Logon.."
      End
      Begin VB.Menu mnuLogoff2 
         Caption         =   "Log&off.."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggleVisible 
         Caption         =   "&Hide Client"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit ISPN Client"
      End
      Begin VB.Menu mnuCancelMenu 
         Caption         =   "(&C a n c e l  p o p u p)"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Dim ISPN_WWWLINK_DEVELOPER  As String
Dim ISPN_WWWLINK_NSTHOME  As String
Dim ISPN_WWWLINK_AUTHOR  As String

Private Sub Client_Close()
'Close the connection - Dont ask me why this is neccessary, but it is :-p
Client.Close

If modISPN.ISPN_LoggedOn = True Then
    
    If HideTerminateMsg = False Then MsgBox "The connection with the remote server was lost.", 48
    HideTerminateMsg = False
    frmLogin.Show
    modISPN.ISPN_LoggedOn = False ' Not logged on anymore if connection is closed!

Else
    
    frmLogonStatus.lblStatus = "Authentication failed."
    frmLogonStatus.ForeColor = vbRed
    frmLogonStatus.tmrClose.Enabled = True
    
    With frmLogin
        'Re-enable controls
       If Not .Visible = True Then .Show
        .txtUserHandle.Enabled = True
        .txtPassword.Enabled = True
        .txtPIN.Enabled = True
        .txtServerIP.Enabled = True
        .txtPort.Enabled = True
        .cmdLogon.Enabled = True
        .chkAutoLogon.Enabled = True
    End With

End If

'frmClient.lblStatus = "Not connected \ Ready."

'Sort out controls
modISPN.SetControlStates


End Sub

Private Sub Client_Connect()
'We have successfully established a connection with the remote server
'
'This however does not mean we have been authenticated, so we send a authentication
'request (prefix $1). This request contains the user handle, password and serverpin
'
'We also update the logon indicator form here to show the user that the server was contacted
'and a logon process was initiated.

'Update status label
frmLogonStatus.lblStatus = "Host found. Logging on..."

'Send authentication request data with split characters
Client.SendData "$1" & ISPN_LocalHandle & Chr(11) & ISPN_LocalPassword & Chr(11) & ISPN_GivenPIN & Chr(11)

'frmClient.lblStatus = "Host found. Logging on..."
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim p As String

'Stores the data sent to us in a variable, in this example: 'p'
Client.GetData p


'Protocol ISPN1 (Client Side -> Server Side)

' | denotes a split character. This is Chr(11)

'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' $1        Normal Logon Request        UserHandle As String|Password As String|ServerPIN As Long|
' $2        Server Echo Logon Requset   UserHandle As String|Password As String|ServerPIN As Long|AdminPIN As Long|
' $3        Contact Information Request |
' %1        Instant Message Service     Recipient As String|MessageBody As String|
' %2        Profile Service (1\3)       UserToQuery As String| (Profile Information)
' %3        Profile Service (2\3)       ProfileBody as String|ProfilePic as Integer (1 or 0)|PhotoFileName as String
' %6        Profile Service (3\3)       UserToQuery As String| (General Information)

' %4        iMail Service (1\2)         |
' %5        iMail Service (2\2)         Recipient As String|MessageBody As String|Attachment As Integer (1 or 0)|AttachedFileName as String|



'Protocol ISPN1 (Server Side -> Client Side) (This data is handled in this sub!)

' | denotes a split character. This is Chr(11)

'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' $1        Logon Type (1) Validated    (No Parameters)
' $3        Contact Information Data    NumberOfContacts As Integer|[1 = online; 2 = offline; 3 = admin online; 4 = admin offline]HandleShortName as String|
' $4        Contact Information Cancel  |
' &1        Alert (User Logged On)      UserHandle As String|
' &2        Alert (User Logged Off)     UserHandle As String|
' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)
' &4        Alert (New iMail)           |
' %1        Instant Message Service     Recipient As String|MessageBody As String|
' %2        Profile Service (1\3)       UserHandle As String|ProfileBody As String| (Profile Push)
' %3        Profile Service (2\3)       | (Update Successful)
' %6        Profile Service (3\3)       UserHandle As String|AccountMade As String|LogonCount As String| (General Information Push)


Dim tmpAlertTxt As String
Dim tmpAlertShl As String

Select Case Left(p, 2)

    Case "$1" ' $1        Logon Type (1) Validated    (No Parameters)

    
        frmLogonStatus.lblStatus = "Connected to remote ISPN server."
        frmLogonStatus.lblStatus.ForeColor = &HC000&
        
        'Set session variable
        modISPN.ISPN_LoggedOn = True
        
        'Save the logon data that was used to connect
        'so it can be default for next time
        SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_ServerIP", modISPN.ISPN_AddressAttachedTo
        SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_ServerPort", modISPN.ISPN_PortAttachedTo
        SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_UserHandle", modISPN.ISPN_LocalHandle
        SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_LocalPassword", modISPN.ISPN_LocalPassword
        SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_ServerPIN", modISPN.ISPN_GivenPIN
        SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_AutoLogon", frmLogin.chkAutoLogon.Value


        'kill the logon form
        Unload frmLogin
        
        'start timer on logon indicator form that closes the form after a specified time.
        frmLogonStatus.tmrClose.Enabled = True
            
        'enable\disable relative menu items
        modISPN.SetControlStates
    
    Case "%1" ' %1        Instant Message Service     From As String|MessageBody As String|
    
        'Extract information from datagram
        Dim tmpFrom As String
        Dim tmpBody As String
        
        tmpFrom = modIMHandler.ExtractIMData(Right(p, Len(p) - 2), modIMHandler.IM_RECIPIENT)
        tmpBody = modIMHandler.ExtractIMData(Right(p, Len(p) - 2), modIMHandler.IM_BODY)
        
        'Parse IM
        modIMHandler.DisplayIM_VirtualWindow TrimHandle(tmpFrom) & GetAttachedServer(True), tmpBody, , IM_ACTION_ADDTEXT
    
    Case "%6" ' %6        Profile Service (3\3)       UserHandle As String|AccountMade As String|LogonCount As String| (General Information Push)
        
        'Update the general information for 'UserHandle'
        modProfileHandler.DisplayGeneralInfo modProfileHandler.ExtractGeneralInfoData(Right(p, Len(p) - 2), modProfileHandler.ISPN_INFO_TARGET), modProfileHandler.ExtractGeneralInfoData(Right(p, Len(p) - 2), modProfileHandler.ISPN_INFO_ACCOUNTCREATED), modProfileHandler.ExtractGeneralInfoData(Right(p, Len(p) - 2), modProfileHandler.ISPN_INFO_LOGONCOUNT)
        
    Case "$3" ' $3        Contact Information Data    NumberOfContacts As Integer|[1 = online; 2 = offline; 3 = admin online; 4 = admin offline]HandleShortName as String|

        modISPN.UpdateContacts Right(p, Len(p) - 2)
    
    Case "$4" ' $4        Contact Information Cancel  |
    
        'MsgBox "No contact information was available! Hit Action, Refresh Contacts to get an update!", 48, "Debug"
        lblWait.Visible = False
        imgWait.Visible = False
        lstContacts.Left = 0
        lstContacts.Visible = True
        lstContacts.Enabled = True
        frmClient.lstContacts.ListItems.Clear
        frmClient.lstContacts.ListItems.Add , "Browser", "Browse " & frmClient.Client.RemoteHostIP, , 5


    Case "&1" ' &1        Alert (User Logged On)      UserHandle As String|
    
        'Check if alert is turned on
        retval = GetSetting("Nexis Software Technologies", "ISPN_Server", "Alerts_User logon alert", 1)
        
        If retval = 1 Then

        'Extract data fields and execute
        tmpAlertTxt = ExtractAlertData(Right(p, Len(p) - 2), ISPN_ALERT_USERLOGON)
        frmAlerter.ShowAlert "'" & tmpAlertTxt & "' just logged on", ISPN_ALERT_USERLOGON, tmpAlertTxt
    
        End If
        
        'Enable that IM window if open
        modIMHandler.DisplayIM_VirtualWindow tmpAlertTxt, , , IM_ACTION_ENABLE
                
    Case "&2" ' &2        Alert (User Logged Off)     UserHandle As String|
        
        'Check if alert is turned on
        retval = GetSetting("Nexis Software Technologies", "ISPN_Server", "Alerts_User logoff alert", 1)
        
        If retval = 1 Then

        'Extract data fields and execute
        tmpAlertTxt = ExtractAlertData(Right(p, Len(p) - 2), ISPN_ALERT_USERLOGOFF)
        frmAlerter.ShowAlert "'" & tmpAlertTxt & "' just logged off", ISPN_ALERT_USERLOGOFF, tmpAlertTxt
    
        End If
            
        'Disable that IM window if open
        modIMHandler.DisplayIM_VirtualWindow tmpAlertTxt, , , IM_ACTION_DISABLE
    
    Case "&3" ' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)

        'Check if alert is turned on
        retval = GetSetting("Nexis Software Technologies", "ISPN_Server", "Alerts_I do not want to recieve network alerts.", 0)
        
        If Not retval = 1 Then

        'Extract data fields and execute
        tmpAlertTxt = ExtractAlertData(Right(p, Len(p) - 2), ISPN_ALERT_UNSPECIFIED, ISPN_ALERT_ALERTTEXT)
        tmpAlertShl = ExtractAlertData(Right(p, Len(p) - 2), ISPN_ALERT_UNSPECIFIED, ISPN_ALERT_ALERTSHELL)
        frmAlerter.ShowAlert tmpAlertTxt, ISPN_ALERT_UNSPECIFIED, tmpAlertShl
 
        End If
        
    Case "&4" ' &4        Alert (New iMail)           |
    
        'Check if alert is turned on
        retval = GetSetting("Nexis Software Technologies", "ISPN_Server", "Alerts_New iMail alert", 1)
        
        If retval = 1 Then
        
        'Show alert
        frmAlerter.ShowAlert "You have new iMail messages!", ISPN_ALERT_NEWIMAIL
    
        End If
        
    Case Else 'Dont know how to handle this data as it is not drawn out in the standard ISPN 1 Protocol
    
        If ISPN_LoggedOn = False Then
            
            frmLogonStatus.lblStatus = "Incorrect server protocol!"
            frmLogonStatus.tmrClose.Enabled = True
            
            With frmLogin
                'Re-enable controls
               If Not .Visible = True Then .Show
                .txtUserHandle.Enabled = True
                .txtPassword.Enabled = True
                .txtPIN.Enabled = True
                .txtServerIP.Enabled = True
                .txtPort.Enabled = True
                .cmdLogon.Enabled = True
                .chkAutoLogon.Enabled = True
            End With
    
        Else
        
            If MsgBox("Unexpected data from server! The server you are attached to may be incompatible with this client version." & vbNewLine & vbNewLine & "Do you want to stay connected to this server?", 48 + vbYesNo) = vbNo Then HideTerminateMsg = True: Client.Close
        
        End If
    
End Select

End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim tmpErrMsgBox As String
Dim tmpErrLgnFrm As String

'The following Winsock Errors are handled in this proceedure
'
'sckBadState 40006 Wrong protocol or connection state for the requested transaction or request.
'sckInvalidArg 40014 The argument passed to a function was not in the correct format or in the specified range.
'sckSuccess 40017 Successful.
'sckUnsupported 40018 Unsupported variant type.
'sckInvalidOp 40020 Invalid operation at current state
'sckOutOfRange 40021 Argument is out of range.
'sckWrongProtocol 40026 Wrong protocol for the requested transaction or request
'sckOpCanceled 1004 The operation was canceled.
'sckInvalidArgument 10014 The requested address is a broadcast address, but flag is not set.
'sckWouldBlock 10035 Socket is non-blocking and the specified operation will block.
'sckInProgress 10036 A blocking Winsock operation in progress.
'sckAlreadyComplete 10037 The operation is completed. No blocking operation in progress
'sckNotSocket 10038 The descriptor is not a socket.
'sckMsgTooBig 10040 The datagram is too large to fit into the buffer and is truncated.
'sckPortNotSupported 10043 The specified port is not supported.
'sckAddressInUse 10048 Address in use.
'sckAddressNotAvailable 10049 Address not available from the local machine.
'sckNetworkSubsystemFailed 10050 Network subsystem failed.
'sckNetworkUnreachable 10051 The network cannot be reached from this host at this time.
'sckNetReset 10052 Connection has timed out when SO_KEEPALIVE is set.
'sckConnectAborted 11053 Connection is aborted due to timeout or other failure.
'sckConnectionReset 10054 The connection is reset by remote side.
'sckNoBufferSpace 10055 No buffer space is available.
'sckAlreadyConnected 10056 Socket is already connected.
'sckNotConnected 10057 Socket is not connected.
'sckSocketShutdown 10058 Socket has been shut down.
'sckTimedout 10060 Socket has been shut down.
'sckConnectionRefused 10061 Connection is forcefully rejected.
'sckNotInitialized 10093 WinsockInit should be called first.
'sckHostNotFound 11001 Authoritative answer: Host not found.
'sckHostNotFoundTryAgain 11002 Non-Authoritative answer: Host not found.
'sckNonRecoverableError 11003 Non-recoverable errors.
'sckNoData 11004 Valid name, no data record of requested type
'No more results can be returned by WSALookupServiceNext


Dim tmpNumber As Long
tmpNumber = Number


'Error handler
Select Case tmpNumber

    Case sckBadState '40006 Wrong protocol or connection state for the requested transaction or request.
        
        tmpErrMsgBox = "Invalid socket state." & vbNewLine & vbNewLine & "Please re-start ISPN client \ your computer. (Error: " & Number & ")"
        tmpErrLgnFrm = "Invalid socket state. (" & Number & ")"
    
    Case sckInvalidArg '40014 The argument passed to a function was not in the correct format or in the specified range.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "The argument passed to a function was not in the correct format or in the specified range. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "Invalid socket state. (" & Number & ")"
    
    Case sckSuccess '40017 Successful.
    
    Case sckUnsupported '40018 Unsupported variant type.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "Unsupported variant type. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "Unsupported variant type. (" & Number & ")"
    
    Case sckInvalidOp '40020 Invalid operation at current state
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "Invalid operation at current state. Please re-start ISPN client \ your computer. (Error: " & Number & ")"
        tmpErrLgnFrm = "UInvalid operation at current state. (" & Number & ")"

    Case sckOutOfRange '40021 Argument is out of range.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "Argument is out of range. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "Argument is out of range.. (" & Number & ")"
    
    Case sckWrongProtocol '40026 Wrong protocol for the requested transaction or request
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "Wrong protocol for the requested transaction or request. Please re-start ISPN client \ your computer. (Error: " & Number & ")"
        tmpErrLgnFrm = "Unsupported variant type.. (" & Number & ")"
    
    Case sckOpCanceled '1004 The operation was canceled.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "The operation was canceled. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "The operation was canceled. (" & Number & ")"
    
    Case sckInvalidArgument '10014 The requested address is a broadcast address, but flag is not set.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "The requested address is a broadcast address, but flag is not set. Check your ISPN connection settings . (Error: " & Number & ")"
        tmpErrLgnFrm = "Unsupported variant type.. (" & Number & ")"
    
    Case sckWouldBlock '10035 Socket is non-blocking and the specified operation will block.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "Unsupported variant type.. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "Unsupported variant type.. (" & Number & ")"
    
    Case sckInProgress '10036 A blocking Winsock operation in progress.
        
        tmpErrMsgBox = "ActiveX Error." & vbNewLine & vbNewLine & "A blocking Winsock operation in progress. Another process is currently using resources required for this operation. Please try again later or restart your computer. (Error: " & Number & ")"
        tmpErrLgnFrm = "Winsock is in use by another process (" & Number & ")"
    
    Case sckAlreadyComplete '10037 The operation is completed. No blocking operation in progress
            
        'Do not need to inform user of this event
    
    Case sckNotSocket '10038 The descriptor is not a socket.
        
        tmpErrMsgBox = "Connection Error." & vbNewLine & vbNewLine & "The descriptor is not a socket. Please check your ISPN connection settings. (Error: " & Number & ")"
        tmpErrLgnFrm = "Server socket error. (" & Number & ")"
    
    Case sckMsgTooBig '10040 The datagram is too large to fit into the buffer and is truncated.
        
        tmpErrMsgBox = "Winsock error." & vbNewLine & vbNewLine & "The datagram is too large to fit into the buffer and is truncated. The action you requested could not be completed successfully. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "Datagram is too large for buffer. (" & Number & ")"
    
    Case sckPortNotSupported '10043 The specified port is not supported.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "The specified port is not supported. Please check your ISPN connection settings. You typed a port that is un-supported. (Error: " & Number & ")"
        tmpErrLgnFrm = "The specified port is invalid. (" & Number & ")"
    
    Case sckAddressInUse '10048 Address in use.
        
        tmpErrMsgBox = "Winsock error." & vbNewLine & vbNewLine & "Address in use. Another program is currently using the port that ISPN requires. Please terminate that application and re-start ISPN client. (Error: " & Number & ")"
        tmpErrLgnFrm = "Local address already in use. (" & Number & ")"
    
    Case sckAddressNotAvailable '10049 Address not available from the local machine.
        
        tmpErrMsgBox = "Winsock error." & vbNewLine & vbNewLine & "Address not available from the local machine. Unable to generate a local address for ISPN client Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "Local address assignment failed. (" & Number & ")"
    
    Case sckNetworkSubsystemFailed '10050 Network subsystem failed.
        
        tmpErrMsgBox = "External error." & vbNewLine & vbNewLine & "Network subsystem failed. Your network connection may be configured incorrectly. Please contact your system administrator. (Error: " & Number & ")"
        tmpErrLgnFrm = "Network subsystem failed. (" & Number & ")"
    
    Case sckNetworkUnreachable '10051 The network cannot be reached from this host at this time.
        
        tmpErrMsgBox = "External error." & vbNewLine & vbNewLine & "The network cannot be reached from this host at this time. Your network connection may be configured incorrectly. Please contact your system administrator. (Error: " & Number & ")"
        tmpErrLgnFrm = "The network cannot be reached from this host at this time. (" & Number & ")"
    
    Case sckNetReset '10052 Connection has timed out when SO_KEEPALIVE is set.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Connection has timed out when SO_KEEPALIVE is set. A connection request was timed out when time out was disabled. Please try re-connecting. (Error: " & Number & ")"
        tmpErrLgnFrm = "Connection was forceably timed out. (" & Number & ")"
    
    Case sckConnectAborted '11053 Connection is aborted due to timeout or other failure.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Connection was aborted due to timeout or other failure. Unspecified connection error. Please contact the ISPN administrator and \ or check your network configuration. (Error: " & Number & ")"
        tmpErrLgnFrm = "Connection was aborted due to timeout or other failure. (" & Number & ")"
    
    Case sckConnectionReset '10054 The connection is reset by remote side.
        
        tmpErrMsgBox = "Connection Lost." & vbNewLine & vbNewLine & "The connection was reset by remote side. The remote server reset the connection. You may lose your current ISPN connection. (Error: " & Number & ")"
        tmpErrLgnFrm = "Server reset connection. (" & Number & ")"
    
    Case sckNoBufferSpace '10055 No buffer space is available.
        
        tmpErrMsgBox = "Resource allocation error." & vbNewLine & vbNewLine & "No buffer space is available. You may need to increase your computer's physical memory, or contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "No buffer space is available. (" & Number & ")"
    
    Case sckAlreadyConnected '10056 Socket is already connected.
        
        tmpErrMsgBox = "Winsock error." & vbNewLine & vbNewLine & "Socket is already connected. An attempt to logon whilst already logged on was made. You may still be able to continue. (Error: " & Number & ")"
        tmpErrLgnFrm = "Socket is already connected. (" & Number & ")"

    Case sckNotConnected '10057 Socket is not connected.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Socket is not connected. Unable to connect to the remote computer. Please try again later. (Error: " & Number & ")"
        tmpErrLgnFrm = "Unable to connect socket. (" & Number & ")"
    
    Case sckSocketShutdown '10058 Socket has been shut down.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Socket has been shut down. The server or a local process has shut down the active connection. (Error: " & Number & ")"
        tmpErrLgnFrm = "Socket was shut down. (" & Number & ")"
    
    Case sckTimedout '10060 Socket has been shut down.
                
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Socket has been shut down. The server or a local process has shut down the active connection. (Error: " & Number & ")"
        tmpErrLgnFrm = "Socket was shut down. (" & Number & ")"
    
    Case sckConnectionRefused '10061 Connection is forcefully rejected.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "Connection is forcefully rejected. The IP Address specified is not a qualifying ISPN server. Please specify a valid server address. (Error: " & Number & ")"
        tmpErrLgnFrm = "Invalid ISPN server! (" & Number & ")"

    Case sckNotInitialized '10093 WinsockInit should be called first.
        
        tmpErrMsgBox = "Internal error." & vbNewLine & vbNewLine & "WinsockInit should be called first. Please contact your vendor. (Error: " & Number & ")"
        tmpErrLgnFrm = "WinsockInit should be called first. (" & Number & ")"
    
    Case sckHostNotFound '11001 Authoritative answer: Host not found.
        
        tmpErrMsgBox = "Connection Error." & vbNewLine & vbNewLine & "Authoritative answer: Host not found. Please check you ISPN settings. (Error: " & Number & ")"
        tmpErrLgnFrm = "Authoritative answer: Host not found. (" & Number & ")"
    
    Case sckHostNotFoundTryAgain '11002 Non-Authoritative answer: Host not found.
        
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Non-Authoritative answer: Host not found. Please check you ISPN settings. (Error: " & Number & ")"
        tmpErrLgnFrm = "Non-Authoritative answer: Host not found. (" & Number & ")"
    
    Case sckNonRecoverableError '11003 Non-recoverable errors.
        
        tmpErrMsgBox = "Unknown error." & vbNewLine & vbNewLine & "Non-recoverable errors. Please restart ISPN client and \ or your computer. (Error: " & Number & ")"
        tmpErrLgnFrm = "Non-recoverable error(s)! (" & Number & ")"
        
    Case sckNoData '11004 Valid name, no data record of requested type
       
        tmpErrMsgBox = "Connection error." & vbNewLine & vbNewLine & "Valid name, no data record of requested type. The server was unable to carry out the requested action. (Error: " & Number & ")"
        tmpErrLgnFrm = "Server protocol error. (" & Number & ")"

    
    Case 10061 ' "Connection is forcefully rejected: 10061"
        
        tmpErrMsgBox = "The remote server could not be located." & vbNewLine & vbNewLine & "Please check the Server IP Address and try again. (Error: " & Number & ")"
        tmpErrLgnFrm = "The remote server could not be located!. (" & Number & ")"
        
    Case 10110 '"No more results can be returned by WSALookupServiceNext" DNS could not be resolved
            
        tmpErrMsgBox = "The server DNS you specified could not be resolved." & vbNewLine & vbNewLine & "Try typing the IP Address rather than the domain name. The ISPN administrator will be able to supply you with this information. (Error: " & Number & ")"
        tmpErrLgnFrm = "The server DNS you specified could not be resolved. (" & Number & ")"
    
    Case Else 'Any other error
            
        tmpErrMsgBox = "ISPN 1 Connection Error. You may be disconnected from the service. (" & Description & " - " & Number & ")"
        tmpErrLgnFrm = "ISPN 1.0 Error (" & Description & ")"
        
End Select

If tmpErrMsgBox = "" Then Exit Sub
If tmpErrLgnFrm = "" Then Exit Sub

'This will display the error to the user
If ISPN_LoggedOn = True Then
       
    MsgBox tmpErrMsgBox, 48, "ISPN Client Error Handler"
        
Client.Close

Else

    'Update the logon indicator
    frmLogonStatus.lblStatus = tmpErrLgnFrm
    frmLogonStatus.lblStatus.ForeColor = vbRed
    frmLogonStatus.tmrClose.Enabled = True
    
    With frmLogin
        'Re-enable controls
        If Not .Visible = True Then .Show
        .txtUserHandle.Enabled = True
        .txtPassword.Enabled = True
        .txtPIN.Enabled = True
        .txtServerIP.Enabled = True
        .txtPort.Enabled = True
        .cmdLogon.Enabled = True
        .chkAutoLogon.Enabled = True
    End With

    'Show\Hide Controls on the main client window
    frmClient.fraLogon.Visible = True
    frmClient.imgWait.Visible = False
    frmClient.lblWait.Visible = False
    
    'Enable\Disable menu items
    frmClient.mnuLogon.Enabled = True
    frmClient.mnuLogoff2.Enabled = True

Client.Close

End If


End Sub

Private Sub Form_Load()
'Set Web Links Variables
ISPN_WWWLINK_NSTHOME = "http://www.angelfire.com/d20/vbfiles"
ISPN_WWWLINK_AUTHOR = "about:<TITLE>ISPN on the web</TITLE>This page has not been added yet. Please visit NST on the web <A HREF=" & Chr(34) & ISPN_WWWLINK_NSTHOME & Chr(34) & ">here</A>"
ISPN_WWWLINK_DEVELOPER = "about:<TITLE>ISPN on the web</TITLE>This page has not been added yet. Please visit NST on the web <A HREF=" & Chr(34) & ISPN_WWWLINK_NSTHOME & Chr(34) & ">here</A>"

'Sorts out all the controls for the relavent logon state (NotLoggedOn)
SetControlStates

'Add system tray icon
PROBas.SystemTrayAddIcon frmClient, frmIconContainer

'Set startup position
frmClient.Height = 5880 'Sometimes the height gets clipped by Systray
frmClient.Top = Screen.Height - (frmClient.Height + 400) '('X' Co-ordinate)
frmClient.Left = Screen.Width - frmClient.Width '('Y' Co-ordinate)

If Not Command = "" Then Me.Hide

'Autologon Routine
chkAutoLogon.Value = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_AutoLogon", frmLogin.chkAutoLogon.Value)

If chkAutoLogon.Value = 1 Then
    
    Dim tmpUserHandle As String
    Dim tmpPassword As String
    Dim tmpServerIP As String
    Dim tmpPIN As String
    Dim tmpPort As String
    
    'Load default logon data from user specific registry key
    tmpUserHandle = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_UserHandle")
    tmpPassword = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_LocalPassword")
    tmpServerIP = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_ServerIP")
    tmpPIN = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_ServerPIN")
    tmpPort = GetSetting("Nexis Software Technologies", "ISPN", "LogonDefault_ServerPort")
    
    'Check for null variables
    If CheckNull(tmpUserHandle) = True Then tmpAutoFail = 1
    If CheckNull(tmpPassword) = True Then tmpAutoFail = 1
    If CheckNull(tmpServerIP) = True Then tmpAutoFail = 1
    If CheckNull(tmpPIN) = True Then tmpAutoFail = 1
    If CheckNull(tmpPort) = True Then tmpAutoFail = 1
    
    'Check for type mismatches
    If IsNumeric(tmpPort) = False Then tmpAutoFail = 1
    If IsNumeric(tmpPIN) = False Then tmpAutoFail = 1
    
    'Failed variable checks
    If tmpAutoFail = 1 Then MsgBox "ISPN Client was unable to log you onto the server automatically." & vbNewLine & vbNewLine & " > Autologon failed because one or more required fields were invalid or missing", 48, "AutoLogon Failure": SaveSetting "Nexis Software Technologies", "ISPN", "LogonDefault_AutoLogon", 0: frmLogin.Show: chkAutoLogon.Value = 0: Exit Sub
    
    'Save variables
    modISPN.ISPN_LocalHandle = tmpUserHandle
    modISPN.ISPN_GivenPIN = tmpPIN
    modISPN.ISPN_LocalPassword = tmpPassword
    modISPN.ISPN_AddressAttachedTo = tmpServerIP
    modISPN.ISPN_PortAttachedTo = tmpPort
    
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


Else

    
    
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Static lngMsg As Long
Static blnFlag As Boolean
Dim result As Long

lngMsg = X / Screen.TwipsPerPixelX

If blnFlag = False Then
    
    blnFlag = True
        
    Select Case lngMsg
        
        Case WM_LBUTTONDBLCLICK 'Double Click
        
            Me.Show
        
        Case WM_RBUTTONUP 'Right Button
        
            PopupMenu mnuTrayMenu, , , , mnuToggleVisible
            result = SetForegroundWindow(Me.hWnd)
    
    End Select
    
    blnFlag = False

End If
End Sub

Private Sub Form_Resize()
If frmClient.WindowState = 1 Then mnuToggleVisible_Click: frmClient.WindowState = 0: frmClient.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1 'Cancel the operation
mnuToggleVisible_Click
End Sub



Private Sub hlnk_Logon_Click()
'If chkAutoLogon.Value = 1 Then

mnuLogon_Click
End Sub


Private Sub lstContacts_DblClick()
'Executes the default action

For f = 1 To lstContacts.ListItems.Count
    If lstContacts.ListItems(f).Selected = True Then n = n + 1
Next f

If n = 0 Then Exit Sub 'No item is selected

On Error Resume Next
If lstContacts.SelectedItem.Key = "Browser" Then
    
    ShowToDo

Else
    
    'Execute the default action
    modISPN.PerfomDefaultAction lstContacts.SelectedItem.text

End If
End Sub

Private Sub lstContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set lstContacts.SelectedItem = Nothing
End Sub

Private Sub lstContacts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Check if a contact is selected and display menu accordingly.
    
If Button = 2 Then
    For f = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(f).Selected = True Then n = n + 1
    Next f
    
    If n = 0 Then Exit Sub 'No item is selected
    
    'If n > 1 Then
    '    mnuRenameServerFile.Enabled = False
    'Else
    '    mnuRenameServerFile.Enabled = True
    'End If
    
    On Error Resume Next
    If lstContacts.SelectedItem.Key = "Browser" Then
        Call Me.PopupMenu(mnuAction, , , , mnuBrowse)
    Else
        Call Me.PopupMenu(mnuUser, , , , mnuIM2)
    End If
    
    
End If

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal, frmClient
End Sub

Private Sub mnuAlerts_Click()
frmAlertSettings.Show
End Sub

Private Sub mnuDebugRefresher_Click()
lstContacts.ListItems.Clear
Client.SendData "$3" & Chr(11)
End Sub

Private Sub mnuExit_Click()
PROBas.SystemTrayDeleteIcon frmClient
End
End Sub

Private Sub mnuHide_Click()
Unload Me
End Sub

Private Sub mnuIM2_Click()
'Compose an instant message to selected contact
modIMHandler.DisplayIM_VirtualWindow modISPN.TrimHandle(lstContacts.SelectedItem.text) & GetAttachedServer(True)


End Sub
Private Sub mnuLogoff_Click()
If ISPN_LoggedOn = True Then
    
    'If MsgBox("You have one or more open windows." & vbNewLine & "This action will cancel any communication with other users and the server." & vbNewLine & vbNewLine & "Continue?", vbYesNo + 48, "Logoff ISPN Server") = vbNo Then Exit Sub

'|-Logoff--------------
    Client.Close
    ISPN_LoggedOn = False
    SetControlStates
    'frmLogin.Show
Else
    
    SetControlStates

End If
End Sub

Private Sub mnuLogoff2_Click()
mnuLogoff_Click
End Sub

Private Sub mnuLogon_Click()
frmLogin.Show 'Show the login dialog
End Sub

Private Sub mnuLogon2_Click()
mnuLogon_Click
End Sub

Private Sub mnuToggleVisible_Click()
frmClient.Visible = Not frmClient.Visible
If frmClient.Visible = False Then mnuToggleVisible.Caption = "Show Client" Else mnuToggleVisible.Caption = "Hide Client"
End Sub

Private Sub mnuViewProfile_Click()
'Shows the User Informatiion Screen, where the profile can be viewed via a link..
modProfileHandler.DisplayGeneralInfo (TrimHandle(lstContacts.SelectedItem.text)), , , True
End Sub

Private Sub mnuWebLinkAuthor_Click()
OpenIt frmClient, ISPN_WWWLINK_AUTHOR
End Sub

Private Sub mnuWebLinkDarkSide_Click()
OpenIt frmClient, ISPN_WWWLINK_NSTHOME
End Sub

Private Sub mnuWebLinkDeveloper_Click()
OpenIt frmClient, ISPN_WWWLINK_DEVELOPER
End Sub

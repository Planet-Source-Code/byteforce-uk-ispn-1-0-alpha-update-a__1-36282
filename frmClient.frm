VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISPN Client 1.0 [Alpha Version]"
   ClientHeight    =   3810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClient.frx":038A
   ScaleHeight     =   3810
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogOff 
      Caption         =   "Log &off"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3555
      TabIndex        =   7
      Top             =   3360
      Width           =   1035
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   90
      Top             =   3345
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraServer 
      Caption         =   "Client"
      Height          =   2295
      Left            =   75
      TabIndex        =   0
      Top             =   1005
      Width           =   4545
      Begin VB.CommandButton cmdSendIM 
         Caption         =   "Cast IM"
         Height          =   405
         Left            =   105
         TabIndex        =   8
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "ISPN 1.1 Status:"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label lblStatus 
         Caption         =   "Not connected \ Ready."
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1455
         TabIndex        =   1
         Top             =   285
         Width           =   2865
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2730
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISPN Client 1.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0 ALPHA"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   615
      Width           =   1425
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NST "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   4695
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "www.angelfire.com/d20/vbfiles"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   2250
      TabIndex        =   2
      Top             =   600
      Width           =   2445
   End
   Begin VB.Menu mnuISPN 
      Caption         =   "&Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCommunity 
         Caption         =   "Services"
         Enabled         =   0   'False
         Begin VB.Menu mnuIMService 
            Caption         =   "Instant Messaging"
            Begin VB.Menu mnuIMEnable 
               Caption         =   "&Enable this service"
               Checked         =   -1  'True
            End
            Begin VB.Menu Dash5 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSendIM 
               Caption         =   "Send an Instant Message"
            End
            Begin VB.Menu mnuIMSettings 
               Caption         =   "Service Settings"
            End
         End
         Begin VB.Menu mnuProfiles 
            Caption         =   "Profiles"
            Enabled         =   0   'False
            Begin VB.Menu mnuProfilesEnable 
               Caption         =   "&Enable this service"
               Checked         =   -1  'True
            End
            Begin VB.Menu Dash6 
               Caption         =   "-"
            End
            Begin VB.Menu mnuEditProfile 
               Caption         =   "Edit my profile"
            End
         End
      End
      Begin VB.Menu mnuISPN2 
         Caption         =   "ISPN"
         Enabled         =   0   'False
         Begin VB.Menu munDirectory 
            Caption         =   "User directory \ find contacts"
         End
         Begin VB.Menu Dash3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComposeiMail 
            Caption         =   "Compose iMail"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuReadiMail 
            Caption         =   "Read iMail"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogon 
         Caption         =   "&Logon"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Log &off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsole 
         Caption         =   "Open console"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu Dash8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "(Cancel Menu)"
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


Private Sub Client_Close()
'Close the connection - Dont ask me why this is neccessary, but it is :-p
Client.Close

If modISPN.ISPN_LoggedOn = True Then
    
    If HideTerminateMsg = False Then MsgBox "The connection with the remote server was lost.", 48
    HideTerminateMsg = False
    frmClient.Hide
    frmLogin.Show
    modISPN.ISPN_LoggedOn = False ' Not logged on anymore if connection is closed!

Else
    
    frmLogonStatus.lblStatus = "Authentication failed."
    frmLogonStatus.tmrClose.Enabled = True
    
    With frmLogin
        'Re-enable controls
        .txtUserHandle.Enabled = True
        .txtPassword.Enabled = True
        .txtPIN.Enabled = True
        .txtServerIP.Enabled = True
        .txtPort.Enabled = True
        .cmdLogon.Enabled = True
        .chkAutoLogon.Enabled = True
    End With

End If

frmClient.lblStatus = "Not connected \ Ready."

'Sort out tray menu
modISPN.DoMenuItems


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

frmClient.lblStatus = "Host found. Logging on..."
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim p As String

'Stores the data sent to us in a variable, in this example: 'p'
Client.GetData p

'Protocol ISPN1
Select Case Left(p, 2)

    Case "$1" ' $1        Logon Type (1) Validated    (No Parameters)

    
        frmLogonStatus.lblStatus = "Connected to remote ISPN server."
        frmClient.lblStatus = "Connected to remote ISPN server."
        
        'Set session variable
        modISPN.ISPN_LoggedOn = True
        
        'Save the logon data that was used to connect
        'so it can be default for next time
        SaveSetting "Nexus Software Technologies", "ISPN", "LogonDefault_ServerIP", modISPN.ISPN_AddressAttachedTo
        SaveSetting "Nexus Software Technologies", "ISPN", "LogonDefault_ServerPort", modISPN.ISPN_PortAttachedTo
        SaveSetting "Nexus Software Technologies", "ISPN", "LogonDefault_UserHandle", modISPN.ISPN_LocalHandle
        SaveSetting "Nexus Software Technologies", "ISPN", "LogonDefault_LocalPassword", modISPN.ISPN_LocalPassword
        SaveSetting "Nexus Software Technologies", "ISPN", "LogonDefault_ServerPIN", modISPN.ISPN_GivenPIN
        SaveSetting "Nexus Software Technologies", "ISPN", "LogonDefault_AutoLogon", frmLogin.chkAutoLogon.Value


        'kill the logon form
        Unload frmLogin
        
        'show the client form (unless user specified differently)
        frmClient.Show
        frmClient.cmdLogOff.Enabled = True
        
        'start timer on logonstatus form that closes the form after a specified time.
        frmLogonStatus.tmrClose.Enabled = True
            
        'enable\disable relative menu items
        modISPN.DoMenuItems
    
    Case "%1" ' %1        Instant Message Service     From As String|MessageBody As String|
    
        'Extract information from datagram
        Dim tmpFrom As String
        Dim tmpBody As String
        
        tmpFrom = modIMHandler.ExtractIMData(Right(p, Len(p) - 2), modIMHandler.IM_RECIPIENT)
        tmpBody = modIMHandler.ExtractIMData(Right(p, Len(p) - 2), modIMHandler.IM_BODY)
        
        'Display in relavent IM window
        frmIM.DisplayIM tmpFrom, Client.RemoteHostIP, tmpBody
        
    
    Case Else 'Dont know how to handle this data
    
        If ISPN_LoggedOn = False Then
            
            frmLogonStatus.lblStatus = "Incorrect server protocol!"
            frmLogonStatus.tmrClose.Enabled = True
            
            With frmLogin
                'Re-enable controls
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
       
    MsgBox tmpErrMsgBox, 48
        
Else

    frmLogonStatus.lblStatus = tmpErrLgnFrm
    frmLogonStatus.tmrClose.Enabled = True
    
    With frmLogin
        'Re-enable controls
        .txtUserHandle.Enabled = True
        .txtPassword.Enabled = True
        .txtPIN.Enabled = True
        .txtServerIP.Enabled = True
        .txtPort.Enabled = True
        .cmdLogon.Enabled = True
        .chkAutoLogon.Enabled = True
    End With

End If


End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdLogOff_Click()
'Close connection
Client.Close

'kill client screen
frmClient.Hide

'Not logged on anymore so set global var
modISPN.ISPN_LoggedOn = False

'Sort out tray menu
modISPN.DoMenuItems

'Show logon screen
frmLogin.Show
End Sub

Private Sub cmdSendIM_Click()
'frmIM.Show
frmIM.StartIM "user2@" & Client.RemoteHostIP
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This proceedure code courtesy of 'Cold Fusionz.net'

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
    
        If ISPN_LoggedOn = True Then
            
            PopupMenu mnuISPN, , , , mnuCommunity
        
        Else
        
            PopupMenu mnuISPN, , , , mnuLogon
        
        End If
        
    result = SetForegroundWindow(Me.hWnd)
End Select
blnFlag = False
End If
End Sub

Private Sub cmdLogon_Click()
lblStatus = "Connecting to " & txtServerIP & ":" & txtPort
'Disable controls
txtUserHandle.Enabled = False
txtPassword.Enabled = False
txtPIN.Enabled = False
txtServerIP.Enabled = False
txtPort.Enabled = False
cmdLogon.Enabled = False
cmdLogOff.Enabled = False

If Client.State <> 0 Then Client.Close

Client.Connect txtServerIP, txtPort

End Sub

Private Sub Form_Load()
'Add icon to tray
Me.Hide
PROBas.SystemTrayAddIcon frmClient
End Sub

Private Sub Form_Terminate()
'PROBas.SystemTrayDeleteIcon frmClient
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Setting Cancel to 1 will 'cancel' the unload form operation. (duh..)
Cancel = 1

'Hides the frmClient form (Object 'Me' can be used as the current form)
Me.Hide
End Sub

Private Sub mnuCancel_Click()
'Cancels the popup menu.
'No code is needed here, by clicking on any enabled item, the popup menu
'will disappear.
End Sub

Private Sub mnuConsole_Click()
frmClient.Show
End Sub

Private Sub mnuExit_Click()
    If ISPN_LoggedOn = True Then
        If MsgBox("You are still logged on to an ISPN server. This action will log you out from the server! Continue?", vbYesNo + vbCritical) = vbYes Then
            PROBas.SystemTrayDeleteIcon frmClient
            End
        End If
    Else
        PROBas.SystemTrayDeleteIcon frmClient
        End
    End If
End Sub

Private Sub mnuLogoff_Click()
'just calls the code for the logoff command button rather than typing out the same code twice
cmdLogOff_Click
End Sub

Private Sub mnuLogon_Click()
frmLogin.Show
End Sub

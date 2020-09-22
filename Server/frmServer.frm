VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISPN Server 1.0  [Alpha Version]"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "ISPNServer"
   MaxButton       =   0   'False
   Picture         =   "frmServer.frx":0CCA
   ScaleHeight     =   6000
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdShowUsers 
      Caption         =   "&Users"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2520
      TabIndex        =   28
      ToolTipText     =   "View a list of logged on users"
      Top             =   3330
      Width           =   1065
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options.."
      Height          =   390
      Left            =   3660
      TabIndex        =   5
      Top             =   3330
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Activity"
      Height          =   2145
      Left            =   60
      TabIndex        =   16
      Top             =   3795
      Width           =   5865
      Begin VB.ListBox lstLog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1785
         Left            =   90
         TabIndex        =   7
         Top             =   255
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   4785
      TabIndex        =   6
      Top             =   3330
      Width           =   1035
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1155
      TabIndex        =   4
      Top             =   3315
      Width           =   1035
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   60
      TabIndex        =   3
      Top             =   3315
      Width           =   1035
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server"
      Height          =   2280
      Left            =   60
      TabIndex        =   9
      Top             =   975
      Width           =   5850
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Text            =   "210"
         Top             =   1800
         Width           =   915
      End
      Begin VB.TextBox txtPIN 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Text            =   "1528"
         Top             =   1800
         Width           =   915
      End
      Begin MSWinsockLib.Winsock Server 
         Index           =   0
         Left            =   5310
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblNextConnection 
         Caption         =   "0"
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   3315
         TabIndex        =   27
         Top             =   1965
         Width           =   1110
      End
      Begin VB.Label Label12 
         Caption         =   "Listening On:"
         Height          =   240
         Left            =   2115
         TabIndex        =   26
         Top             =   1965
         Width           =   1170
      End
      Begin VB.Label Label13 
         Caption         =   "Topmost Index:"
         Height          =   240
         Left            =   2115
         TabIndex        =   25
         Top             =   1725
         Width           =   1170
      End
      Begin VB.Label lblTopMost 
         Caption         =   "0"
         ForeColor       =   &H00008080&
         Height          =   240
         Left            =   3315
         TabIndex        =   24
         Top             =   1725
         Width           =   1110
      End
      Begin VB.Label lblWaiting 
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1575
         TabIndex        =   23
         Top             =   1200
         Width           =   2865
      End
      Begin VB.Label Label10 
         Caption         =   "Awaiting LgnData:"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Label lblConnected 
         Caption         =   "0"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1575
         TabIndex        =   21
         Top             =   285
         Width           =   2865
      End
      Begin VB.Label lblRefused 
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1575
         TabIndex        =   20
         Top             =   510
         Width           =   2865
      End
      Begin VB.Label lblIPAddress 
         Caption         =   "0.0.0.0"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1575
         TabIndex        =   19
         Top             =   750
         Width           =   2865
      End
      Begin VB.Label lblConnectedNow 
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1575
         TabIndex        =   18
         Top             =   975
         Width           =   2865
      End
      Begin VB.Label Label11 
         Caption         =   "Connected Now:"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   975
         Width           =   1380
      End
      Begin VB.Label Label9 
         Caption         =   "Server Port:"
         Height          =   240
         Left            =   1140
         TabIndex        =   14
         Top             =   1545
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Server IP Address:"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Server PIN:"
         Height          =   240
         Left            =   135
         TabIndex        =   12
         Top             =   1545
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "Users refused:"
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   510
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Users Validated:"
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   285
         Width           =   1380
      End
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmServer.frx":1799
      Stretch         =   -1  'True
      Top             =   0
      Width           =   795
   End
   Begin VB.Label hlnk_DarkSideHome 
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
      Left            =   3450
      MouseIcon       =   "frmServer.frx":DBD7
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Visit the [DS]Dark_Side Internet pages"
      Top             =   570
      Width           =   2445
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   5985
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0 ALPHA"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   15
      TabIndex        =   8
      Top             =   585
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISPN Server 1.0"
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
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   1200
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdShowUsers_Click()
frmUsers.Show
End Sub

Private Sub cmdStart_Click()
If IsNumeric(txtPort) = False Then MsgBox "Invalid port. A port must be a number from 1 to 32667", 16: Exit Sub

txtPIN.Enabled = False
txtPort.Enabled = False

'Check state before listening
If Server(0).State <> 0 Then Server(0).Close

'Set the servers port
Server(0).LocalPort = txtPort

'Reset the ISPN_TopLevelCtlId variable back to 0
ISPN_TopLevelCtlId = 0

'Start listening for connection requests
frmServer.lblNextConnection = 0
frmServer.lblWaiting = 0
frmServer.lblConnected = 0
frmServer.lblConnectedNow = 0
frmServer.lblRefused = 0

Server(0).Close
Server(0).Listen

AddLogItem "Started ISPN 1.0 Alpha SERVER: " & Now
AddLogItem "Opened server array element 0"

cmdStart.Enabled = False
cmdStop.Enabled = True
cmdClose.Enabled = False
cmdShowUsers.Enabled = True
End Sub

Public Sub AddLogItem(txt As String)
'Clear the list box if there are more than 199 entries
If lstLog.ListCount >= 200 Then lstLog.Clear

'Add the item to the running log in lstLog
lstLog.AddItem txt
lstLog.ListIndex = lstLog.ListCount - 1
End Sub

Private Sub cmdStop_Click()
If lblConnectedNow = 1 Then If MsgBox("Are you sure you want to stop the server? There is " & lblConnectedNow & " connected user." & vbNewLine & vbNewLine & "This action will logout all users attached to this server! Continue?", vbYesNo + 48, "Stop ISPN server?") = vbNo Then Exit Sub
If lblConnectedNow > 1 Then If MsgBox("Are you sure you want to stop the server? There are " & lblConnectedNow & " connected users." & vbNewLine & vbNewLine & "This action will logout all users attached to this server! Continue?", vbYesNo + 48, "Stop ISPN server?") = vbNo Then Exit Sub

On Error GoTo errcont

'Close all open connections
'Add item to running log
AddLogItem "Closing connections.."


'Delete virtual winsock control (number Index) from array and reset the variables
For delctl = 1 To ISPN_TopLevelCtlId
    Server(delctl).Close
    Unload Server(delctl)
    ISPNUSER_MemberHandle(delctl) = ""
    ISPNUSER_LoggedOn(delctl) = False
Next

'Now reset the base control
ISPNUSER_MemberHandle(0) = ""
ISPNUSER_LoggedOn(0) = False
Server(0).Close

errcont:

'Add item to running log
AddLogItem "Disconnected all users and stopped server at " & Now

'Reset items
txtPIN.Enabled = True
txtPort.Enabled = True
cmdStop.Enabled = False
cmdStart.Enabled = True
cmdClose.Enabled = True
cmdShowUsers.Enabled = False

lblConnected = 0
lblConnectedNow = 0
lblWaiting = 0
lblRefused = 0

'Variables are reset in cmdStart_Click
End Sub

Private Sub cmdOptions_Click()
frmServiceOptions.Show
End Sub

Private Sub Form_Load()
'Retrieve the local machines IP address as reported by Winsock.
lblIPAddress = Server(0).LocalIP
End Sub

Private Sub hlnk_DarkSideHome_Click()
Shell "C:\program files\internet explorer\iexplore.exe http://www.angelfire.com/d20/vbfiles", vbNormalFocus
End Sub

Public Sub Server_Close(Index As Integer)

If ISPNUSER_MemberHandle(Index) = "" Or ISPNUSER_MemberHandle(Index) = Chr(11) Then 'Dont know who user is or is not logged on
    
    'Add an item to the running log
    AddLogItem "User disconnected before validation completed (" & Now & ")" & "IP: " & Server(Index).RemoteHostIP

Else
    
    'Add an item to the running log
    AddLogItem "'" + ISPNUSER_MemberHandle(Index) + "' disconnected at " & Now & " (IP: " & Server(Index).RemoteHostIP & ")"
    lblConnectedNow = lblConnectedNow - 1
    'Protocol ISPN1 (Server Side -> Client Side)

    ' | signifies a split character. This is Chr(11)
    
    'Prefix     Meaning                     Parameters
    '-------------------------------------------------
    '
    ' &1        Alert (User Logged On)      UserHandle As String|
    ' &2        Alert (User Logged Off)     UserHandle As String|
    ' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)

    'Send alert out
    Call modServerAlertHandler.CastAlert("&2" & ISPNUSER_MemberHandle(Index) & Chr(11), Index, 0, ISPN_TopLevelCtlId)
    
End If

If ISPNUSER_MemberHandle(Index) = Chr(11) Then lblWaiting = lblWaiting - 1

'Required
Server(Index).Close

'Delete virtual winsock control (number Index) from array
If Not Index = 0 Then Unload Server(Index)

'Add an item to the running log
AddLogItem "Server index " & Index & " was unloaded"

'Sort out memory space
If ISPN_TopLevelCtlId = Index Then ISPN_TopLevelCtlId = ISPN_TopLevelCtlId - 1
lblTopMost = ISPN_TopLevelCtlId

ISPNUSER_MemberHandle(Index) = ""
ISPNUSER_LoggedOn(Index) = False

'Send out a list of contacts to all connected users
Call modServerISPN.CastMembersContacts(0, True)
End Sub

Private Sub Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim NewWSArrayIndex As Long


'Accept current connection request for authentication and start listening for the next connection request.
'=========================================================================================================

'This code looks through the variables and checks if there are any available
'indexes up for grabs that we can use to load the control for the next user to log on with
'before just loading a new control with a index of
'ISPN_TopLevelCtlId + 1. This saves on memory and will prevent eventual system failure.
'
'NOTE:
'This IS NOT finding a free control index and logging on with that, it is finding a free
'control index for the NEXT user to connect on. The user is attached to whatever control is listening..

For findfree = 0 To ISPN_TopLevelCtlId + 1
    
    If ISPNUSER_MemberHandle(findfree) = "" Then
        
        If findfree = Index Then GoTo nxt ' The free index found is not usable. It is this control!
            
        'Assign free control index
        NewWSArrayIndex = findfree: Exit For
        
    End If

nxt:

Next

If NewWSArrayIndex = ISPN_TopLevelCtlId + 1 Then

    'We do not have any gaps in the array, so update the top level control index
    ISPN_TopLevelCtlId = ISPN_TopLevelCtlId + 1

End If
    
lblNextConnection = NewWSArrayIndex
    
'=========================================================================================
'Ear-mark this index so if the logon string is not passed before this sub is called again,
'then the index that was found here does not get used twice.. preventing potential errors
'and \ or security breaches...
'This also lets other proceedures know that that control is awaiting logon data, and may
'be a wrong client version or the connection timed out..

ISPNUSER_MemberHandle(Index) = Chr(11)

'=========================================================================================






'Virtual Control creation =================================================

    'Add item to running log
    AddLogItem "Listening for next connection on index " & NewWSArrayIndex
    
    DoEvents
    
    'Create virtual control (if its not 0)
    If Not NewWSArrayIndex = 0 Then Load Server(NewWSArrayIndex)
    
    DoEvents
    
    'Start listening on the new control for the next connection request
    If Server(NewWSArrayIndex).State <> 0 Then Server(NewWSArrayIndex).Close
    If Server(Index).State <> 0 Then Server(Index).Close
    
    DoEvents
    
    Server(NewWSArrayIndex).Listen
    
    DoEvents
'=========================================================================

lblTopMost = ISPN_TopLevelCtlId

'Accepting this connection request ======================================

    'Update server status label: Awaiting LgnData
    'This user will be connected, but not authenticated.
    'Another feature will allow the server operator \ or automated process
    'to close connections that do not provide logon data within a timeout
    'period.. This would also saves on memory, and gives the server greater
    'capacity.
    
    lblWaiting = lblWaiting + 1    'Add item to running log
    AddLogItem "Accepted Connection Request " & requestID & " on index " & Index & " with IP: " & Server(Index).RemoteHostIP

    
    'Accept the connection on this control
    DoEvents
    Server(Index).Accept requestID
    DoEvents


'=========================================================================

End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'If the data is not a protocol command then close the connection
On Error GoTo killer

Dim p As String
Server(Index).GetData p

'Protocol ISPN1 (Client Side -> Server Side) (This data is handled in this sub!)
'
' | denotes a split character. This is Chr(11)
'
' $ = Compatible Server Language
' % = Service Carrier
' ? = Client Information Request
'
'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' $1        Normal Logon Request        UserHandle As String|Password As String|ServerPIN As Long|
' $2        Server Echo Logon Requset   UserHandle As String|Password As String|ServerPIN As Long|AdminPIN As Long|
' $3        Contact Information Request |
' $5        Browse Network Request      |
'
' %1        Instant Message Service     Recipient As String|MessageBody As String|
' %2        Profile Service (1\3)       UserToQuery As String| (Profile Information)
' %3        Profile Service (2\3)       ProfileBody as String|ProfilePic as Integer (1 or 0)|PhotoFileName as String
' %6        Profile Service (3\3)       UserToQuery As String| (General Information)
' %4        iMail Service (1\2)         |
' %5        iMail Service (2\2)         Recipient As String|MessageBody As String|Attachment As Integer (1 or 0)|AttachedFileName as String|
'
' ?1        Info Request (Services)     sHandle As Sting|ServiceIM As Integer (1=On)|ServiceProfiles As Integer (1=On)|ServiceiMail As Integer (1=On)|
' ?2        Info Request (WinVer)       sHandle As Sting|WindowsVersion As String|
' ?3        Info Request (PhysicalRAM)  sHandle As Sting|PhysicalRAM As String|
' ?4        Info Request (WindowsUserN) sHandle As Sting|WinUserName As String|


'Protocol ISPN1 (Server Side -> Client Side)
'
' | denotes a split character. This is Chr(11)
'
' $ = Compatible Server Language
' & = Alert
' % = Service Carrier
' ? = Client Information Request
'
'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' $1        Logon Type (1) Validated    (No Parameters)
' $3        Contact Information Data    NumberOfContacts As Integer|[1 = online; 2 = offline; 3 = admin online; 4 = admin offline]HandleShortName as String|
' $4        Contact Information Cancel  |
' $5        Browse Network Users Data   NumberOfHandles As Integer|[1 = online; 2 = offline; 3 = admin online; 4 = admin offline]HandleShortName as String|
' $5        Browse Network: NoAccess    Reason As String|
'
' &1        Alert (User Logged On)      UserHandle As String|
' &2        Alert (User Logged Off)     UserHandle As String|
' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)
' &4        Alert (New iMail)           |
' %1        Instant Message Service     Recipient As String|MessageBody As String|
' %2        Profile Service (1\3)       UserHandle As String|ProfileBody As String| (Profile Push)
' %3        Profile Service (2\3)       | (Update Successful)
' %6        Profile Service (3\3)       UserHandle As String|AccountMade As String|LogonCount As String| (General Information Push)
'
' ?1        Info Request (Services)     UserWhoIsQuerying As Sting|
' ?2        Info Request (WinVer)       UserWhoIsQuerying As Sting|
' ?3        Info Request (PhysicalRAM)  UserWhoIsQuerying As Sting|
' ?4        Info Request (WindowsUserN) UserWhoIsQuerying As Sting|


Select Case Left(p, 2)

    Case "$1" '$1 = Logon Request Type 1 (Normal User Logon)

        'Add item to running log
        AddLogItem "Validating logon request.."
        
        If ISPNUSER_LoggedOn(Index) = False Then
        
            'Validate the incoming string
            If ValidateRequest(Right(p, Len(p) - 2), Index) = True Then
                        
                
                Server(Index).SendData "$1" 'Send sucessful logon message
                
                ISPNUSER_LoggedOn(Index) = True 'Set varible to show user is logged on
                                
                lblConnected = lblConnected + 1 'Update server status label: Users Validated
                lblConnectedNow = lblConnectedNow + 1 'Update server status label: Connected Now
                lblWaiting = lblWaiting - 1 'Update server status label: Awaiting LgnData
            
                'Add item to running log
                AddLogItem "'" & ISPNUSER_MemberHandle(Index) & "' successfully connected (IP: " & Server(Index).RemoteHostIP & ")"
                                                            
                'Update user log
                modUserQuery.SaveUserInfo ISPNUSER_MemberHandle(Index), modGlobals.ISPN_INFO_LOGONCOUNT, (modUserQuery.GetUserInfo(ISPNUSER_MemberHandle(Index), modGlobals.ISPN_INFO_LOGONCOUNT) + 1)
                
                'Send alert out to logged on users
                
                'Protocol ISPN1 (Server Side -> Client Side)

                ' | signifies a split character. This is Chr(11)
                
                'Prefix     Meaning                     Parameters
                '-------------------------------------------------
                '
                ' &1        Alert (User Logged On)      UserHandle As String|
                ' &2        Alert (User Logged Off)     UserHandle As String|
                ' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)

                Call modServerAlertHandler.CastAlert("&1" & ISPNUSER_MemberHandle(Index) & Chr(11), Index, 0, ISPN_TopLevelCtlId)
                
                'Send contact list to the user
                Call modServerISPN.CastMembersContacts(Index, False)
    
                
            Else
                
                lblRefused = lblRefused + 1 'Update server status label: Refused
                lblWaiting = lblWaiting - 1 'Update server status label: Awaiting LgnData
                GoTo killer 'Disconnect the user
            
            End If
            
        Else
        'Unexpected data or the user is already logged on, so disconnect as it may not be a correct client
killer:
        
        'Add items to running log
        AddLogItem "Invalid Protocol Parameter \ Illegal Client Request (IP: " & Server(Index).RemoteHostIP & ")"
        AddLogItem "The client was disconnected."
        
        'Check if the ISPN_TopLevelCtlId variable needs to be clipped
        If ISPN_TopLevelCtlId = Index Then ISPN_TopLevelCtlId = ISPN_TopLevelCtlId - 1
        
        'Close the connection
        Server(Index).Close
        
        'If its a virtual control then delete it from memory
        If Not Index = 0 Then Unload Server(Index)
        
        'Send alert out to logged on users is the user is logged on..
        
        'Protocol ISPN1 (Server Side -> Client Side)

        ' | signifies a split character. This is Chr(11)
        
        'Prefix     Meaning                     Parameters
        '-------------------------------------------------
        '
        ' &1        Alert (User Logged On)      UserHandle As String|
        ' &2        Alert (User Logged Off)     UserHandle As String|
        ' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)

        If ISPNUSER_MemberHandle(Index) = "" Then GoTo noalrt
        If ISPNUSER_MemberHandle(Index) = Chr(11) Then GoTo noalrt
        
        lblConnected = lblConnected - 1

        Call modServerAlertHandler.CastAlert("&2" & ISPNUSER_MemberHandle(Index) & Chr(11), Index, 0, ISPN_TopLevelCtlId)
        
noalrt:
        
        'Reset variables
        ISPNUSER_MemberHandle(Index) = ""
        ISPNUSER_LoggedOn(Index) = False
        
        
        End If
            
        'Send out a contact list update to all logged in users
        Call modServerISPN.CastMembersContacts(0, True)
        
    
    Case "%1" 'Instant Message Cast Request (PLAINTEXT)
    
        'Cast the Instant Message
        Call modServerIMHandler.CastIM(Right(p, Len(p) - 2), Index)
        
    
    Case "%2" 'Instant Message Cast Request (RTF)
    
        MsgBox "Rich Text Formatting is not available with this server version. Please update your product.", 48
    
    Case "%6" ' %6        Profile Service (3\3)       UserToQuery As String|  (General Information)

        'Get information about the specified user and send back to client
        Call modUserQuery.PushGeneralInfo(Index, Left(Right(p, Len(p) - 2), Len(Right(p, Len(p) - 2)) - 1))
        
        
    Case "$3" ' $3        Contact Information Request |
        
        Call modServerISPN.CastMembersContacts(Index, False)
        
    Case Else 'Unexpected Data
        
        'Add items to running log
        AddLogItem "Unknown protocol data from '" + ISPNUSER_MemberHandle(Index) + "'. (IP: " & Server(Index).RemoteHostIP & ")"
        AddLogItem "The client was disconnected"
        
        'Check if the ISPN_TopLevelCtlId variable needs to be clipped
        If ISPN_TopLevelCtlId = Index Then ISPN_TopLevelCtlId = ISPN_TopLevelCtlId - 1
        
        'Close the connection
        Server(Index).Close
        
        'If its a virtual control then delete it from memory
        If Not Index = 0 Then Unload Server(Index)
        
        'Send alert out to logged on users is the user is logged on..
        
        'Protocol ISPN1 (Server Side -> Client Side)

        ' | signifies a split character. This is Chr(11)
        
        'Prefix     Meaning                     Parameters
        '-------------------------------------------------
        '
        ' &1        Alert (User Logged On)      UserHandle As String|
        ' &2        Alert (User Logged Off)     UserHandle As String|
        ' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)

        If ISPNUSER_MemberHandle(Index) = "" Then GoTo noalrt2
        If ISPNUSER_MemberHandle(Index) = Chr(11) Then GoTo noalrt2
        
        lblConnected = lblConnected - 1

        Call modServerAlertHandler.CastAlert("&2" & ISPNUSER_MemberHandle(Index) & Chr(11), Index, 0, ISPN_TopLevelCtlId)
noalrt2:
        
        'Reset variables
        ISPNUSER_MemberHandle(Index) = ""
        ISPNUSER_LoggedOn(Index) = False
        
End Select

End Sub

Private Sub tmrClientTimeOut_Timer()

End Sub

Private Sub Server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Add item to running log
AddLogItem "Winsock error: " & Description & " (" & Number & ")"
End Sub

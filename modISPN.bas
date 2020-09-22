Attribute VB_Name = "modISPN"
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

'ISPN Logon Info: Variables
Global ISPN_LoggedOn As Boolean

Global ISPN_LocalHandle As String
Global ISPN_AddressAttachedTo As String

Global ISPN_PortAttachedTo As Integer
Global ISPN_LocalPassword As String
Global ISPN_GivenPIN As Integer

'ISPN Services Constants
Global Const ISPN_SERVICE_IM = 0
Global Const ISPN_SERVICE_PROFILE = 1
Global Const ISPN_SERVICE_IMAIL = 2
          
'ISPN Alerts: Constants
Global Const ISPN_ALERT_UNSPECIFIED = 0

Global Const ISPN_ALERT_USERLOGON = 1
Global Const ISPN_ALERT_USERLOGOFF = 2
Global Const ISPN_ALERT_NEWIMAIL = 3
Global Const ISPN_ALERT_IM = 4

Global Const ISPN_ALERT_ALERTTEXT = 1
Global Const ISPN_ALERT_ALERTSHELL = 2


'Internal variables used for various things
Global HideTerminateMsg As Boolean

Public Function TrimHandle(sInput As String) As String
Dim tmpUserHandle As String

tmpUserHandle = sInput

'Trim User Handle down to short type
If Not InStr(vbNull, tmpUserHandle, "@") = 0 Then

    'Is full userhandle type
    'Extract short user handle
    For eXtract = 1 To Len(tmpUserHandle)
        
        If Right(Left(tmpUserHandle, eXtract), 1) = "@" Then tmpUserHandle = Left(tmpUserHandle, eXtract - 1): w = 1: Exit For
        
    Next eXtract
    
    'Check to see if the eXtract process trimmed the userhandle
    If Not w = 1 Then MsgBox "Invalid IM Recipient. You cannot recieve this instant message. Please check you are running the latest version of ISPN Client and ISPN Server.", 48, "Type Mismatch": Exit Function

End If
TrimHandle = tmpUserHandle
End Function

Public Function PerfomDefaultAction(sHandle As String)
Dim UsrOption As Integer
'Performs the default action as specified by the user in the options dialog

'Installed Services Compatible with this proceedure
'
'   ISPN_SERVICE_IM
'   ISPN_SERVICE_PROFILE
'   ISPN_SERVICE_IMAIL


'Get setting from registry
UsrOption = ISPN_SERVICE_IM 'Quick cheat while developing

'Execute action
Select Case UsrOption
    
    Case ISPN_SERVICE_IM
    
        'Shows a new\existing IM window for the handle
        Call modIMHandler.DisplayIM_VirtualWindow(TrimHandle(sHandle) & GetAttachedServer(True), , True)
    
    Case ISPN_SERVICE_PROFILE
        
        'Views 'sHandle's profile
        ShowToDo
    
    Case ISPN_SERVICE_IMAIL

        'Composes an internal mail message to 'sHandle'
        ShowToDo
    
End Select
End Function

Public Function GetAttachedServer(Optional AddAtSymbol As Boolean) As String
'Simply returns the server attached to string all in one
'used for full handle types
Select Case AddAtSymbol

    Case True

        GetAttachedServer = "@" & modISPN.ISPN_AddressAttachedTo & ":" & modISPN.ISPN_PortAttachedTo
        
    Case Else
    
        GetAttachedServer = modISPN.ISPN_AddressAttachedTo & ":" & modISPN.ISPN_PortAttachedTo

End Select
End Function

Public Function CheckNull(Expression As String, Optional FailOnSplitChar As Boolean) As Boolean
If Expression = "" Then CheckNull = True
If Expression = Space(Len(Expression)) Then CheckNull = True
If FailOnSplitChar = True Then
    If Not InStr(vbNull, Expression, Chr(11)) = 0 Then CheckNull = True
End If
End Function

Public Sub LogoffISPN()
ShowToDo
End Sub

Public Sub ShowToDo()
'This sub is for any unfinished parts
MsgBox "Please update your product. This feature is not available", 48
End Sub

Public Sub UpdateContacts(pData As String, Optional HideAdminUsers As Boolean)
'Protocol ISPN 1
'
' Prefix    Meaning                    Parameters
' -----------------------------------------------
'
' $3        Contact Information Data    NumberOfContacts As Integer|[1 = online; 2 = offline; 3 = admin online; 4 = admin offline]HandleShortName as String|

Dim tmpNumberOfContacts As Integer
Dim tmpHandleShortName As String
Dim tmpData As String

'Handle States
Dim tmpNormalOnline As Integer
Dim tmpNormalOffline As Integer
Dim tmpAdminOnline As Integer
Dim tmpAdminOffline As Integer
tmpNormalOnline = 1
tmpNormalOffline = 2
tmpAdminOnline = 3
tmpAdminOffline = 4

frmClient.lstContacts.ListItems.Clear 'TODO: Add a checker in here so wedont get flickering, so only add items that are not already present
frmClient.lstContacts.ListItems.Add , "Browser", "Browse " & frmClient.Client.RemoteHostIP, , 5

tmpData = pData 'Copy pData to tmpData
tmpNumberOfContacts = Left(tmpData, 1) 'Set the number of contacts we are adding

tmpData = Right(tmpData, Len(tmpData) - 2) 'Trim off the NumberOfContacts parameter, so we are left with just the contacts

For mainloop = 1 To tmpNumberOfContacts '(have to look for x amount of contacts)

    For findsplit = 1 To Len(tmpData) 'Start loop to find position of split char
    
        If Right(Left(tmpData, findsplit), 1) = Chr(11) Then 'gets a single character at point (findsplit) within tmpData
        
        tmpHandleShortName = Right(Left(tmpData, findsplit - 1), Len(Left(tmpData, findsplit - 1)) - 1)
        
        'add user to contacts with appropriate image for the online\offline status of the user
        
        Select Case Left(Left(tmpData, findsplit - 1), 1) 'Get the online\offline status of the contact
        
        Case tmpNormalOnline
        frmClient.lstContacts.ListItems.Add , "1" & tmpHandleShortName, tmpHandleShortName, , 1  'Add to lstContacts
        frmClient.lstContacts.ListItems(frmClient.lstContacts.ListItems.Count).Left = 20
        'tmpDebug = tmpDebug & tmpHandleShortName & ": type 0; online" & vbNewLine
        
        Case tmpNormalOffline
        frmClient.lstContacts.ListItems.Add , "0" & tmpHandleShortName, tmpHandleShortName, , 2 'Add to lstContacts
        frmClient.lstContacts.ListItems(frmClient.lstContacts.ListItems.Count).Left = 20
        'tmpDebug = tmpDebug & tmpHandleShortName & ": type 0; offline" & vbNewLine
                
        Case tmpAdminOnline
        If HideAdminUsers = False Then frmClient.lstContacts.ListItems.Add , "1" & tmpHandleShortName, tmpHandleShortName, , 3 'Add to lstContacts
        frmClient.lstContacts.ListItems(frmClient.lstContacts.ListItems.Count).Left = 20
        'tmpDebug = tmpDebug & tmpHandleShortName & ": type 1; online" & vbNewLine
        
        Case tmpAdminOffline
        If HideAdminUsers = False Then frmClient.lstContacts.ListItems.Add , "0" & tmpHandleShortName, tmpHandleShortName, , 4 'Add to lstContacts
        frmClient.lstContacts.ListItems(frmClient.lstContacts.ListItems.Count).Left = 20
        'tmpDebug = tmpDebug & tmpHandleShortName & ": type 1; offline" & vbNewLine
        
        End Select
        tmpData = Right(tmpData, Len(tmpData) - findsplit) 'chop off the found field so we can start that process again with the next contact field
        Exit For
        
        End If
        
    Next

Next

'Debug message just verifies the extracted fields
'MsgBox "Contacts: " & vbNewLine & tmpDebug, 64, "Public Sub UpdateContacts(pData As String, Optional HideAdminUsers As Boolean) As String - Debug Message"

With frmClient
    
    .lblWait.Visible = False
    .imgWait.Visible = False
    .lstContacts.Left = 0
    .lstContacts.Visible = True
    .lstContacts.Enabled = True

End With
End Sub

Public Sub SetControlStates()
'A sub to set properties of various different controls depending on what the logon state is.


If ISPN_LoggedOn = True Then
    
    With frmClient
           
        'Update system tray icon
        PROBas.SystemTrayAddIcon frmClient, frmClient, True
        
        'Client Menu
        .mnuLogoff.Enabled = True
        .mnuLogon.Enabled = False
        .mnuOrganiseContacts.Enabled = True
        
        'Services Menu
        .mnuBrowse.Enabled = True
        .mnuiM.Enabled = True
        .mnuProfile.Enabled = True
        .mnuAlerts.Enabled = True
        .mnuiMail.Enabled = True
        
        'User Menu
        .mnuIM2.Enabled = True
        .mnuiMail2.Enabled = True
        .mnuOrganiseContacts2.Enabled = True
        .mnuViewProfile.Enabled = True
        
        
        'Tray Menu
        .mnuCheckInbox.Enabled = True
        .mnuIM3.Enabled = True
        .mnuBrowseISPN.Enabled = True
        .mnuLogoff2.Enabled = True
        .mnuLogon2.Enabled = False
        
        'Main Window
        .Caption = "ISPN Client - [Logged On]"
        .lblAttachedTo = "Attached To: " & modISPN.ISPN_AddressAttachedTo
        .lblAttachedToHighlight = modISPN.ISPN_AddressAttachedTo
        .lblHandle = "Handle: " & modISPN.ISPN_LocalHandle
        .lblHandleHighlight = modISPN.ISPN_LocalHandle
        .lblAttachedTo.Visible = True
        .lblAttachedToHighlight.Visible = True
        .lblHandle.Visible = True
        .lblHandleHighlight.Visible = True
        .imgContactsTitle.Visible = True
        .fraLogon.Visible = False
    
    End With

    
    
Else
    
    With frmClient
        
        'Update system tray icon
        PROBas.SystemTrayAddIcon frmClient, frmIconContainer, True
        
        'Client Menu
        .mnuLogoff.Enabled = False
        .mnuLogon.Enabled = True
        .mnuOrganiseContacts.Enabled = False
        
        'Services Menu
        .mnuBrowse.Enabled = False
        .mnuiM.Enabled = False
        .mnuProfile.Enabled = False
        .mnuAlerts.Enabled = False
        .mnuiMail.Enabled = False
        
        'User Menu
        .mnuIM2.Enabled = False
        .mnuiMail2.Enabled = False
        .mnuOrganiseContacts2.Enabled = False
        .mnuViewProfile.Enabled = False
        
        'Tray Menu
        .mnuCheckInbox.Enabled = False
        .mnuIM3.Enabled = False
        .mnuBrowseISPN.Enabled = False
        .mnuLogoff2.Enabled = False
        .mnuLogon2.Enabled = True
        
        'Main Window
        .Caption = "ISPN Client - [Not Logged On]"
        .lblAttachedTo = "Attached To: (Not logged on)"
        .lblAttachedToHighlight = "(Not logged on)"
        .lblHandle = "Handle: (Not logged on)"
        .lblHandleHighlight = modISPN.ISPN_LocalHandle
        .lblAttachedTo.Visible = False
        .lblAttachedToHighlight.Visible = False
        .lblHandle.Visible = False
        .lblHandleHighlight.Visible = False
        .imgContactsTitle.Visible = False
        .lstContacts.Left = -4000
        .lstContacts.Visible = False
        .lstContacts.Enabled = False
        .lstContacts.ListItems.Clear
        .lblWait.Visible = False 'Not required; this would show the Please Wait label when no action is taking place
        .imgWait.Visible = False 'Not required; this would show the Please Wait label when no action is taking place
        .fraLogon.Visible = True

    End With
    
End If
End Sub

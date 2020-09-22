Attribute VB_Name = "modIMHandler"
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com
'
'Protocol ISPN1 (Server Side)
'
' | signifies a split character. This is Chr(11)
'
'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' $1        Normal Logon Request        UserHandle As String|Password As String|ServerPIN As Long|
' $2        Server Echo Logon Requset   UserHandle As String|Password As String|ServerPIN As Long|AdminPIN As Long|
' %1        Instant Message Service     Recipient As String|MessageBody As String|
' %2        Profile Service (1\2)       UserToQuery As String|
' %4        Profile Service (2\2)       ProfileBody as String|ProfilePic as Integer (1 or 0)|PhotoFileName as String
' %4        iMail Service (1\2)         |
' %5        iMail Service (2\2)         Recipient As String|MessageBody As String|Attachment As Integer (1 or 0)|AttachedFileName as String|
'
'
'Protocol ISPN1 (Client Side)
'
'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' %1        Instant Message Service     From As String|MessageBody As String|

Global Const IM_ACTION_ADDTEXT = 0 '(Default)
Global Const IM_ACTION_ENABLE = 1
Global Const IM_ACTION_DISABLE = 2

Global Const IM_RECIPIENT = 1
Global Const IM_BODY = 2


Public Sub SendIMToHandle(IMRecipient As String, IMBody As String)
' %1        Instant Message Service     Recipient As String|MessageBody As String|

'Check for null string
If IMRecipient = "" Then MsgBox "Unspecified user handle. Please specify a handle to connect to.", 16: Exit Sub

'Sends a datagram to the server that then gets casted to the recipient
frmClient.Client.SendData "%1" & IMRecipient & Chr(11) & IMBody & Chr(11)

End Sub

Public Function ExtractIMData(pData As String, DataField As Integer) As String
Dim tmpRecipient As String
Dim tmpBody As String

Dim tmpData As String

tmpData = pData 'Copy pData to tmpData

For mainloop = 1 To 3 '(have to look for 3 variables)

    For findsplit = 1 To Len(tmpData) 'Start loop to find position of split char
    
        If Right(Left(tmpData, findsplit), 1) = Chr(11) Then 'gets a single character at point (findsplit) within tmpData
        
        'found it
        Select Case mainloop
        
        Case IM_RECIPIENT 'Recipient
        tmpRecipient = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        Case IM_BODY 'Message Body
        tmpBody = Left(tmpData, findsplit - 1) 'allocate the field to a variable
                
        End Select
        tmpData = Right(tmpData, Len(tmpData) - findsplit) 'chop off the found field
        Exit For
        
        End If
        
    Next

Next

'Debug message just verifies the extracted fields
'MsgBox "Recipient: " & tmpRecipient & vbNewLine & "Message: " & tmpBody, 64, "Public Function ExtractIMData(pData As String, DataField As Integer) As String - Output"

If DataField = IM_RECIPIENT Then ExtractIMData = tmpRecipient
If DataField = IM_BODY Then ExtractIMData = tmpBody

End Function

Public Function DisplayIM_VirtualWindow(sUserHandle As String, Optional sTextToAdd As String, Optional NormaliseWindowState As Boolean, Optional cAction As Integer)
'Links to an IM window that is already open, or opens a new one
'==============================================================

Dim tmpFoundHandle As Boolean
Dim tmpUserHandle As String
Dim tmpAtHost As String

'Assign variables
tmpUserHandle = sUserHandle
tmpUserHandle = modISPN.TrimHandle(tmpUserHandle)
tmpAtHost = GetAttachedServer

'Search through windows to see if we have a conversation open already
For Each Form In VB.Forms
    
    If LCase(Right(Form.Caption, Len("conversation"))) = LCase("conversation") Then 'Found a IM window
    
        'Check to make sure that the window we are looking at is the target
        If Len(tmpUserHandle) + Len(" - Conversation") >= Len(Form.Caption) Then 'Possible match
            
            If LCase(Left(Form.Caption, Len(tmpUserHandle))) = LCase(tmpUserHandle) Then 'Is match
            
                tmpFoundHandle = True 'Found a window already open with that handle
                
                If cAction = IM_ACTION_ENABLE Then Call Form.SetState(False)
                If cAction = IM_ACTION_DISABLE Then Call Form.SetState(True)
                If cAction = IM_ACTION_ADDTEXT Then Call Form.DisplayIM(tmpUserHandle, tmpAtHost, sTextToAdd)
                
                'Restore the form's window state to '0-Normal'
                If NormaliseWindowState = True Then Form.WindowState = 0
                
            End If
            
        End If
        
        
    End If
    
Next
  
'----------------------------------------------------------------------------------------

If tmpFoundHandle = False Then 'Didnt find a window already open so open a new IM window.
    
    If cAction = IM_ACTION_ENABLE Then Exit Function
    If cAction = IM_ACTION_DISABLE Then Exit Function

    'Create a virtual IM window in memory
    Dim ISPN_IMSTACK_WindowHandle As New frmIM
    
    'Open the virtual IM window with the specified text and userhandle
    Call ISPN_IMSTACK_WindowHandle.DisplayIM(tmpUserHandle, tmpAtHost, sTextToAdd)
    
    'Show IM alert?
    If Not sTextToAdd = "" Then
                
        Dim tmpAlertMsg As String
        'See if IM Alerts are turned on, if so show alert window
        'otherwise give direct focus to the IM window
        
        retval = GetSetting("Nexis Software Technologies", "ISPN_Server", "Alerts_IM Alert", 1)
        
        If retval = 1 Then
            
            'Trim message to fit in alert window
            If Len(tmpUserHandle & ":" & sTextToAdd) >= 27 Then
                
                'trim
                tmpAlertMsg = tmpUserHandle & ":" & sTextToAdd
                tmpAlertMsg = Left(tmpAlertMsg, 25) & ".."
                
            Else
                
                tmpAlertMsg = tmpUserHandle & ":" & sTextToAdd
            
            End If
            
            'Show IM alert window
            frmAlerter.ShowAlert tmpAlertMsg, modISPN.ISPN_ALERT_IM, tmpUserHandle
            
            
        Else
        
            'Just give normal focus to window
            ISPN_IMSTACK_WindowHandle.WindowState = 0
            
        End If
    Else
    
        'Just give normal focus to window
        ISPN_IMSTACK_WindowHandle.WindowState = 0
    
    End If
    
End If

End Function

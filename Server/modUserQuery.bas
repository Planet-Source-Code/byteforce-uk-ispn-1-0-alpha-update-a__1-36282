Attribute VB_Name = "modUserQuery"

Global Const IM_RECIPIENT = 1
Global Const IM_BODY = 2

Public Function CheckUserHandle(UserHandle As String) As Boolean

'standardise the user handle to all lowercase letters
UserHandle = LCase(UserHandle)

'check if user exists
If UserHandle = "admin" Then CheckUserHandle = True: Exit Function
If UserHandle = "guest" Then CheckUserHandle = True: Exit Function
If UserHandle = "david" Then CheckUserHandle = True: Exit Function
If UserHandle = "matt" Then CheckUserHandle = True: Exit Function
If UserHandle = "jason" Then CheckUserHandle = True: Exit Function
If UserHandle = "bob" Then CheckUserHandle = True: Exit Function
If UserHandle = "joe" Then CheckUserHandle = True: Exit Function

If UserHandle = "paul" Then CheckUserHandle = True: Exit Function
If UserHandle = "rachel" Then CheckUserHandle = True: Exit Function
If UserHandle = "julia" Then CheckUserHandle = True: Exit Function
If UserHandle = "adam" Then CheckUserHandle = True: Exit Function


CheckUserHandle = False
End Function

Public Sub PushGeneralInfo(CtlIndex As Integer, sQueryHandle As String)
'**Bounces back general information about a specified user (sQueryHandle)**
 
'Protocol ISPN1 (Client Side -> Server Side)

' | signifies a split character. This is Chr(11)

'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' %6        Profile Service (3\3)       UserToQuery As String|  (General Information)

frmServer.AddLogItem "General Information Query (" & sQueryHandle & " > " & ISPNUSER_MemberHandle(CtlIndex)


'Protocol ISPN1 (Server Side -> Client Side)

' | denotes a split character. This is Chr(11)

'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' %6        Profile Service (3\3)       UserHandle As String|AccountMade As String|LogonCount As String| (General Information Push)


If CheckUserHandle(sQueryHandle) = False Then
    
    'Add item to running log
     frmServer.AddLogItem "Query failed: (Unknown user " & sQueryHandle
    
    'Fire blanks back @ the destination socket
    frmServer.Server(CtlIndex).SendData "%6" & sQueryHandle & Chr(11) & Chr(11) & Chr(11)
    
    Exit Sub

End If

'Gather information and send
frmServer.Server(CtlIndex).SendData "%6" & sQueryHandle & Chr(11) & GetUserInfo(sQueryHandle, ISPN_INFO_ACCOUNTCREATED) & Chr(11) & GetUserInfo(sQueryHandle, ISPN_INFO_LOGONCOUNT) & Chr(11)

End Sub

Public Function GetUserInfo(UserHandle As String, InfoField As Integer) As String
'General Information
    'ISPN_INFO_ACCOUNTCREATED = 0
    'ISPN_INFO_LOGONCOUNT = 1
    'ISPN_INFO_IMCOUNT = 2

'Service Information
    'ISPN_INFO_IMENABLED = 3
    'ISPN_INFO_IMAILENABLED = 4
    'ISPN_INFO_PROFILEENABLED = 5

Select Case InfoField

    Case ISPN_INFO_ACCOUNTCREATED
    
        'Read information from registry
        GetUserInfo = GetSetting("Nexis Software Technologies", "ISPN_Server", "Log_AccountCreated-" & UserHandle, "This information is currently unavailable")
        
    Case ISPN_INFO_LOGONCOUNT
        
        'Read information from registry
        GetUserInfo = GetSetting("Nexis Software Technologies", "ISPN_Server", "Log_LogonCount-" & UserHandle, 0)
    
    Case ISPN_INFO_IMCOUNT
    
        'Read information from registry
        GetUserInfo = GetSetting("Nexis Software Technologies", "ISPN_Server", "Log_IMCount-" & UserHandle, 0)

    Case ISPN_INFO_IMENABLED
     
        'Read information from target
        
        
    Case ISPN_INFO_IMAILENABLED
        
        'Read information from target
    
    Case ISPN_INFO_PROFILEENABLED
        
        'Read information from target
        
    
End Select

End Function

Public Sub SaveUserInfo(UserHandle As String, InfoField As Integer, sDataToWrite As String)
'General Information
    'ISPN_INFO_ACCOUNTCREATED = 0
    'ISPN_INFO_LOGONCOUNT = 1
    'ISPN_INFO_IMCOUNT = 2

'Service Information
    'ISPN_INFO_IMENABLED = 3
    'ISPN_INFO_IMAILENABLED = 4
    'ISPN_INFO_PROFILEENABLED = 5

Select Case InfoField

    Case ISPN_INFO_ACCOUNTCREATED
    
        'Save information to registry
        SaveSetting "Nexis Software Technologies", "ISPN_Server", "Log_AccountCreated-" & UserHandle, sDataToWrite
        
    Case ISPN_INFO_LOGONCOUNT
        
        'Save information to registry
        SaveSetting "Nexis Software Technologies", "ISPN_Server", "Log_LogonCount-" & UserHandle, sDataToWrite
    
    Case ISPN_INFO_IMCOUNT
    
        'Save information to registry
        SaveSetting "Nexis Software Technologies", "ISPN_Server", "Log_IMCount-" & UserHandle, sDataToWrite

    Case ISPN_INFO_IMENABLED
     
        MsgBox "This infomation is not stored on the server side. Please enable \ disable client services on the individual client(s).", 48
    
    Case ISPN_INFO_IMAILENABLED
        
        MsgBox "This infomation is not stored on the server side. Please enable \ disable client services on the individual client(s).", 48
    
    Case ISPN_INFO_PROFILEENABLED
        
        MsgBox "This infomation is not stored on the server side. Please enable \ disable client services on the individual client(s).", 48
        
End Select

End Sub

Public Function CheckUserPwd(UserHandle As String, GivenPwd As String) As Boolean
'standardise the user handle to all lowercase letters
UserHandle = LCase(UserHandle)

If UserHandle = "admin" Then
If GivenPwd = "nst" Then CheckUserPwd = True
Exit Function
End If

If UserHandle = "guest" Then
If GivenPwd = "nst" Then CheckUserPwd = True
Exit Function
End If

If UserHandle = "david" Then
If GivenPwd = "nst" Then CheckUserPwd = True
Exit Function
End If

If UserHandle = "matt" Then
If GivenPwd = "oracle20" Then CheckUserPwd = True
Exit Function
End If


If UserHandle = "joe" Then
If GivenPwd = "1" Then CheckUserPwd = True
Exit Function
End If

If UserHandle = "bob" Then
If GivenPwd = "bob" Then CheckUserPwd = True
Exit Function
End If

If UserHandle = "jason" Then
If GivenPwd = "kebab" Then CheckUserPwd = True
Exit Function
End If

CheckUserPwd = False
End Function
Public Function IsAccountEnabled(UserHandle As String) As Boolean
'standardise the user handle to all lowercase letters
UserHandle = LCase(UserHandle)

If UserHandle = "admin" Then
IsAccountEnabled = True
Exit Function
End If

If UserHandle = "guest" Then
IsAccountEnabled = False
Exit Function
End If

If UserHandle = "david" Then
IsAccountEnabled = True
Exit Function
End If

If UserHandle = "matt" Then
IsAccountEnabled = True
Exit Function
End If

If UserHandle = "joe" Then
IsAccountEnabled = True
Exit Function
End If

If UserHandle = "bob" Then
IsAccountEnabled = True
Exit Function
End If

If UserHandle = "jason" Then
IsAccountEnabled = True
Exit Function
End If

IsAccountEnabled = False
End Function

Public Function ValidateRequest(pData As String, wsIdx As Integer) As Boolean
'If there is an error then the input string is not the right format so fail it.
On Error GoTo nopass

'Processes a logon request
'Split character is Chr(11)
  
'There should be 3 fields within pData;
'UserHandle,Password,PIN

'We have to split the pData string up in order to extract these three fields
'This is possible by looking for the split character within the string.
'which is defined in Chr(11).

Dim tmpUsrHandle As String
Dim tmpUsrPwd As String
Dim tmpPIN As String

Dim tmpData As String
Dim strFailReason As String

tmpData = pData 'Copy pData to tmpData

For mainloop = 1 To 3 '(have to look for 3 variables)

    For findsplit = 1 To Len(tmpData) 'Start loop to find position of split char
    
        If Right(Left(tmpData, findsplit), 1) = Chr(11) Then 'gets a single character at point (findsplit) within tmpData
        
        'found it
        Select Case mainloop
        
        Case 1 'User Handle
        tmpUsrHandle = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        Case 2 'User Password
        tmpUsrPwd = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        Case 3 'Server PIN
        tmpPIN = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        End Select
        tmpData = Right(tmpData, Len(tmpData) - findsplit) 'chop off the found field
        Exit For
        
        End If
        
    Next

Next

'Debug message just verifies the extracted fields
'MsgBox "Username: " & tmpUsrHandle & vbNewLine & "Password: " & tmpUsrPwd & vbNewLine & "PIN: " & tmpPIN, 64, "Public Function ValidateLogin(pData as String) As Boolean - Output"

'Check the given data by calling other functions
If CheckUserHandle(tmpUsrHandle) = False Then strFailReason = "User does not exist": GoTo nopass 'Check if the user exists
If IsUserLoggedOn(tmpUsrHandle) = True Then strFailReason = "User already logged on": GoTo nopass 'Check to see if the user is already logged on
If CheckUserPwd(tmpUsrHandle, tmpUsrPwd) = False Then strFailReason = "Incorrect password": GoTo nopass 'Check Password
If Not frmServer.txtPIN = tmpPIN Then strFailReason = "Incorrect PIN": GoTo nopass 'Check server PIN

'Set the members handle
ISPNUSER_MemberHandle(wsIdx) = tmpUsrHandle

'Set the function's return value to TRUE
ValidateRequest = True

Exit Function
nopass:
If Err <> 0 Then
frmServer.AddLogItem "Validation failed: " + Err.Description + " (" & Err.Number & ")"
Else
frmServer.AddLogItem "Validation failed: " + strFailReason
End If
ValidateRequest = False
End Function

Public Function CreateUser(UniqueName As String, Password As String)
ShowToDo
End Function

Public Sub ShowToDo()
'This sub is for any unfinished parts
MsgBox "Please update your product. This feature is not available", 48
End Sub

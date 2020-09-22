Attribute VB_Name = "modServerISPN"
Public Sub CastMembersContacts(CtlIndex As Integer, Optional SendToAll As Boolean)
'Sends out a contact list to one or all users logged on (SendToAll As Boolean)

Dim tmpCtlRange As Integer
Dim tmpCtlIndex As Integer

'Deal with parameters and set the cast range accordingly
If SendToAll = False Then
    
    'We arent sending a contact list to everyone, just an individual
    'so set the range to then same as CtlIndex (the index the user
    'is attached to)
    tmpCtlRange = CtlIndex
    tmpCtlIndex = CtlIndex

Else
    
    'Set the range: from 0 to ISPN_TopLevelCtlID (All users)
    tmpCtlRange = ISPN_TopLevelCtlId
    tmpCtlIndex = 0
    
End If


For CurrentUser = tmpCtlIndex To tmpCtlRange 'Loop for however many users are logged on

    'Check to make sure that the CurrentUser index is attached to a user..
    'if not then skip the cast and look at the next index
    If ISPNUSER_MemberHandle(CurrentUser) = "" Or ISPNUSER_MemberHandle(CurrentUser) = Chr(11) Then GoTo nocastc

    'Add item to running log
    frmServer.AddLogItem "Casting out contact information to '" & ISPNUSER_MemberHandle(CurrentUser) & "'"
    
    'Protocol ISPN 1 (Data from Server Side to Client Side)
    '
    ' Prefix    Meaning                    Parameters
    ' -----------------------------------------------
    '
    ' $3        Contact Information Data    NumberOfContacts As Integer|[1 = online; 2 = offline; 3 = admin online; 4 = admin offline]HandleShortName as String|
    ' $4        Contact Information Cancel  |
    
    Dim CData As String
    Dim tmpNumberOfContacts As Integer
    
    'Reset Vars
    tmpNumberOfContacts = 0
    CData = ""
    
    'Loop for amount of users potentially logged in, and add any found users to the CData string
    'except for the CurrentUser. (We dont want the client to see his\her self in the contacts list)
    For FindUsers = 0 To ISPN_TopLevelCtlId
        
        If Not ISPNUSER_MemberHandle(FindUsers) = "" Then
            
            'Check to make sure that this user is not the user we have just found
            If LCase(ISPNUSER_MemberHandle(FindUsers)) = LCase(ISPNUSER_MemberHandle(CurrentUser)) Then GoTo nxtchk
            
            'Add user to string that will be sent out to the client (contact list)
            CData = CData & "1" & ISPNUSER_MemberHandle(FindUsers) & Chr(11)
            
            tmpNumberOfContacts = tmpNumberOfContacts + 1 'Update counter
        
        End If
    
nxtchk:
    
    Next FindUsers
    
    'There are no users signed in except for the user requesting this update so exit the
    'sub before the a list is sent out, preventing error
    If tmpNumberOfContacts = 0 Then
        
        'Add item to running log
        frmServer.AddLogItem "No users logged on except '" & ISPNUSER_MemberHandle(CurrentUser) & "'. Cast cancelled."
        
        

        'Send cancelled notification to client (if there is a client attached to that ctlindex)
        If ISPNUSER_MemberHandle(CurrentUser) = "" Then GoTo nocastc
        If ISPNUSER_MemberHandle(CurrentUser) = Chr(11) Then GoTo nocastc
        
        DoEvents
        
        frmServer.Server(CurrentUser).SendData "$4" & Chr(11)
        
        DoEvents
        
        Exit Sub
    
    End If
    
    'Send list to client
    DoEvents
    
    frmServer.Server(CurrentUser).SendData "$3" & tmpNumberOfContacts & Chr(11) & CData
    
    DoEvents
    
    'Add item to running log
    frmServer.AddLogItem "Sucessfully sent contact information to '" & ISPNUSER_MemberHandle(CurrentUser) & "'"

nocastc: 'No user logged on at Index CurrentUser

Next

End Sub

Public Function IsUserLoggedOn(SearchUser As String) As Boolean
'Returns True or false is the user specified in SearchUser is already logged on..
'
'..Does not use the ISPNUSER_LoggedOn() variable.. for reasons i cant be bothered to explain.

For uCompare = 0 To ISPN_TopLevelCtlId
If LCase(ISPNUSER_MemberHandle(uCompare)) = LCase(SearchUser) Then IsUserLoggedOn = True: Exit For
Next uCompare

End Function

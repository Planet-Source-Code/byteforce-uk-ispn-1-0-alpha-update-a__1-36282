Attribute VB_Name = "modServerIMHandler"
Public Sub CastIM(pData As String, CtlIndex As Integer, Optional GlobalMSG As Boolean)
            
Dim tmpIMRecipient As String
Dim tmpIMBody As String
Dim tmppData As String
Dim tmpUserIdx As Integer
Dim tmpAtHost As String
tmppData = pData


            '-Extract fields
            tmpIMRecipient = ExtractIMData(tmppData, IM_RECIPIENT)
            
            tmpIMBody = ExtractIMData(tmppData, IM_BODY)
            
            'Add item to running log
            frmServer.AddLogItem "Casting IM (" & ISPNUSER_MemberHandle(CtlIndex) & " > " & tmpIMRecipient & ").."
            
            'Update user log
            modUserQuery.SaveUserInfo ISPNUSER_MemberHandle(Index), modGlobals.ISPN_INFO_IMCOUNT, (modUserQuery.GetUserInfo(ISPNUSER_MemberHandle(Index), modGlobals.ISPN_INFO_IMCOUNT) + 1)
            
            '-Find the recipients index
            tmpUserIdx = -1
            
            For finduser = 0 To ISPN_TopLevelCtlId
                
                If ISPNUSER_MemberHandle(finduser) = "" Then GoTo nxt2:
                If ISPNUSER_MemberHandle(finduser) = Chr(11) Then GoTo nxt2:
                
                
                'Look for a match in the users logged on.
                'Searches for the following formats..
                '
                '   -Exact match (Not case sensitive)
                '   -user@127.0.0.1
                '   -user@localhostname
                '   -user@customdns.co.uk
                '   -user@127.0.0.1:port
                '   -user@localhostname:port
                '   -user@customdns.co.uk:port
                
                
                If LCase(ISPNUSER_MemberHandle(finduser)) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If
                
                If LCase(ISPNUSER_MemberHandle(finduser) & "@" & frmServer.Server(0).LocalIP) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If
                
                If LCase(ISPNUSER_MemberHandle(finduser) & "@" & frmServer.Server(0).LocalHostName) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If
                
                If LCase(ISPNUSER_MemberHandle(finduser) & "@" & ISPN_CustomDNS) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If

                If LCase(ISPNUSER_MemberHandle(finduser) & "@" & frmServer.Server(0).LocalIP & ":" & frmServer.Server(0).LocalPort) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If
                
                If LCase(ISPNUSER_MemberHandle(finduser) & "@" & frmServer.Server(0).LocalHostName & ":" & frmServer.Server(0).LocalPort) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If
                
                If LCase(ISPNUSER_MemberHandle(finduser) & "@" & ISPN_CustomDNS & ":" & frmServer.Server(0).LocalPort) = LCase(tmpIMRecipient) Then
                
                    tmpUserIdx = finduser
                    
                    Exit For
                
                End If

nxt2:
            
            Next
            
            'Set AtHost variable
            If ISPN_CustomDNS = "" Then tmpAtHost = frmServer.Server(0).LocalIP Else tmpAtHost = ISPN_CustomDNS
            tmpAtHost = tmpAtHost & ":" & frmServer.Server(0).LocalPort 'Add on the port identifier
            
            'See if user was found or not, if not then abort IM cast
            If tmpUserIdx = -1 Then frmServer.AddLogItem "Cast Failed: Unknown user @" & tmpAtHost: Exit Sub
                        
            '-Cast IM message to recipient
            
            'See protocol.txt for more information
            frmServer.Server(tmpUserIdx).SendData "%1" & ISPNUSER_MemberHandle(CtlIndex) & "@" & tmpAtHost & Chr(11) & tmpIMBody & Chr(11)
            
            'Add item to running log
            frmServer.AddLogItem "Successfully sent IM"

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


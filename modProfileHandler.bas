Attribute VB_Name = "modProfileHandler"
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com
    
Global Const ISPN_INFO_TARGET = 0
Global Const ISPN_INFO_ACCOUNTCREATED = 1
Global Const ISPN_INFO_LOGONCOUNT = 2


Public Function RequestGeneralInfo(sHandle As String)
'Protocol ISPN1 (Client Side -> Server Side)

' | signifies a split character. This is Chr(11)

'Prefix     Meaning                     Parameters
'-------------------------------------------------
'
' %6        Profile Service (3\3)       UserToQuery As String|  (General Information)
'

'Request the information from the server.
frmClient.Client.SendData "%6" & TrimHandle(sHandle) & Chr(11)

'The information is updated into the virtual window by modProfileHandler.DisplayGeneralInfo(sHandle As String, ByVal sAccountMade As String, ByVal sLogonCount As String)
End Function

Public Function DisplayGeneralInfo(sHandle As String, Optional ByVal sAccountMade As String, Optional ByVal sLogonCount As String, Optional bNewRequest As Boolean)
Dim tmpFoundHandle As Boolean
Dim tmpUserHandle As String
Dim tmpAtHost As String

'Assign variables
tmpUserHandle = sHandle
tmpUserHandle = modISPN.TrimHandle(tmpUserHandle)


'Search through windows to see if we have a conversation open already
For Each Form In VB.Forms
    
    If LCase(Right(Form.Caption, Len("information"))) = LCase("information") Then 'Found a IM window
    
        'Check to make sure that the window we are looking at is the target
        If Len(tmpUserHandle) + Len(" - information") >= Len(Form.Caption) Then 'Possible match
            
            If LCase(Left(Form.Caption, Len(tmpUserHandle))) = LCase(tmpUserHandle) Then 'Is match
            
                tmpFoundHandle = True 'Found a window already open with that handle
                
                If bNewRequest = False Then
                    
                    'Add information to window that is already open
                    Call Form.UpdateGeneralInfo(sAccountMade, sLogonCount)
                    
                Else
                                    
                    'Refresh screen
                    Call Form.GetUserInfo(sHandle)
                                    
                End If
            
            End If
            
        End If
        
        
    End If
    
Next
  
'----------------------------------------------------------------------------------------
newreq:

If tmpFoundHandle = False Then 'Didnt find a window already open so open a new IM window.
    
    'Create a virtual user info window in memory
    Dim ISPN_USERINFOSTACK_WindowHandle As New frmUserInfo
    
    'Open the virtual user info window
    Call ISPN_USERINFOSTACK_WindowHandle.GetUserInfo(sHandle)
    
End If

End Function

Public Function ExtractGeneralInfoData(pData As String, DataField As Integer) As String
'General Information
    'ISPN_INFO_TARGET = 0
    'ISPN_INFO_ACCOUNTCREATED = 1
    'ISPN_INFO_LOGONCOUNT = 2

Dim tmpTarget As String
Dim tmpAccMade As String
Dim tmpLgnCount As String

Dim tmpData As String

tmpData = pData 'Copy pData to tmpData

For mainloop = 1 To 3 '(have to look for 3 variables)

    For findsplit = 1 To Len(tmpData) 'Start loop to find position of split char
    
        If Right(Left(tmpData, findsplit), 1) = Chr(11) Then 'gets a single character at point (findsplit) within tmpData
        
        'found it
        Select Case mainloop
        
        Case ISPN_INFO_TARGET + 1 'Recipient
        tmpTarget = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        Case ISPN_INFO_ACCOUNTCREATED + 1 'Message Body
        tmpAccMade = Left(tmpData, findsplit - 1) 'allocate the field to a variable
                
        Case ISPN_INFO_LOGONCOUNT + 1 'Message Body
        tmpLgnCount = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        End Select
        tmpData = Right(tmpData, Len(tmpData) - findsplit) 'chop off the found field
        Exit For
        
        End If
        
    Next

Next


If DataField = ISPN_INFO_TARGET Then ExtractGeneralInfoData = tmpTarget
If DataField = ISPN_INFO_ACCOUNTCREATED Then ExtractGeneralInfoData = tmpAccMade
If DataField = ISPN_INFO_LOGONCOUNT Then ExtractGeneralInfoData = tmpLgnCount

End Function


Attribute VB_Name = "modAlertHandler"
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Public Function ExtractAlertData(pData As String, AlertType As Integer, Optional DataField As Integer) As String
'Extracts required fields from pData

' The new iMail alert does not require parameters. Maybe ill add some in like 'NewMailCount As Integer'


If AlertType = ISPN_ALERT_USERLOGON Or AlertType = ISPN_ALERT_USERLOGOFF Then
    
    'Just trim the last split char off and return the parameter
    ExtractAlertData = Left(pData, Len(pData) - 1)
    Exit Function

End If

Dim tmpAlertText As String

Dim tmpShellString As String


tmpData = pData 'Copy pData to tmpData

For mainloop = 1 To 2 '(have to look for 3 variables)

    For findsplit = 1 To Len(tmpData) 'Start loop to find position of split char
    
        If Right(Left(tmpData, findsplit), 1) = Chr(11) Then 'gets a single character at point (findsplit) within tmpData
        
        'found it
        Select Case mainloop
        
        Case ISPN_ALERT_ALERTTEXT 'Alert Text
        tmpRecipient = Left(tmpData, findsplit - 1) 'allocate the field to a variable
        
        Case ISPN_ALERT_ALERTSHELL 'ShellString
        tmpBody = Left(tmpData, findsplit - 1) 'allocate the field to a variable
                
        End Select
        tmpData = Right(tmpData, Len(tmpData) - findsplit) 'chop off the found field
        Exit For
        
        End If
        
    Next

Next

'Return the requested field
If DataField = ISPN_ALERT_ALERTTEXT Then ExtractAlertData = tmpAlertText
If DataField = ISPN_ALERT_ALERTSHELL Then ExtractAlertData = tmpShellString
End Function


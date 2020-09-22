Attribute VB_Name = "modServerAlertHandler"

Public Sub CastAlert(sParameters As String, sException As Integer, CtlIndex As Integer, Optional CtlRange As Integer)
Dim tmpCtlRange As Integer

tmpCtlRange = CtlRange
If CtlRange = 0 Then tmpCtlRange = CtlIndex


For spanit = CtlIndex To tmpCtlRange
    
    If spanit = sException Then GoTo noalrter
        
        'Check to see if a user is attached to that socket before attempting to cast the alert
        If ISPNUSER_MemberHandle(spanit) = "" Then GoTo noalrter
        If ISPNUSER_MemberHandle(spanit) = Chr(11) Then GoTo noalrter
        
        'Add item to running log
        frmServer.AddLogItem "Casting alert to '" & ISPNUSER_MemberHandle(spanit) & "'"
        
        'Protocol ISPN1 (Server Side -> Client Side)
        
        ' | signifies a split character. This is Chr(11)
        
        'Prefix     Meaning                     Parameters
        '-------------------------------------------------
        '
        ' &1        Alert (User Logged On)      UserHandle As String|
        ' &2        Alert (User Logged Off)     UserHandle As String|
        ' &3        Alert (Standard)            AlertText As String|ShellString As String|  (when ShellString begins with !, this will send a TCP\IP packet containing the rest of the string)
        
        'Send alert to client
        DoEvents
        
        frmServer.Server(spanit).SendData sParameters
        
        DoEvents
        
        'Add item to running log
        frmServer.AddLogItem "Sucessfully sent alert to '" & ISPNUSER_MemberHandle(spanit) & "'"
    
noalrter:

Next

End Sub

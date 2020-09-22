Attribute VB_Name = "modGlobals"
Global ISPNUSER_MemberHandle(200) As String
Global ISPNUSER_LoggedOn(200) As Boolean

Global ISPN_CustomDNS As String 'For use with full user handles (someone@[ISPN_CustomDNS:serverport])
Global ISPN_TopLevelCtlId As Integer 'Top level control index counter


'Constants

'For use with modUserQuery.GetUserInfo =========

'General Information
Global Const ISPN_INFO_ACCOUNTCREATED = 0
Global Const ISPN_INFO_LOGONCOUNT = 1
Global Const ISPN_INFO_IMCOUNT = 2

'Service Information
Global Const ISPN_INFO_IMENABLED = 3
Global Const ISPN_INFO_IMAILENABLED = 4
Global Const ISPN_INFO_PROFILEENABLED = 5

'===============================================

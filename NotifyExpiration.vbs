' Author: Dario Alves Junior - dairoajr@gmail.com (github)
' Date: 06/07/2015
' Purpose: To send e-mail and sms to registered users in Active Directory,
'          3 days before the password expires, until it is changed.
Option Explicit

Dim adoCommand, adoConnection, strBase, strFilter, strAttributes
Dim objRootDSE, strDNSDomain, strQuery, adoRecordset
Dim dtmDate1, dtmDate2, intDays, strName, strEmail
Dim lngSeconds1, str64Bit1, lngSeconds2, str64Bit2
Dim objShell, lngBiasKey, lngBias, k
Dim objDomain, objMaxPwdAge, lngHighAge, lngLowAge, sngMaxPwdAge
Dim objDate, dtmPwdLastSet, dtmExpires
Dim arrEmails, strItem, strPrefix
Dim agora, strSms, strEmail_a, arrEmails_a, strCNName, distinguishedName, strEmailNOC

' Number of days the password expiration check.
intDays = 3

' Checks the number of days of validity of the password for the policy.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("DefaultNamingContext")
Set objDomain = GetObject("LDAP://" & strDNSDomain)
Set objMaxPwdAge = objDomain.MaxPwdAge

' Fix Bug IADslargeInteger.
lngHighAge = objMaxPwdAge.HighPart
lngLowAge = objMaxPwdAge.LowPart
If (lngLowAge < 0) Then
    lngHighAge = lngHighAge + 1
End If

' Convert of 100-nanoseconds to days.
sngMaxPwdAge = -((lngHighAge * 2^32) _
    + lngLowAge)/(600000000 * 1440)

' Determines the last password change.
' No precesses user whose password has expired.

'Defined manually when not using the default policy.
sngMaxPwdAge = 63 
dtmDate1 = DateAdd("d", - sngMaxPwdAge, Now())
dtmDate2 = DateAdd("d", intDays - sngMaxPwdAge, Now())

'Fetch Time Machine recording local time.
'This feature switches to daylight saving time.

Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngBias = lngBias + (lngBiasKey(k) * 256^k)
    Next
End If

' Convert Dates to UTC.
dtmDate1 = DateAdd("n", lngBias, dtmDate1)
dtmDate2 = DateAdd("n", lngBias, dtmDate2)

' Calculate second numbers since  1/1/1601.
lngSeconds1 = DateDiff("s", #1/1/1601#, dtmDate1)
lngSeconds2 = DateDiff("s", #1/1/1601#, dtmDate2)

' Convert second numbers to string
' and convert to 100-nanosecond intervals.
str64Bit1 = CStr(lngSeconds1) & "0000000"
str64Bit2 = CStr(lngSeconds2) & "0000000"

' Configure ADO objects.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection

' Define the domain.
strBase = "<LDAP://" & strDNSDomain & ">"

' Configure the filter
' Filters only user who want to receive notifications (alternativeNotification)
strFilter = "(&(objectCategory=person)(objectClass=user)" _
    & "(pwdLastSet>=" & str64Bit1 & ")" _
    & "(pwdLastSet<=" & str64Bit2 & ")" _
	& "(alternativeNotification=1)" _ 
    & "(!userAccountControl:1.2.840.113556.1.4.803:=2)" _
    & "(!userAccountControl:1.2.840.113556.1.4.803:=65536)" _
    & "(!userAccountControl:1.2.840.113556.1.4.803:=32)" _
    & "(!userAccountControl:1.2.840.113556.1.4.803:=48))"

' Define search attributes.
strAttributes = "cn, sAMAccountName,mail,proxyAddresses,pwdLastSet,alternativeMail,alternativeMobile, distinguishedName"

' Build Query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Execute Query.
Set adoRecordset = adoCommand.Execute

' Records.
Do Until adoRecordset.EOF
	distinguishedName = adoRecordset.Fields("distinguishedName").Value
	If InStr(UCase(distinguishedName),"DEMITIDOS") = 0 AND InStr(UCase(distinguishedName),"AFASTADOS") = 0 Then
		strName = adoRecordset.Fields("sAMAccountName").Value
		strCNName = adoRecordset.Fields("cn").Value
		strEmail_a = adoRecordset.Fields("alternativeMail").Value
		strSms = adoRecordset.Fields("alternativeMobile").Value
		' Determines when the password expires.
		If (TypeName(adoRecordset.Fields("pwdLastSet").Value) = "Object") Then
			Set objDate = adoRecordset.Fields("pwdLastSet").Value
			dtmPwdLastSet = Integer8Date(objDate, lngBias)
		Else
			dtmPwdLastSet = #1/1/1601#
		End If
		
		dtmExpires = DateAdd("d", sngMaxPwdAge, dtmPwdLastSet)
		
		' Send e-mail
		If (strEmail_a <> "") Then
			arrEmails_a = Split(strEmail_a,",")
			For Each strItem In arrEmails_a
				WScript.Echo strItem & " - " & strCNName
				strEmailNOC = strEmailNOC & strItem & " - " & strCNName & " - " & FormatDate(dtmExpires) & "<br>"
				Call SendEmailMessage(strItem, strCNName, strName, dtmExpires)
			Next
		End If

		' Send SMS 
		'If(strSms <> "") Then
		'	Call SendSmsMessage(strSms, strName, dtmExpires)
		'End If
	End If
    adoRecordset.MoveNext
Loop

' Send adminsitrative e-mail
Call SendNOCEmailMessage("admin@yourdomain",strEmailNOC)

' Close the connections.
adoRecordset.Close
adoConnection.Close

Function Integer8Date(ByVal objDate, ByVal lngBias)
    Dim lngAdjust, lngDate, lngHigh, lngLow
    lngAdjust = lngBias
    lngHigh = objDate.HighPart
    lngLow = objDate.LowPart
    If (lngLow < 0) Then
        lngHigh = lngHigh + 1
    End If
    If (lngHigh = 0) And (lngLow = 0) Then
        lngAdjust = 0
    End If
    lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
        + lngLow) / 600000000 - lngAdjust) / 1440
    On Error Resume Next
    Integer8Date = CDate(lngDate)
    If (Err.Number <> 0) Then
        On Error GoTo 0
        Integer8Date = #1/1/1601#
    End If
    On Error GoTo 0

End Function

Sub SendEmailMessage(ByVal strDestEmail, ByVal strCN, ByVal strNTName, ByVal dtmDate)
    Dim objMessage
	Dim strHTML
	
    If (strDestEmail = "") Then
        Exit Sub
    End If

    Set objMessage = CreateObject("CDO.Message") 
	objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "yourmailserver"

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	strHTML = "<!DOCTYPE HTML>"
	strHTML = strHTML & "<HTML>"
	strHTML = strHTML & "<HEAD>"
	strHTML = strHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF8"" />"
	strHTML = strHTML & "</HEAD>"
	strHTML = strHTML & "<BODY>"
	strHTML = strHTML & "<b>Dear(a) " & strCN & "</b><br><br>"
	
	strHTML = strHTML & "The password to your account DOMAIN\" & strNTName & " will expire on " & FormatDate(dtmDate) & ".<br><br>"
	
	strHTML = strHTML & "</BODY>"
	strHTML = strHTML & "</HTML>"
	
	objMessage.BodyPart.charset = "UTF-8"
    objMessage.Subject = "Your Company - Password Expiration Warning"
	objMessage.From="admin@yourcompany"
    objMessage.To = strDestEmail 
	objMessage.HTMLBody = strHTML
	objMessage.HTMLBodyPart.Charset = "UTF-8"
	
	objMessage.Configuration.Fields.Update
    objMessage.Send
	set objMessage=nothing
End Sub

Sub SendNOCEmailMessage(ByVal strDestEmail, ByVal strEmailNOC)
    Dim objMessage
	Dim strHTML
	
    If (strDestEmail = "") Then
        Exit Sub
    End If

    Set objMessage = CreateObject("CDO.Message") 
	objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "yourmailserverdomain"

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	strHTML = "<!DOCTYPE HTML>"
	strHTML = strHTML & "<HTML>"
	strHTML = strHTML & "<HEAD>"
	strHTML = strHTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF8"" />"
	strHTML = strHTML & "</HEAD>"
	strHTML = strHTML & "<BODY>"
		
	strHTML = strHTML & "<b>Password Expiration Notification Notice</b><br><br>"
	strHTML = strHTML & "<b>Users below will have their expired password in less than 3 days:</b><br><br>"
	
	strHTML = strHTML & strEmailNOC
	
	strHTML = strHTML & "<br><br><b>This e-mail is only informative, it does not require action.</b><br><br>"
	strHTML = strHTML & "</BODY>"
	strHTML = strHTML & "</HTML>"
	
	objMessage.BodyPart.charset = "UTF-8"
    objMessage.Subject = "Your Company - Password Expiration Warning"
	objMessage.From="admin@yourcompany"
    objMessage.To = strDestEmail 
	objMessage.HTMLBody = strHTML
	objMessage.HTMLBodyPart.Charset = "UTF-8"
	
	objMessage.Configuration.Fields.Update
    objMessage.Send
	set objMessage=nothing
End Sub

Function FormatDate(ByVal dtmDate)
    Dim strYear, strMonth, strDay, strHour, strMinute
    strYear = CStr(Year(dtmDate))
    strMonth = PadLeft(CStr(Month(dtmDate)), 2, "0")
    strDay = PadLeft(CStr(Day(dtmDate)), 2, "0")
	strHour = PadLeft(CStr(Hour(dtmDate)), 2, "0")
	strMinute = PadLeft(CStr(Minute(dtmDate)), 2, "0")
    FormatDate = strDay & "/" & strMonth & "/" & strYear & " as " & strHour & ":" & strMinute & " hs"
End Function

Function PadLeft(ByVal strValue, ByVal intSize, ByVal strMask)
    PadLeft = RIGHT(String(intSize, strMask) & strValue, intSize)
End Function

Sub SendSmsMessage(ByVal strDestMobile, ByVal strNTName, ByVal dtmDate)
    Dim strMessage
	Dim objConnection, objRecordSet

    If (strDestMobile = "") Then
        Exit Sub
    End If

    strMessage = "Your Company - The password to access your account YOURDOMAIN\" & strNTName & " will expire " & FormatDate(dtmDate)
	
	' Implement here your SMS routines.
End Sub

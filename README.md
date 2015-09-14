# ADPasswordExpirationNotification
Script to send email to users about password expiration

Please change this settings:
```
' Number of days the password expiration check.
intDays = 3

and

' Defined manually when not using the default policy.
sngMaxPwdAge = 63 

Others:
Created a new field in the AD (alternativeNotification), Filters only user who want to receive notifications.
You can implement or comment this:

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
    
```

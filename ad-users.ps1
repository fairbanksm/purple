# AD User Query Script
#  Author:  Meghan Fairbanks
# Contact:  meghan.fairbanks@transcore.com
# Updated:  12/1/21

#variables for mail message
$from = ""
$to = @("")
$subject = "Active Directory Users Query"
$body = "Monthly AD users query."
$smtpserver = ""

$filepath = "c:\script\ad-users.csv"

#get all enabled accounts
Get-ADUser -Filter 'enabled -eq $true' -Properties * | Select-Object Name, LastLogonDate, Enabled, PasswordExpired, AccountExpirationDate, whenCreated | Sort-Object LastLogonDate -Descending | Export-Csv $filepath -NoTypeInformation

#get members of domain admin security group
"`"`"`n`"Domain Admins`",`"LastLogonDate`",`"Enabled`",`"PasswordExpired`",`"AccountExpirationDate`",`"whenCreated`"" | out-file c:\script\ad-users.csv -Append -Encoding ASCII
Get-ADGroupMember -Identity "Domain Admins" -Recursive | %{Get-ADUser -Identity $_.distinguishedName -Properties *} | Sort-Object LastLogonDate -Descending | Select-Object Name, LastLogonDate, Enabled, PasswordExpired, AccountExpirationDate, whenCreated | Export-Csv $filepath -Append

#get members of dev admins group
"`"`"`n`"Dev Admins`",`"LastLogonDate`",`"Enabled`",`"PasswordExpired`",`"AccountExpirationDate`",`"whenCreated`"" | out-file c:\script\ad-users.csv -Append -Encoding ASCII
Get-ADGroupMember -Identity "Dev Admins" -Recursive | %{Get-ADUser -Identity $_.distinguishedName -Properties *} | Sort-Object LastLogonDate -Descending | Select-Object Name, LastLogonDate, Enabled, PasswordExpired, AccountExpirationDate, whenCreated | Export-Csv $filepath -Append

#send mail message with attachments
Send-MailMessage -From $from -To $to -Subject $subject -Body $body -SmtpServer $smtpserver -Attachments $filepath
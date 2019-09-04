Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


$recipient = "customerservice@mnpass.net"
$mailrecipients = @("meghan.fairbanks@transcore.com", "mike.solomonson@mnpass.net")
$smtpserver = "10.202.107.133"

#set days of the week - this script shoyuld be run on Friday
$monday = (get-date).AddDays(-4).ToString('MM/dd/yyy')
$tuesday = (get-date).AddDays(-3).ToString('MM/dd/yyy')
$wednesday = (get-date).AddDays(-2).ToString('MM/dd/yyy')
$thursday = (get-date).AddDays(-1).ToString('MM/dd/yyy')
$friday = (get-date).AddDays(-0).ToString('MM/dd/yyy')


#weekly morning numbers
$starttime1 = " 6:30"
$endtime1 = " 8:00"
$mornmessage = ("Count of emails sent to Customer Serveice between" + $starttime1 + " and" + $endtime1 + ".")

$mornmessage | Out-File C:\script\morning-count.txt
$monday | Out-File C:\script\morning-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($monday + $starttime1) -End ($monday + $endtime1) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\morning-count.txt -Append

$tuesday | Out-File C:\script\morning-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($tuesday + $starttime1) -End ($tuesday + $endtime1) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\morning-count.txt -Append

$wednesday | Out-File C:\script\morning-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($wednesday + $starttime1) -End ($wednesday + $endtime1) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\morning-count.txt -Append

$thursday | Out-File C:\script\morning-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($thursday + $starttime1) -End ($thursday + $endtime1) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\morning-count.txt -Append

$friday | Out-File C:\script\morning-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($friday + $starttime1) -End ($friday + $endtime1) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\morning-count.txt -Append


#weekly evening numbers
$starttime2 = " 17:00"
$endtime2 = " 18:00"
$evemessage = ("Count of emails sent to Customer Serveice between" + $starttime2 + " and" + $endtime2 + ".")

$evemessage | Out-File C:\script\evening-count.txt
$monday | Out-File C:\script\evening-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($monday + $starttime2) -End ($monday + $endtime2) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\evening-count.txt -Append

$tuesday | Out-File C:\script\evening-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($tuesday + $starttime2) -End ($tuesday + $endtime2) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\evening-count.txt -Append

$wednesday | Out-File C:\script\evening-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($wednesday + $starttime2) -End ($wednesday + $endtime2) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\evening-count.txt -Append

$thursday | Out-File C:\script\evening-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($thursday + $starttime2) -End ($thursday + $endtime2) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\evening-count.txt -Append

$friday | Out-File C:\script\evening-count.txt -Append
Get-MessageTrackingLog -ResultSize Unlimited -Start ($friday + $starttime2) -End ($friday + $endtime2) `
-Recipient $recipient -EventID RECEIVE | Measure-Object | Format-Table -Property count -AutoSize | Out-File C:\script\evening-count.txt -Append



#send mail message with attachments
$body = ("Weekly customer service email counts.")
Send-MailMessage -From meghan.fairbanks@mnpass.net -To $mailrecipients -Subject "Weekly Customer Service Count" -Body $body -SmtpServer $smtpserver -Attachments ("c:\script\morning-count.txt", "C:\script\evening-count.txt")
# purple

**ad-users.ps1**  
This script will query AD and create a .csv with users sorted by their last logon date.  The script also queries for date created, account expiration date (if applicable), as well as booleans for account enabled status and password expiration status.  Accounts are broken down into three sections.
* All accounts where enabled = true. 
* All members of the Dev Admins group.
* All members of the Domain Admins group.

**weekly-count.ps1**  
This script will query exchange for the $recipient mailbox and get a count of messages sent within the specified times. 

 


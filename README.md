# Get-Mailbox-And-Stats-Info-And-Output

There are 2 sample scripts in this repo:

- ```OutputMailboxAndStatsInfo.ps1``` => to dump mailbox information for ALL mailboxes on an organization, browsing database by database to avoid loading ALL mailboxes of an organization into a Powershell variable, mailbox statistics, and corresponding AD account info (such as SID, Is the account disabled)
> NOTE: you can customize the output file with your own preferences by changing the $OutputFile variable. You can remove the $strDate too if you don't want to append date/time information within your output file name.

```powershell
$strDate = Get-Date -Format "_MMddyyyy_HHmmss"
$OutputFile = "C:\temp\test_$StrDate.csv"
```

=> [Download OutputMailboxAndStatsInfo.ps1 here](https://raw.githubusercontent.com/SammyKrosoft/Get-Mailbox-And-Stats-Info-And-Output/main/OutputMailboxAndStatsInfo.ps1) (or from the repository)

- ```OutputMailboxesInfoFromSpecificOU.ps1``` => same as above, but only for mailboxes on a specific Organizational Unit, and no mailbox stats on this sample.
> NOTE: on this script targetting an OU, I didn't put any parameters (yet). Change/hard code the desired OU on the $strOU variable: 

```powershell
$strOrgUnit = "OU=CanadaUsers,DC=CanadaDrey,DC=ca"
```

=> [Download OutputMailboxesInfoFromSpecificOU.ps1 here](https://raw.githubusercontent.com/SammyKrosoft/Get-Mailbox-And-Stats-Info-And-Output/main/OutputMailboxesInfoFromSpecificOU.ps1) (or from the repository)

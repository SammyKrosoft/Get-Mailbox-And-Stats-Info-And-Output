# Get-Mailbox-And-Stats-Info-And-Output

There are 2 sample scripts in this repo:

- OutputMailboxAndStatsInfo.ps1 => to dump mailbox information for ALL mailboxes on an organization, browsing database by database to avoid loading ALL mailboxes of an organization into a Powershell variable, mailbox statistics, and corresponding AD account info (such as SID, Is the account disabled)

- OutputMailboxesInfoFromSpecificOU.ps1 => same as above, but only for mailboxes on a specific Organizational Unit, and no mailbox stats on this sample.
> NOTE: on this script targetting an OU, I didn't put any parameters (yet). Change/hard code the desired OU on the $strOU variable: 

```powershell
$strOrgUnit = "OU=CanadaUsers,DC=CanadaDrey,DC=ca"
```


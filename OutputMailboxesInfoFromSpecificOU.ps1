# Definition of output file name and path - First I create a variable holding a date string in "_MonthDateYear_HourMinutesSeconds" format
#Then I use that date string inside the file name string, which helps differentiating different runs
$strDate = Get-Date -Format "_MMddyyyy_HHmmss"
$OutputFile = "C:\temp\test_$StrDate.csv"

# Adding PowerShell Exchange Management tools into the session
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Setting the AD search scope to the root of the forest
Set-ADServerSettings -ViewEntireForest $true

# Setting the OU we want to export mailboxes from
$strOrgUnit = "OU=CanadaUsers,DC=CanadaDrey,DC=ca"

# Saving all database objects into a variable to browse all mailboxes database by database
#=> we will load many mailbox objects into a variable (aka memory) but only database by database instead of loading all mailboxes of the organization all at once
#which risks to exhaust the computer RAM
$Databases = Get-MailboxDatabase

#Setting up variables (counters, count of DBs, empty collection) before the loop, it will be easier to use progress bar later
$Dbprogresscounter = 0
$DatabasesCount = $Databases.Count
Write-Host "Found $DatabasesCount databases..." -ForegroundColor Green
$ObjectCollectionToExport = @()
# Start of the database loop
Foreach ($database in $Databases) {
    # Our progress bar with the database count, and the progress counter
    write-progress -Id 1 -Activity "Parsing databases" -Status "Now in database $($database.Name), $($DatabasesCount-$DBProgresscounter) databases left... ..." -PercentComplete $($Dbprogresscounter/$DatabasesCount*100)
    $Mailboxes = $null
    $Stats = $Null
    $User = $Null
    # Here we load all mailboxes of the current database being "scanned", with only a subset of properties => quicker, and uses less memory than if ew put the whole mailbox object into the variable
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $Database.Name -OrganizationalUnit $strOrgUnit -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"} | Select Name,Alias,PrimarySMTPAddress,RecipientTypeDetails,RecipientType, LitigationHoldEnabled, IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota, RetainDeletedItemsFor, UseDatabaseQuotaDefaults, SingleItemRecoveryEnabled, RecoverableItemsQuota, CustomAttribute1
    
    If ($Mailboxes -eq $null -or $Mailboxes -eq "") {
        Write-Host "No mailboxes found on database $($Database.Name) ... moving on to next database (if any)" -ForegroundColor DarkRed -BackgroundColor Yellow
    } Else {
    
        Write-Host "Found $($Mailboxes.count) mailboxes on database $($Database.name) ..." -ForegroundColor Green    
        # Same as before the database loop, setting up variables (counters, count of mailboxes, empty collection) before the loop, it will be easier to use progress bar later
        $mbxCount = $Mailboxes.count
        $mbxCounter = 0
        # Mailboxes loop to get stats and properties
        Foreach ($mbx in $Mailboxes) {
            # Another progress bar with the mailboxes count, which is child of the database progress bar, it will be hierarchically under the database progress bar
            write-progress -ParentId 1 -Activity "Getting mailbox stats..." -status "Getting stats for mailbox $($mbx.name), $($mbxCount-$mbxCounter) mailboxes left..." -PercentComplete $($mbxCounter/$mbxCount*100)
           # Getting mailbox' associated user account propertiess, also to later populate our PSCustomObject with user properties such as AD Alias or SID info...
            $user = Get-User $mbx.Name | Select SID, AccountDisabled
            # Creating the PSCustomObject, directly with the hash table with the properties we want to provision
            $Object = New-Object -TypeName PSObject -Property @{
                # Put the properties you require here ... in this example, it's a mix of properties of the mailbox itself, properties of the mailbox statistics, and properties of the user account associated with the mailbox
                # Mailbox properties we want to put into our PSCustomObject
                MAilboxName = $mbx.Name
                MailboxAlias = $mbx.Alias
                RecipientType = $mbx.RecipientTypeDetails
                Database = $mbx.Database
                PrimarySMTPAddress = $mbx.PrimarySMTPAddress
                CustomAttribute1 = $mbx.CustomAttribute1
                # Mailbox Statistics properties we want to put into our PSCustomObject
                # AD User information we want to put into our PSCustomObject
                ADAccountDisabled = $user.AccountDisabled
            }
            # Add the current PSCustomObject into the collection
            $ObjectCollectionToExport += $Object
            # Increment the mailbox counter to update the progress bar in next iteration of the Foreach Mailbox loop
            $mbxCounter++
        }
    }
    # Increment the database counter (parent of the mailbox counter) to update the progress bar in the next iteration of the Foreach Database loop
    $Dbprogresscounter++
}

# Exporting the object, specifying each properties to ensure we will have a custom order of the column on the resulting CSV file
$ObjectCollectionToExport | select MailboxName,MailboxAlias,PrimarySMTPAddress,CustomAttribute1,ADAccountDisabled | Export-Csv $OutputFile -NoTypeInformation -Encoding 'UTF8'
notepad $OutputFile

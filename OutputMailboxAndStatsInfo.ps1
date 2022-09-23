# Definition of output file name and path - 
$strDate = Get-Date -Format "_MMddyyyy_HHmmss"
$OutputFile = "C:\temp\test_$StrDate.csv"

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

Set-ADServerSettings -ViewEntireForest $true

$Databases = Get-MailboxDatabase

$Dbprogresscounter = 0
$DatabasesCount = $Databases.Count
Write-Host "Found $DatabasesCount databases..." -ForegroundColor Green
$ObjectCollectionToExport = @()
Foreach ($database in $Databases) {
    write-progress -Id 1 -Activity "Parsing databases" -Status "Now in database $($database.Name), $($DatabasesCount-$DBProgresscounter) databases left... ..." -PercentComplete $($Dbprogresscounter/$DatabasesCount*100)
    $Mailboxes = $null
    $Stats = $Null
    $User = $Null
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $Database.Name -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"} | Select Name,PrimarySMTPAddress,RecipientTypeDetails,RecipientType, LitigationHoldEnabled, IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota, RetainDeletedItemsFor, UseDatabaseQuotaDefaults, SingleItemRecoveryEnabled, RecoverableItemsQuota
    
    If ($Mailboxes -eq $null -or $Mailboxes -eq "") {
        Write-Host "No mailboxes found on database $($Database.Name) ... moving on to next database (if any)" -ForegroundColor DarkRed -BackgroundColor Yellow
    } Else {
    
        Write-Host "Found $($Mailboxes.count) mailboxes on database $($Database.name) ..." -ForegroundColor Green    
        $mbxCount = $Mailboxes.count
        $mbxCounter = 0
        Foreach ($mbx in $Mailboxes) {
            write-progress -ParentId 1 -Activity "Getting mailbox stats..." -status "Getting stats for mailbox $($mbx.name), $($mbxCount-$mbxCounter) mailboxes left..." -PercentComplete $($mbxCounter/$mbxCount*100)
            $stats = Get-MailboxStatistics $mbx.Name | Select Lastlogontime, TotalItemSize, Itemcount, TotalDeletedItemSize
            $user = Get-User $mbx.Name | Select SID, AccountDisabled
            $Object = New-Object -TypeName PSObject -Property @{
                RecipientType = $mbx.RecipientTypeDetails
                LitigationHoldEnabled = $mbx.LitigationHoldEnabled
                IssueWarningQuota = $mbx.IssueWarningQuota
                ProhibitSendQuota = $mbx.ProhibitSendQuota
                ProhibitSendReceiveQuota = $mbx.ProhibitSendReceiveQuota
                RetainDeletedItemsFor = $mbx.RetainDeletedItemsFor
                UseDatabaseQuotaDefaults = $mbx.UseDatabaseQuotaDefaults
                SingleItemRecoveryEnabled = $mbx.SingleItemRecoveryEnabled
                RecoverableItemsQuota = $mbx.RecoverableItemsQuota
                UseDatabaseRetentionDefaults = $mbx.UseDatabaseRetentionDefaults
                Database = $mbx.Database
                PrimarySMTPAddress = $mbx.PrimarySMTPAddress
                IsMailboxEnabled = $mbx.IsMailboxEnabled

                Lastlogontime = $stats.Lastlogontime
                TotalItemSize = $stats.TotalItemSize
                Itemcount = $stats.Itemcount
                TotalDeletedItemSize = $stats.TotalDeletedItemSize
                
                SID = $user.SID
                ADAccountDisabled = $user.AccountDisabled
                
            }
            $ObjectCollectionToExport += $Object
            $mbxCounter++
        }
    }
    $Dbprogresscounter++
}


$ObjectCollectionToExport | select "TotalItemSize", "ProhibitSendReceiveQuota", "LitigationHoldEnabled", "RetainDeletedItemsFor", "UseDatabaseQuotaDefaults", "UseDatabaseRetentionDefaults", "SID", "Database", "ProhibitSendQuota", "Lastlogontime", "RecoverableItemsQuota", "IssueWarningQuota", "RecipientType", "Itemcount", "SingleItemRecoveryEnabled", "TotalDeletedItemSize", "ADAccountDisabled", "IsMailboxEnabled", "PrimarySMTPAddress" | Export-Csv $OutputFile -NoTypeInformation -Encoding 'UTF8'

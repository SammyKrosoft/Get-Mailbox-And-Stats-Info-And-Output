<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>

#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()

#NOTE: This script requires the ActiveDirectory module to be installed and imported in the current session.
#It usually requires the following Windows feature to be installed:
#    PS:\> Install-WindowsFeature -Name "RSAT-AD-PowerShell" –IncludeAllSubFeature
#then to import ActiveDirectory module:
#    PS:\>Import-Module ActiveDirectory
#
# We can add a routing to check if ActiveDirectory module is already installed and/or imported (not done in this sample)

<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

#$OutputFile = "D:\Antonio-test\test.csv"
$strDate = Get-Date -Format "_MMddyyyy_HHmmss"
$OutputFile = "C:\temp\test_$StrDate.csv"

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

Set-ADServerSettings -ViewEntireForest $true

# Storing the forest root for future use with Get-ADUser
$RootADdef = ([ADSI]'LDAP://RootDSE').defaultNamingContext

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
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $Database.Name -Filter {RecipientTypeDetails -ne "DiscoveryMailbox,SystemMailbox"} | Select SamAccountName,Name,PrimarySMTPAddress,RecipientTypeDetails, LitigationHoldEnabled, IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota, RetainDeletedItemsFor, UseDatabaseQuotaDefaults, SingleItemRecoveryEnabled, RecoverableItemsQuota, IsMailboxEnabled, Database, UseDatabaseRetentionDefaults
    
    If ($Mailboxes -eq $null -or $Mailboxes -eq "") {
        Write-Host "No mailboxes found on database $($Database.Name) ... moving on to next database (if any)" -ForegroundColor DarkRed -BackgroundColor Yellow
    } Else {
    
        Write-Host "Found $($Mailboxes.count) mailboxes on database $($Database.name) ..." -ForegroundColor Green    
        $mbxCount = $Mailboxes.count
        $mbxCounter = 0
        Foreach ($mbx in $Mailboxes) {
            write-progress -ParentId 1 -Activity "Getting mailbox stats..." -status "Getting stats for mailbox $($mbx.name), $($mbxCount-$mbxCounter) mailboxes left..." -PercentComplete $($mbxCounter/$mbxCount*100)
            $stats = Get-MailboxStatistics $mbx.SamAccountName | Select @{Label=”Lastlogontime”;Expression={$_.Lastlogontime.tostring("yyyy-MM-dd")}}, TotalItemSize, Itemcount, TotalDeletedItemSize
            # Storing the $mbx.SamAccountName in a variable as it's easier to use within the Filter part of Get-ADUser -Filter rathehr than $mbx.property
            $MailboxIDForGetADUser = $mbx.SamAccountName
            # Calling Get-ADUser with SearchBase using the $RootADdef variable defined at the beginning of the script
            $user = Get-ADUser -SearchBase $RootADdef.ToString() -SearchScope Subtree -Properties * -Filter "Name -eq `"$MailboxIDForGetADUser`"" | Select @{Name="lastLogon";Expression={[datetime]::FromFileTime($_.'lastLogon')}}, sid, enabled, userAccountControl, PasswordExpired
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
                ADAccountStatus = $user.Enabled
                ADLastLogon = $user.lastLogon
                ADUserAccountControl = $user.userAccountControl
                PasswordExpired = $user.PasswordExpired

            }
            $ObjectCollectionToExport += $Object
            $mbxCounter++
        }
    }
    $Dbprogresscounter++
}


$ObjectCollectionToExport | select "TotalItemSize", "ProhibitSendReceiveQuota", "LitigationHoldEnabled", "RetainDeletedItemsFor", "UseDatabaseQuotaDefaults", "UseDatabaseRetentionDefaults", "SID", "Database", "ProhibitSendQuota", "Lastlogontime", "RecoverableItemsQuota", "IssueWarningQuota", "RecipientType", "Itemcount", "SingleItemRecoveryEnabled", "TotalDeletedItemSize", "ADAccountStatus", "IsMailboxEnabled", "PrimarySMTPAddress", "ADLastLogon", "ADUserAccountControl", "PasswordExpired" | Export-Csv $OutputFile -NoTypeInformation -Encoding 'UTF8'

<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>

#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
$TotalMinutes = $([math]::round($($StopWatch.Elapsed.TotalMinutes),2))
Write-Host "The script took $TotalMinutes to execute" 
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>

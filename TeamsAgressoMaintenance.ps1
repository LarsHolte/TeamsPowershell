############## TEAMS AGRESSO CREATION AND MAINTENANCE ################## 
# Automaintenance of Team memberships from Agresso Unit4 (role view). 
# 31.10.2019 - LH

# If debug = $true the script will output messages to host and not make any changes
$debug = $false

# Path to script
$scriptPath = "C:\Tasks\Teams-Agresso\"

# Agresso Unit4 Role = Team ID mapping list - add more as needed.
# TeamID is not stored elsewhere, so need to manually maintain this list as new Teams are added.
# When a new Team has been created add the orgID to the TeamID in the list below.
$orgRoleTeamId = @{"AD_ANS2_A833" = "c94cf5ce-40c2-4618-bc42-05f78561663c"} # Økonomi og organisasjon

# Local "database-file" of the groups ExternalDirectoryObjectId that already has been processed
$alreadyProcessedFile = $scriptPath + "AlreadyProcessedSharepointSharingTeams.txt"
$alreadyProcessed = Get-Content $alreadyProcessedFile

# SecurePassword files
$securePasswordFileAgresso = $scriptPath + "SecurePasswordAgresso.txt"
$securePasswordFileOffice365 = $scriptPath + "SecurePasswordSVC_O365TeamCreator.txt"

### Email report lines, for storing report/error messages
$emailReportItems = New-Object System.Collections.ArrayList

### Logger function (write log to a "log" directory in the script relative path)
function Write-Log { param( [string]$logText )
    if ($debug)
    { 
        Write-Host $logText
    }
    else
    {
        $logFullPath = (Split-Path $script:MyInvocation.MyCommand.Path) + "\log\" + (Get-Date -Format "yyyy-MM-dd") + ".txt" 
        $logLine = (Get-Date -Format "yyyy-MM-dd HH:mm:ss ") + $logText
        Write-Output $logLine | Out-File $logFullPath  -Append -Encoding utf8
    }
}

function SendSMTPMail { param( [string[]]$smtpBodyItems )
    $smtpFrom = "<Teams.Agresso.Maintenance@stavanger.kommune.no>"
    $smtpTo = "admin-adresse@stavanger.kommune.no"
    $smtpSubject = "Teams Agresso maintenance script report"
    $smtpServer = "smtp-serverhost"
    $smtpPort = "25"
    $smtpBody = "<p/>The following was reported when running the $($scriptPath)TeamEmpMaintenance.ps1 script on $env:COMPUTERNAME <p/><p/>"
    foreach($item in $smtpBodyItems)
    {
        $smtpBody += $item.ToString()
        $smtpBody += "<p/><p/>"
    }
    Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $smtpSubject -Body $smtpBody -BodyAsHtml -Encoding UTF8 -SmtpServer $smtpServer -Port $smtpPort
    Write-Log "INF: Sent admin email to $smtpTo containing $($emailReportItems.Count.ToString()) items."
}

Write-Log "INF: Script started"
$Error.Clear()

### Get externally stored secret key
$sqlCon = New-Object System.Data.SQLClient.SQLConnection("Data Source=SERVERHOST;Initial Catalog=SVGKOMM;Integrated Security=True;ApplicationIntent=ReadOnly")
$sqlCmd = New-Object System.Data.SQLClient.SQLCommand("SELECT config_value FROM config WHERE config_data = N'config_common_key'", $sqlCon)
$sqlCon.Open()
[byte[]]$Key = [System.Text.Encoding]::UTF8.GetBytes($sqlCmd.ExecuteScalar())
$sqlCon.Close()
if ($Key.Count -ne 32) # Failsafe
{
    $logMsg = "ERR: Failed to get key from SQL, verify server $serverName and database $databaseName is available."
    Write-Log $logMsg
    $emailReportItems.Add($logMsg)   
    SendSMTPMail($emailReportItems)
    Break
}

### Get all users from role view [AgrProdM5].[sk_ad_roller] that contains the role AD_ANS2_% (recursive membership in a level 2 department)
$serverName = "SERVERHOST"
$databaseName = "AgrProdM5"
$userName = "SQLUSER"
$securePassword = Get-Content $securePasswordFileAgresso | ConvertTo-SecureString -Key $Key
$credentialsAgresso = New-Object System.Management.Automation.PSCredential -ArgumentList $userName, $securePassword
$query = "SELECT bruker, rolle FROM sk_ad_roller WHERE (rolle like N'AD_ANS2_%')"
$connString = "Server=$serverName;Database=$databaseName;User ID=$($credentialsAgresso.GetNetworkCredential().UserName);Password=$($credentialsAgresso.GetNetworkCredential().Password);"
$dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$dataAdapter.SelectCommand = New-Object System.Data.SqlClient.SqlCommand ($query,$connString)
$commandBuilder = New-Object System.Data.SqlClient.SqlCommandBuilder $dataAdapter
$dtUsers = New-Object System.Data.DataTable
[void]$dataAdapter.Fill($dtUsers)
if ($dtUsers.Rows.Count -le 50) # Failsafe
{
    $ex = $Error[0]
    $logMsg = "ERR: Failed to get more than 50 users with role from SQL, verify server $serverName and database $databaseName is available. Last error reported was: $ex"
    Write-Log $logMsg
    $emailReportItems.Add($logMsg)   
    SendSMTPMail($emailReportItems)
    Break
}
# Also get organization structure from org view [AgrProdM5].[sk_ad_organisasjon] to lookup department names
$query = "SELECT orgenhet, visningsnavn FROM sk_ad_organisasjon"
$dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$dataAdapter.SelectCommand = New-Object System.Data.SqlClient.SqlCommand ($query,$connString)
$commandBuilder = New-Object System.Data.SqlClient.SqlCommandBuilder $dataAdapter
$dtOrg = New-Object System.Data.DataTable
[void]$dataAdapter.Fill($dtOrg)
if ($dtOrg.Rows.Count -le 500) # Failsafe
{
    $ex = $Error[0]
    $logMsg = "ERR: Failed to get more than 500 departments from SQL, verify server $serverName and database $databaseName is available. Last error reported was: $ex"
    Write-Log $logMsg
    $emailReportItems.Add($logMsg)   
    SendSMTPMail($emailReportItems)
    Break
}

### Office 365 credentials and modules loading
$credentialsOffice365 = New-Object System.Management.Automation.PSCredential("SVCKONTO@stavanger.kommune.no", $(Get-Content $securePasswordFileOffice365 | ConvertTo-SecureString -Key $Key))
# Active Directory module
Import-Module ActiveDirectory
# Microsoft Online
Import-Module MSOnline
Connect-MSOLService -Credential $credentialsOffice365
$so = New-PSSessionOption -IdleTimeout 600000
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credentialsOffice365 -Authentication Basic -AllowRedirection -SessionOption $so
$ImportResults = Import-PSSession $Session
# Microsoft Teams
Connect-MicrosoftTeams -Credential $credentialsOffice365
Start-Sleep -Seconds 15 # Because of errors with the first New-Team creation, sometimes resulting in: Error occurred while executing Code: GeneralException Message: Failed to start/restart provisioning of Team
# Sharepoint online
Connect-SPOService -Url https://stavangerkommune-admin.sharepoint.com -Credential $credentialsOffice365
# Test we are able to connect to and retrieve some information
$testGroup = Get-UnifiedGroup -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
if ([string]::IsNullOrEmpty($testGroup.Name))
{
    $logMsg = "ERR: Failed to connect to Teams, ending script execution"
    Write-Log $logMsg
    $emailReportItems.Add($logMsg)   
    SendSMTPMail($emailReportItems)
    Break
}

################## TEAM MAINTENANCE ##################
Write-Log "INF: Starting Teams maintenance"
foreach ($orgRole in $orgRoleTeamId.Keys)
{
    # Reset
    $Error.Clear()
    $teamId = $null
    $orgId = $null
    $orgName = $null
    $unifiedGroup = $null
    $team = $null

    $teamId = $orgRoleTeamId[$orgRole]
    $orgId = $orgRole.Split('_')[2] # orgRole = AD_ANS2_A833, orgId = A833 (should maybe been presented directly in the sql view from Agresso)
    $orgName = $dtOrg.Select("orgenhet = '$orgId'").visningsnavn

    Write-Log "INF: Verifying $orgName TeamId: $teamId orgId: $orgId"
        
    # Have to get both the unified group (for Exchange specific attributes), and the Team (for attributes like ShowInTeamsSearchAndSuggestions)
    $unifiedGroup = Get-UnifiedGroup -Identity $teamId
    if ($unifiedGroup -eq $null) # Team has probably been deleted in Office365. Add to e-mail report of items to investigate. (User may have accidentally deleted the Team)
    {
        $emailReportItems.Add("Team gr.$orgName no longer exists in Office 365. A user may have accidentally deleted the Team. Verify that the deletion was correct and then remove the Team from the maintenance script.")
        continue # Skip further processing of this Team
    }
    $team = Get-Team -GroupId $teamId 

    # Verify Team name is correct, and if not, rename Team back to orgName if any owners renamed the Team. Department Teams should only be renamed in Agresso Unit4
    if ($unifiedGroup.DisplayName -ne "gr.$orgName") # Team has been renamed by an owner, rename back to Agresso Unit4 organizational name
    {
        $unifiedGroupName = $unifiedGroup.DisplayName
        Write-Log "INF: Setting DisplayName to gr.$orgName from $unifiedGroupName"
        if($debug -ne $true)
        {
            Set-UnifiedGroup -Identity $teamId -DisplayName "gr.$orgName"
        }
    }
    if ($unifiedGroup.AccessType -ne "Private") # All groups that are automatically maintained should be private, check and reset if any owner changed this
    {
        Write-Log "INF: Setting AccessType to Private for gr.$orgName"
        if($debug -ne $true)
        {
            Set-UnifiedGroup -Identity $teamId -AccessType "Private"
        }
    }
    if ($unifiedGroup.HiddenFromExchangeClientsEnabled -ne $True) # Do not show Team in Outlook clients
    {
        Write-Log "INF: Setting HiddenFromExchangeClientsEnabled for gr.$orgName"
        if($debug -ne $true)
        {
            Set-UnifiedGroup -Identity $teamId -HiddenFromExchangeClientsEnabled
        }
    }
    if ($unifiedGroup.HiddenFromAddressListsEnabled -ne $True) # Do not show Team in Exchange Address Book
    {
        Write-Log "INF: Setting HiddenFromAddressListsEnabled for gr.$orgName"
        if($debug -ne $true)
        {
            Set-UnifiedGroup -Identity $teamId -HiddenFromAddressListsEnabled $True
        }
    }
    if ($team.ShowInTeamsSearchAndSuggestions -ne $False) # Do not show Team in search and suggestions
    {
        Write-Log "INF: Setting ShowInTeamsSearchAndSuggestions for gr.$orgName to False"
        if($debug -ne $true)
        {
            Set-Team -GroupId $teamId -ShowInTeamsSearchAndSuggestions $False
        }
    }
    $teamDescription = "Automatisk vedlikeholdt Team"
    if ($team.Description -ne $teamDescription) # Update description
    {
        Write-Log "INF: Setting Team description for gr.$orgName to $teamDescription"
        if($debug -ne $true)
        {
            Set-Team -GroupId $teamId -Description $teamDescription
        }
    }
}
Write-Log "INF: Finished Teams creation and policy maintenance"
################## TEAM MEMBERSHIPS MAINTENANCE ##################
Write-Log "INF: Starting Teams membership maintenance"
Start-Sleep -Seconds 10 # Allow new Teams to finish intialize before updating memberships
foreach ($orgRole in $orgRoleTeamId.Keys)
{
    # Reset
    $Error.Clear()
    $teamId = $null
    $orgId = $null
    $orgName = $null
    $unifiedGroup = $null
    $team = $null
    $agressoSamAccounts = $null
    $upnList  = $null
    $unifiedGroupOwnersUPN = $null
    $unifiedGroupMembersUPN = $null
    $ex = $null
    $logMsg = $null
    $teamName = $null

    $teamId = $orgRoleTeamId[$orgRole]
    $orgId = $orgRole.Split('_')[2] # orgRole = AD_ANS2_A833, orgId = A833 (should maybe been presented directly in the sql view from Agresso)
    $orgName = $dtOrg.Select("orgenhet = '$orgId'").visningsnavn
    $teamName = "gr.$orgName"

    Write-Log "INF: Updating $orgName TeamId: $teamId orgId: $orgId"

    # Get existing Team owners
    $unifiedGroupOwnersUPN = (Get-UnifiedGroupLinks -Identity $teamId -LinkType Owner | Select -ExpandProperty WindowsLiveID).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $orgName with Team ID $teamId skipped. Error while retrieving Team owners. Team either does not exists or is missing at least one owner! Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Get existing Team members
    $unifiedGroupMembersUPN = (Get-UnifiedGroupLinks -Identity $teamId -LinkType Members | Select -ExpandProperty WindowsLiveID).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $orgName with Team ID $teamId skipped. Error while retrieving Team members. Team either does not exists or is missing at least one member! Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Get existing Agresso Unit4 members (Agresso only stores SamAccountName, find the UserPrincipalName by lookup in Active Directory and add upn to list)
    $agressoSamAccounts = $dtUsers.Select("rolle = '$orgRole'").bruker 
    $upnMemberList = New-Object System.Collections.ArrayList
    foreach($usr in $agressoSamAccounts)
    {
        $usrAD = $null
        $usrAD = Get-ADUser -Identity $usr -Properties UserPrincipalName, msExchRecipientTypeDetails, Enabled -ErrorAction Continue
        if($Error) # An error occured, probably a "Cannot find an object with identity: 'SAMACCOUNTNAME' under: 'DC=svgkomm,DC=svgdrift,DC=no'." type message
        {
            $ex = $Error[0]
            if ($ex.FullyQualifiedErrorId -like "*ADIdentityNotFoundException*") # User was not found. Just log to file and clear Error, this is an expected error.
            {
                $logMsg = "INF: $usr from Agresso Unit4 with role $orgRole was not found in Active Directory. User has been deleted or not created yet. Can not add user to $teamName"
                Write-Log $logMsg
                $Error.Clear()
            }
            Continue # Skip this user
        }
        if ($usrAD.Enabled -and ($usrAD.msExchRecipientTypeDetails -eq "2147483648")) # Check if user is active and has Office 365 user mailbox (=Teams license)
        {
            $upnMemberList.Add($usrAD.UserPrincipalName.ToLower())
        }
        else # Just log that we are not adding the user to the team userlist. User has probably recently been deactivated
        {
            $logMsg = "INF: $usr from Agresso Unit4 with role $orgRole is either not enabled in Active Directory or is missing Office 365 mailbox. Can not add user to $teamName"
            Write-Log $logMsg
        }
    }
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $orgName with Team ID $teamId skipped. Error while looking up Agresso Unit4 members in Active Directory. Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    if($upnMemberList.Count -le 50) # Failsafe in case Active Directory lookups failed (and upnMemberList is empty), expecting more than 50 users in these Teams
    {
        $logMsg = "ERR: $orgName with Team ID $teamId skipped. Memberlist only contains $($upnMemberList.Count.ToString()) members, expecting more than 50."
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Add users that do not exist as members already
    foreach($user in $upnMemberList)
    {
        $Error.Clear()
        if(!$unifiedGroupMembersUPN.Contains($user))
        {
            if($debug -ne $true)
            {
                Add-TeamUser -GroupId $teamId -Role Member -User $user.ToString()
            }
            if($Error) 
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error adding member $user to $orgName Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Added member $user to $teamName"
            }
        }
    }
    # TODO:Add owners - this information is not presented from Agresso Unit4 and cannot be implemented

    # Remove owners that do not exist in Agresso Unit4 with the organization role
    foreach($user in $unifiedGroupOwnersUPN)
    {
        if(!$upnMemberList.Contains($user))
        {
            if($debug -ne $true)
            {
                Remove-UnifiedGroupLinks -Identity $teamId -LinkType Owners -Links $user -Confirm:$false
            }
            if($Error)
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error removing owner $user from $teamName Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Removed owner $user from $teamName"
            }
        }
    }
    # Remove members that do not exist in Agresso Unit4 with the organization role
    foreach($user in $unifiedGroupMembersUPN)
    {
        $Error.Clear()
        if(!$upnMemberList.Contains($user))
        {
            if($debug -ne $true)
            {
                Remove-UnifiedGroupLinks -Identity $teamId -LinkType Members -Links $user -Confirm:$false
            }
            if($Error) 
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error removing member $user from $teamName Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Removed member $user from $teamName"
            }
        }
    }
}
Write-Log "INF: Finished Teams membership maintenance"

################## TEAM DEFAULT SHARING PERMISSIONS - ALL TEAMS ##################
Write-Log "INF: Starting Teams emp default sharing permissions maintenance"
### AZURE AD
Import-Module AzureADPreview
Connect-AzureAD -Credential $credentialsOffice365

foreach ($orgRole in $orgRoleTeamId.Keys)
{
    $teamId = $orgRoleTeamId[$orgRole]

    # Set Teams default sharing permissions if not already set (setting "ExternalUserAndGuestSharing")
    if (!$alreadyProcessed.Contains($teamId)) # Skip if already processed
    {
        $Error.Clear()
        $spoSite = Get-SPOSite -Identity ((Get-UnifiedGroup -Identity $teamId).SharePointSiteURL)
        if($Error)
        {
            $ex = $Error[0]
            Write-Host "ERR: Error message was: $ex"
            Continue
        }
        $spoSite | Set-SPOSite -SharingCapability ExternalUserAndGuestSharing
        # Add the ExternalDirectoryObjectId to the "database-file" of Teams already processed
        Write-Output $teamId >> $alreadyProcessedFile
        $teamName = $spoSite.Title
        Write-Log "INF: TeamID $teamId default permissions setup. TeamName: $teamName"
    }
}
Write-Log "INF: Finished Teams default sharing permissions maintenance"

################## TEAM EMAIL REPORT TO ADMINS ##################
if ($emailReportItems.Count -gt 0)
{
    SendSMTPMail($emailReportItems)
}
Write-Log "INF: Script completed"

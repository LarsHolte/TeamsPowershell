################## TEAM CREATION AND MAINTENANCE ################## 

# Local "database-file" of the groups ExternalDirectoryObjectId that already has been processed
$alreadyProcessedFile = "C:\Tasks\Teams-Departments\AlreadyProcessedSharepointSharingTeams.txt"
$alreadyProcessed = Get-Content $alreadyProcessedFile

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

# Email report info. Send to admin if $emailReportItems.Count > 0
$smtpFrom = "<Teams.Maintenance@stavanger.kommune.no>"
$smtpTo = "admin-adresse@stavanger.kommune.no"
$smtpSubject = "Teams maintenance script report"
$smtpServer = "smtp-serverhost"
$smtpPort = "25"
$emailReportItems = New-Object System.Collections.ArrayList

Write-Log "INF: Script started"
$Error.Clear()

# Get all departments from SQL, as this location is synced with Active Directory and also contains a boolean true if department is Office 365 enabled
$serverName = "SERVERHOST"
$databaseName = "DATABASENAME"
$query = "SELECT department_name FROM departments WHERE (department_office365 = 1)"
$connString = "Server=$serverName;Database=$databaseName;Integrated Security=SSPI;"
$dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$dataAdapter.SelectCommand = New-Object System.Data.SqlClient.SqlCommand ($query,$connString)
$commandBuilder = New-Object System.Data.SqlClient.SqlCommandBuilder $dataAdapter
$dt = New-Object System.Data.DataTable
[void]$dataAdapter.Fill($dt)
[System.Collections.ArrayList]$allDepsFromSQL = $dt | Select-Object -ExpandProperty department_name
if ($allDepsFromSQL.Count -le 100) # Failsafe, should be about 300 departments
{
    $ex = $Error[0]
    Write-Log "ERR: Failed to get more than 100 departments from SQL, verify server $serverName and database $databaseName is available. Last error reported was: $ex"
    Break
}

### CREDENTIALS ###
$Username = "SVCKONTO@stavanger.kommune.no"
$secpasswd = ConvertTo-SecureString PWD -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($Username,$secpasswd)

### ACTIVE DIRECTORY ###
Import-Module ActiveDirectory

### MICROSOFT ONLINE ###
Import-Module MSOnline
Connect-MSOLService -Credential $Credential
$so = New-PSSessionOption -IdleTimeout 600000
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection -SessionOption $so
$ImportResults = Import-PSSession $Session

### MICROSOFT TEAMS ###
Connect-MicrosoftTeams -Credential $Credential
Start-Sleep -Seconds 15 # Because of errors with the first New-Team creation, sometimes resulting in: Error occurred while executing Code: GeneralException Message: Failed to start/restart provisioning of Team

### SHAREPOINT ONLINE ###
Connect-SPOService -Url https://stavangerkommune-admin.sharepoint.com -Credential $Credential

# Verify we are able to connect to and retrieve Group information
$testGroup = Get-UnifiedGroup -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
if ([string]::IsNullOrEmpty($testGroup.Name))
{
    Write-Log "ERR: Failed to connect to Teams, ending script execution"
    Break # End script
}

################## TEAM CREATION AND MAINTENANCE ##################
Write-Log "INF: Starting Teams creation and maintenance"
[System.Collections.ArrayList]$allDepsFromSQLClone = $allDepsFromSQL.Clone() # Using a clone to be able to remove items from the original array (when detecting departments that has disabled Teams)
foreach ($depFromSQL in $allDepsFromSQLClone)
{
    # Reset
    $Error.Clear()
    $depFromAD = $null
    $depTeamID = $null
    $depName = $null
    $unifiedGroup = $null
    $newTeamID = $null
    $unifiedGroup = $null
    $team = $null

    # Get the actual Active Directory group object (department)
    $depFromAD = Get-ADGroup -Filter {Name -eq $depFromSQL } -SearchBase 'OU=Departments,OU=SVGKOMM,OU=Organization,DC=svgkomm,DC=svgdrift,DC=no' -Properties DistinguishedName,SamAccountName,extensionAttribute3
    if ($depFromAD -eq $null) # Failsafe
    {
        $logMsg = "ERR: Failed to get Active Directory group object $depFromSQL. Departments in AD and SQL should match! Department has probably been deleted in AD but not in SQL table SVGKOMM.groups"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue # Skip this department
    }
    $depName = $depFromAD.Name
    # Custom extensionAttribute is used to store the actual GroupID of the Team, used for checking if department already has a Team, or if it has been renamed. If "0" then Team for the department is disabled.
    $depTeamID = $depFromAD.extensionAttribute3 
    if($depTeamID -eq $null) # If Active Directory object does not contain a Team ObjectID in extensionAttribute3
    {
        # Verify that no un-linked Team with the same name already exists
        $unifiedGroup = Get-UnifiedGroup -Identity "gr.$depName" -ErrorAction SilentlyContinue
        if($unifiedGroup -ne $null) # If True Team already exists -> update membership, if False Team does not exist -> create new Team)
        {
            $existingID = $unifiedGroup.ExternalDirectoryObjectId
            Write-Log "INF: TeamID found for un-linked department $depName. Creating link to Team $depName with TeamID $existingID"
            #TODO: Check if we have any first, and do manual verification before this is enabled
            #Set-ADGroup -Identity $depFromAD -Replace @{extensionAttribute3=$unifiedGroup.ExternalDirectoryObjectId} # Link the TeamID to the Active Directory department group
        }
        else # No Team exists. Create a new Team and link it to the AD group
        {
            Write-Log "INF: TeamID not found for $depName. Creating new Team: gr.$depName"
            $newTeamID = (New-Team -DisplayName "gr.$depName" -Visibility Private -Description "Alle $depName").GroupId

            if([string]::IsNullOrEmpty($newTeamID))
            {
                Write-Log "ERR: Team creation for $depName failed"
                Continue
            }
            Set-ADGroup -Identity $depFromAD -Replace @{extensionAttribute3=$newTeamID} # Link the TeamID to the Active Directory department group
            if($Error)
            {
                Write-Log "ERR: Team created but unable to update Active Directory department $depName with TeamID $newTeamID. AD department should update next time this script is run (name matching). Investigate if the problem persist"
            }
        }
    }
    else # Active Directory already contains the Team objectId
    {
        if ($depTeamID -eq "0")
        {
            Write-Log "INF: TeamID $depTeamID found on existing AD department $depName. Team for this department is disabled. Skipping."
            # Remove the department from $allDepsFromSQL so it is not processed in other parts of this script
            $allDepsFromSQL.Remove($depFromSQL)
            Continue
        }
        Write-Log "INF: TeamID $depTeamID found on existing AD department $depName"
        
        # Have to get both the unified group (for Exchange specific attributes), and the Team (for attributes like ShowInTeamsSearchAndSuggestions)
        $unifiedGroup = Get-UnifiedGroup -Identity $depTeamID
        if ($unifiedGroup -eq $null) # Team has probably been deleted in Office365. Add AD department to e-mail report of items to investigate. (User may have accidentally deleted the Team, or the AD department should also be deleted)
        {
            $emailReportItems.Add("$depName exists in Active Directory with Team ID $depTeamID, but the Team no longer exists in Office 365. A user may have accidentally deleted the Team, or the department should also be deleted.")
            continue # Skip further processing of this Team
        }
        $team = Get-Team -GroupId $depTeamID 

        # Verify Team name is correct, and if not, rename Team back to AD depName if any owners renamed the Team. Department Teams should only be renamed by renaming the AD Department group using Windows Adminstrasjon
        if ($unifiedGroup.DisplayName -ne "gr.$depName") # Team has been renamed by an owner, rename back to Active Directory department group name
        {
            $unifiedGroupName = $unifiedGroup.DisplayName
            Write-Log "INF: Setting DisplayName to gr.$depName from $unifiedGroupName"
            Set-UnifiedGroup -Identity $depTeamID -DisplayName "gr.$depName"
        }
        if ($unifiedGroup.AccessType -ne "Private") # All groups that are automatically maintained should be private, check and reset if any owner changed this
        {
            Write-Log "INF: Setting AccessType to Private for gr.$depName"
            Set-UnifiedGroup -Identity "gr.$depName" -AccessType "Private"
        }
        if ($unifiedGroup.HiddenFromExchangeClientsEnabled -ne $True) # Do not show Team in Outlook clients
        {
            Write-Log "INF: Setting HiddenFromExchangeClientsEnabled for gr.$depName"
            Set-UnifiedGroup -Identity "gr.$depName" -HiddenFromExchangeClientsEnabled
        }
        if ($unifiedGroup.HiddenFromAddressListsEnabled -ne $True) # Do not show Team in Exchange Address Book
        {
            Write-Log "INF: Setting HiddenFromAddressListsEnabled for gr.$depName"
            Set-UnifiedGroup -Identity "gr.$depName" -HiddenFromAddressListsEnabled $True
        }
        if ($team.ShowInTeamsSearchAndSuggestions -ne $False) # Do not show Team in search and suggestions
        {
            Write-Log "INF: Setting ShowInTeamsSearchAndSuggestions for gr.$depName to False"
            Set-Team -GroupId $depTeamID -ShowInTeamsSearchAndSuggestions $False
        }
        $teamDescription = "Automatisk vedlikeholdt Team"
        if ($team.Description -ne $teamDescription) # Update description
        {
            Write-Log "INF: Setting Team description for gr.$depName to $teamDescription"
            Set-Team -GroupId $depTeamID -Description $teamDescription
        }
        # Always set Team picture to the known auto-maintained department picture 
        # (TODO: THIS DOES NOT WORK. The service account needs to be member of the team to change picture, global admin is not enough..gg MS, re-write and use GraphAPI instead.. MS actually removed this command..)
        #Set-TeamPicture -GroupId $depTeamID -ImagePath C:\Tasks\Teams-Departments\TeamIcon.png
    }
}
Write-Log "INF: Finished Teams creation and policy maintenance"
################## TEAM MEMBERSHIPS MAINTENANCE ##################
Write-Log "INF: Starting Teams membership maintenance"
Start-Sleep -Seconds 10 # Allow new Teams to finish intialize before updating memberships
foreach ($depFromSQL in $allDepsFromSQL)
{
    # Init
    $Error.Clear()
    $depFromAD = $null
    $depTeamID = $null
    $unifiedGroupOwnersUPN = $null
    $unifiedGroupMembersUPN = $null
    $departmentMembersUPN = $null
    $depManagedBy = $null
    $departmentOwnersUPN = $null
    $ex = $null
    $logMsg = $null

    # Get information from Active Directory and Microsoft Teams
    $depFromAD = Get-ADGroup -Filter {Name -eq $depFromSQL } -SearchBase 'OU=Departments,OU=SVGKOMM,OU=Organization,DC=svgkomm,DC=svgdrift,DC=no' -Properties DistinguishedName,SamAccountName,ManagedBy,extensionAttribute3
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $depFromSQL department not found in Active Directory, but exists in SQL database SVGKOMM.groups. Departments in AD and SQL should match! Either create in AD, or delete in SQL. Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Custom extensionAttribute is used to store the actual GroupID of the Team, used for checking if department already has a Team, or if it has been renamed. If "0" then Team for the department is disabled.
    $depTeamID = $depFromAD.extensionAttribute3
    # Get existing Team owners
    $unifiedGroupOwnersUPN = (Get-UnifiedGroupLinks -Identity $depTeamID -LinkType Owner | Select -ExpandProperty WindowsLiveID).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $depFromSQL department with Team ID $depTeamID skipped. Error while retrieving Team owners. Team either does not exists or the AD role department contains no valid Office365 users (0 users returned). Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Get existing Team members
    $unifiedGroupMembersUPN = (Get-UnifiedGroupLinks -Identity $depTeamID -LinkType Members | Select -ExpandProperty WindowsLiveID).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $depFromSQL department skipped. Error while retrieving Team members. Team either does not exists, or the AD department contains no valid Office365 users. Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Get existing AD members
    $departmentMembersUPN = (Get-ADGroupMember -Identity $depFromSQL | Get-ADUser -Properties UserPrincipalName, msExchRecipientTypeDetails | Where {($_.msExchRecipientTypeDetails) -eq "2147483648"} | Select -ExpandProperty UserPrincipalName).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $depFromSQL department skipped. Error while retrieving AD department members. AD department either does not exists, or it contains no valid Office365 users. Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }
    # Get existing AD owners (members of .ManagedBy group in AD)
    $depManagedBy = $depFromAD.ManagedBy
    $departmentOwnersUPN = (Get-ADGroupMember -Identity $depManagedBy | Get-ADUser -Properties UserPrincipalName, msExchRecipientTypeDetails | Where {($_.msExchRecipientTypeDetails) -eq "2147483648"} | Select -ExpandProperty UserPrincipalName).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        $logMsg = "ERR: $depFromSQL department skipped. Error while retrieving AD role department members. Either the AD security group is missing the ManagedBy ROL_Department link, or it contains no valid Office365 users. Error message was: $ex"
        Write-Log $logMsg
        $emailReportItems.Add($logMsg)
        Continue
    }

    
    # Add users that do not exist as members already
    foreach($user in $departmentMembersUPN)
    {
        $Error.Clear()
        if(!$unifiedGroupMembersUPN.Contains($user))
        {
            Add-TeamUser -GroupId $depTeamID -Role Member -User $user.ToString()
            if($Error) 
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error adding member $user to $depFromSQL Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Added member $user to $depFromSQL"
            }
        }
    }
    # Add users that do not exists as owners
    foreach($user in $departmentOwnersUPN)
    {
        if(!$unifiedGroupOwnersUPN.Contains($user))
        {
            Add-TeamUser -GroupId $depTeamID -Role Owner -User $user.ToString()
            if($Error)
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error adding owner $user to $depFromSQL Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Added owner $user to $depFromSQL"
            }
        }
    }
    # Remove users that do not exist as member or owner (some users are only members of the AD department.ManagedBy role group)
    foreach($user in $unifiedGroupMembersUPN)
    {
        $Error.Clear()
        if(!$departmentMembersUPN.Contains($user) -and !$departmentOwnersUPN.Contains($user))
        {
            Remove-UnifiedGroupLinks -Identity $depTeamID -LinkType Members -Links $user -Confirm:$false
            if($Error) 
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error removing member $user from $depFromSQL Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Removed member $user from $depFromSQL"
            }
        
        }
    }
    # Remove users that do not exist as owner
    foreach($user in $unifiedGroupOwnersUPN)
    {
        if(!$departmentOwnersUPN.Contains($user))
        {
            Remove-UnifiedGroupLinks -Identity $depTeamID -LinkType Owners -Links $user -Confirm:$false
            if($Error)
            {
                $ex = $Error[0]
                $logMsg = "ERR: Error removing owner $user from $depFromSQL Exception was: $ex"
                Write-Log $logMsg
                $emailReportItems.Add($logMsg)
            }
            else
            {
                Write-Log "INF: Removed owner $user from $depFromSQL"
            }
        }
    }
}
Write-Log "INF: Finished Teams membership maintenance"

################## TEAM DEFAULT SHARING PERMISSIONS - ALL TEAMS ##################
Write-Log "INF: Starting Teams default sharing permissions maintenance"
### AZURE AD
Import-Module AzureADPreview
Connect-AzureAD -Credential $Credential

# All Teams have the prefix "gr."
$AzureADGroupIDs = Get-AzureADGroup -All $true | Where-Object {$_.ObjectType -eq "Group" -and $_.MailEnabled -eq $true -and $_.DisplayName -like 'gr.*'} | Select -ExpandProperty ObjectId
foreach ($depTeamID in $AzureADGroupIDs) 
{
    # Set Teams default sharing permissions if not already set (setting "ExternalUserAndGuestSharing")
    if (!$alreadyProcessed.Contains($depTeamID)) # Skip if already processed
    {
        $Error.Clear()
        $spoSite = Get-SPOSite -Identity ((Get-UnifiedGroup -Identity $depTeamID).SharePointSiteURL)
        if($Error)
        {
            $ex = $Error[0]
            Write-Host "ERR: Error message was: $ex"
            Continue
        }
        $spoSite | Set-SPOSite -SharingCapability ExternalUserAndGuestSharing
        # Add the ExternalDirectoryObjectId to the "database-file" of Teams already processed
        Write-Output $depTeamID >> $alreadyProcessedFile
        $teamName = $spoSite.Title
        Write-Log "INF: TeamID $depTeamID default permissions setup. TeamName: $teamName"
    }
}

Write-Log "INF: Finished Teams default sharing permissions maintenance"

################## TEAM EMAIL REPORT TO ADMINS ##################
if ($emailReportItems.Count -gt 0)
{
    $smtpBody = "<p/>The following error messages was reported when running the TeamDepartmentMaintenance.ps1 script on $env:COMPUTERNAME <p/><p/>"
    foreach($item in $emailReportItems)
    {
        $smtpBody += $item.ToString()
        $smtpBody += "<p/><p/>"
    }
    Send-MailMessage -From $smtpFrom -To $smtpTo -Subject $smtpSubject -Body $smtpBody -BodyAsHtml -Encoding UTF8 -SmtpServer $smtpServer -Port $smtpPort
}

Write-Log "INF: Script completed"

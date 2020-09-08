<#
---------------------------------------------------------------------------------------------
    Copyright © 2020  SWORD TECHNOLOGIES.  All rights reserved.
---------------------------------------------------------------------------------------------
    PowerShell Source Code

    NAME: Provision_Teams_And_Migrate_SharePoint.ps1
    AUTHOR: Benoît Jester - SWORD TECHNOLOGIES (benoit.jester@sword-group.com)
    DATE  : 19/03/2020

    USAGE:
        * Create Teams and channels, based on a csv file
        * Migrate SharePoint content in Teams channels
        * Migrate SharePoint content in Teams SharePoint sites
        * Migrate SharePoint content in SharePoint sites

    EXAMPLE: .\Provision_Teams_And_Migrate_SharePoint.ps1 -testOnly $false -migrationPhase "PoC" -teamsAndChannelsFilePath "C:\Temp\Scripting\Teams_SharePoint_Migration\testBJE.csv" -processMigration $true -processProvisioning $false
    
    PREREQUISITES: https://docs.microsoft.com/en-us/microsoftteams/private-channels-life-cycle-management#install-the-latest-teams-powershell-module-from-the-powershell-test-gallery
---------------------------------------------------------------------------------------------
#>

# Script parameters
param (
    [Parameter(Position=0,mandatory=$false)]
    [bool]$TestOnly = $false,
    [Parameter(Position=1,mandatory=$true)]
    [string]$TeamsAndChannelsFilePath,
    [Parameter(Position=2,mandatory=$true)]
    [string]$MigrationPhase,
    [Parameter(Position=3,mandatory=$false)]
    [switch]$ProcessProvisioning = $false,
    [Parameter(Position=4,mandatory=$false)]
    [switch]$ProcessMigration = $false
)

# Include "Functions" library
. ".\functions\Functions"

# Team parameters
$teamVisibility = "Private"
$teamAllowGiphy = $true
$teamGiphyContentRating = "Strict"
$teamAllowStickersAndMemes = $true
$teamAllowCustomMemes = $true
$teamAllowGuestCreateUpdateChannels = $false
$teamAllowGuestDeleteChannels = $false
$teamAllowCreateUpdateChannels = $false
$teamAllowDeleteChannels = $false
$teamAllowAddRemoveApps = $false
$teamAllowCreateUpdateRemoveTabs = $false
$teamAllowCreateUpdateRemoveConnectors = $false
$teamAllowUserEditMessages = $true
$teamAllowUserDeleteMessages = $true
$teamAllowOwnerDeleteMessages = $true
$teamAllowTeamMentions = $true
$teamAllowChannelMentions = $true
$teamShowInTeamsSearchAndSuggestions = $false

try {

    # Connect to Teams
    if ($ProcessProvisioning.IsPresent -or $ProcessMigration.IsPresent) {
        Connect-MicrosoftTeams -Credential $destinationMigrationCredentials
        Connect-ExchangeOnline -Credential $destinationMigrationCredentials
    }

    if ($ProcessMigration.IsPresent) {
        # Connect to SharePoint admin
        $pnpAdminConnection = Connect-PnPOnline -Url $targetTenantURL -Credentials $destinationMigrationCredentials
    }
} catch {
    exit
}

# Import csv data
$teamsAndChannelsToCreate = Import-Csv "$TeamsAndChannelsFilePath" -Delimiter ";" -Encoding 'UTF8' | `
                            Where-Object {$_."Migration Phase" -eq $MigrationPhase `
                            -and $_."Site To Migrate" -eq "Yes" `
                            -and $_."To Be Migrated" -ne "No, delete" `
                            -and $_."To Be Migrated" -ne "No, archive" `
                            -and $_."Migrate To" -ne ""}

##################################
# Process structure provisioning #
##################################
if (!$ProcessProvisioning.IsPresent -and !$ProcessMigration.IsPresent) {exit}

$teamsAndChannelsToCreate | ForEach-Object {

    $itemToProcess = @{
        MigrateTo = $_."Migrate To"
        SourceSiteUrl = $_."Source Site URL"
        TargetSiteUrl = $_."Target SPO Site URL"
        SourceSiteName = $_."Site Name"
        SourceListName = $_."List Name"
        ShouldMigrateToTeams = $_."Migrate To" -eq "Teams"
        ShouldMigrateToSPTeams = $_."Migrate To" -eq "ModernTeamSite (Teams)"
        ShouldMigrateToSPOnly = $_."Migrate To" -eq "ModernTeamSite (SPO only)"
        TargetTeamsName = $_."SPO / Teams Name"
        SourceContainerType = $_."Container Type"
        TargetChannelName = $_."Channel Name"
        TargetChannelIsPrivate = $_."Channel Is Private" -eq "Yes"
        TargetChannelOwner1 = $_."Channel Owner #1"
        TargetChannelOwner2 = $_."Channel Owner #2"
        TargetChannelOwner = ($_."Channel Owner #1") -or ($_."Channel Owner #2")
        TargetTeamsOwner1 = $_."Team Owner #1"
        TargetTeamsOwner2 = $_."Team Owner #2"
    }

    # Info
    Write-Verbose ""
    Write-Verbose ("======= Processing: " + $itemToProcess.SourceSiteName + "/" + $itemToProcess.SourceListName )
   
    #########################
    # Data consistency test #
    #########################
    if ($testOnly) {
        Write-Verbose "Testing data consistency for channel:" + $itemToProcess.TargetTeamsName + "/" + $itemToProcess.ChannelName
    }

    if (($itemToProcess.ShouldMigrateToTeams -or $itemToProcess.ShouldMigrateToSPTeams) -and [string]::IsNullOrEmpty($itemToProcess.TargetTeamsName -eq $null)){
        Write-Warning "Teams name is missing"
        if (!$testOnly){Continue}
    }

    if ($itemToProcess.ShouldMigrateToTeams -and [string]::IsNullOrEmpty($itemToProcess.TargetChannelName)) {
        Write-Warning "Channel name is missing"
        if (!$testOnly){Continue}
    }

    if ($itemToProcess.ShouldMigrateToTeams -and [string]::IsNullOrEmpty($itemToProcess.TargetChannelIsPrivate)) {
        Write-Warning "Channel privacy setting is missing"            
        if (!$testOnly){Continue}
    }

    if ($itemToProcess.TargetChannelIsPrivate -and [string]::IsNullOrEmpty($itemToProcess.TargetChannelOwner)) {
        Write-Warning "An owner is missing for the private channel"
        if (!$testOnly){Continue}
    }

    # Skip if test only
    if ($testOnly) {  Write-Verbose ""; Continue }

    ########################
    # Process provisioning #
    ########################
    if ($ProcessProvisioning.IsPresent) {

        if (!$itemToProcess.ShouldMigrateToSPOnly) {

            ###############
            # Create Team #
            ###############
            # Test if the Team already exists 
            $teamsName = $itemToProcess.TargetTeamsName
            $processedTeam = $null
            $processedTeam = MicrosoftTeams\Get-Team -DisplayName $teamsName #Sharegate & MicrosoftTeams modules have the same 'Get-Team' cmdlet so we need to precise which module to use 

            if ($processedTeam.Count -gt 1) {
                Write-Warning "Multiple Teams exist with the name '$teamsName'"
                Continue
            }

            # Create Team if necessary
            if (!$processedTeam) {
                # Info
                Write-Verbose "Creating team: $teamsName"
                
                # Create Team
                $processedTeam = New-Team -Displayname "$teamsName" -Visibility $teamVisibility
            
                # Wait
                Write-Verbose "Waiting 30 seconds ..."
                Start-Sleep -Seconds 30
                Write-Verbose "Team '$teamsName' successfully created"

                if ($processedTeam -eq $null) {
                    Write-Error "The Team '$teamsName' couldn't be created"
                    Continue
                }

            } else {
                Write-Warning "Team '$teamsName' already exists"            
            }

            # Set team settings
            Set-Team    -GroupId $processedTeam.GroupId `
                        -Visibility $teamVisibility `
                        -AllowGiphy $teamAllowGiphy `
                        -GiphyContentRating $teamGiphyContentRating `
                        -AllowStickersAndMemes $teamAllowStickersAndMemes `
                        -AllowCustomMemes $teamAllowCustomMemes `
                        -AllowGuestCreateUpdateChannels $teamAllowGuestCreateUpdateChannels `
                        -AllowGuestDeleteChannels $teamAllowGuestDeleteChannels `
                        -AllowCreateUpdateChannels $teamAllowCreateUpdateChannels `
                        -AllowDeleteChannels $teamAllowDeleteChannels `
                        -AllowAddRemoveApps $teamAllowAddRemoveApps `
                        -AllowCreateUpdateRemoveTabs $teamAllowCreateUpdateRemoveTabs `
                        -AllowCreateUpdateRemoveConnectors $teamAllowCreateUpdateRemoveConnectors `
                        -AllowUserEditMessages $teamAllowUserEditMessages `
                        -AllowUserDeleteMessages $teamAllowUserDeleteMessages `
                        -AllowOwnerDeleteMessages $teamAllowOwnerDeleteMessages `
                        -AllowTeamMentions $teamAllowTeamMentions `
                        -AllowChannelMentions $teamAllowChannelMentions `
                        -ShowInTeamsSearchAndSuggestions $teamShowInTeamsSearchAndSuggestions

            ######################
            # Add owners to Team #
            ######################
            $teamOwner1 = $itemToProcess.TargetTeamsOwner1
            Retry-Command -ScriptBlock {
                if ($teamOwner1) { 

                    if (Get-User $teamOwner1 -ErrorAction SilentlyContinue) {
                        Write-Verbose "Adding user '$teamOwner1' as Team owner ..."
                    
                        Add-TeamUser -GroupId $processedTeam.GroupId -User $teamOwner1 -Role Owner

                        Write-Verbose "User '$teamOwner1' successfully added"
                    } else {
                        Write-Warning "User '$teamOwner1' doesn't exist in the tenant. Skipping ..."
                    }

                }
            }

            $teamOwner2 = $itemToProcess.TargetTeamsOwner2
            Retry-Command -ScriptBlock {
                if ($teamOwner2) {

                    if (Get-User $teamOwner2 -ErrorAction SilentlyContinue) {
                        Write-Verbose "Adding user '$teamOwner2' as Team owner ..."
                        
                        Add-TeamUser -GroupId $processedTeam.GroupId -User $teamOwner2 -Role Owner

                        Write-Verbose "User '$teamOwner2' successfully added"
                    } else {
                        Write-Warning "User '$teamOwner2' doesn't exist in the tenant. Skipping ..."
                    }                        
                }
            }

            ###################
            # Create channels #
            ###################
            if ($itemToProcess.ShouldMigrateToTeams) {

                Write-Verbose ""

                # Test if the channel already exists
                $channelName = $itemToProcess.TargetChannelName
                $channel = (Get-TeamChannel -GroupId $processedTeam.GroupId | Where-Object { $_.DisplayName -eq $channelName })
                $teamChannelExists = $channel -ne $null
                $channelIsPrivate = $itemToProcess.TargetChannelIsPrivate

                # Create private channel if necessary
                if (!$teamChannelExists -and $channelIsPrivate) {

                    # Scenario: private channel
                    Write-Verbose "Creating private channel '$channelName' ..."

                    New-TeamChannel -GroupId $processedTeam.GroupId -DisplayName $channelName -MembershipType "Private" -Description $channelName

                    Write-Verbose "Waiting for private group creation ..."
                    Start-Sleep -Seconds 10
                    Write-Verbose "Private channel '$channelName' successfully created"

                    # Set private channel owner 1
                    Retry-Command -ScriptBlock {

                        $channelOwner1 = $itemToProcess.TargetChannelOwner1
                        
                        if ($channelOwner1) {

                            if (Get-User $channelOwner1 -ErrorAction SilentlyContinue) {
                                Write-Verbose "Adding user '$channelOwner1' as channel owner ..."
                                
                                Add-TeamUser -GroupId $processedTeam.GroupId -User $channelOwner1
                                Start-Sleep -Seconds 10
                                Add-TeamChannelUser -GroupId $processedTeam.GroupId -DisplayName "$channelName" -User $channelOwner1
                                Start-Sleep -Seconds 10
                                Add-TeamChannelUser -GroupId $processedTeam.GroupId -DisplayName "$channelName" -User $channelOwner1 -Role Owner
                                
                                Write-Verbose "User '$channelOwner1' successfully added" -ForegroundColor Green
                            } else {
                                Write-Verbose "User '$channelOwner1' doesn't exist in the tenant. Skipping ..."
                            }
                        }
                    }

                    # Set private channel owner 2
                    Retry-Command -ScriptBlock {

                        $channelOwner2 = $itemToProcess.TargetChannelOwner2
                        
                        if ($channelOwner2) {

                            if (Get-User $channelOwner2 -ErrorAction SilentlyContinue) {
                                # Info
                                Write-Verbose "Adding user '$channelOwner2' as channel owner ..."

                                Add-TeamUser -GroupId $processedTeam.GroupId -User $channelOwner2
                                Start-Sleep -Seconds 10
                                Add-TeamChannelUser -GroupId $processedTeam.GroupId -DisplayName "$channelName" -User $channelOwner2
                                Start-Sleep -Seconds 10
                                Add-TeamChannelUser -GroupId $processedTeam.GroupId -DisplayName "$channelName" -User $channelOwner2 -Role Owner
                                
                                # Info      
                                Write-Verbose "User '$channelOwner2' successfully added" -ForegroundColor Green
                            } else {
                                Write-Verbose "User '$channelOwner2' doesn't exist in the tenant. Skipping ..."
                            }
                        }
                    }
                    
                } elseif (!$teamChannelExists -and !$channelIsPrivate) {

                    # Scenario: public channel
                    Write-Verbose "Creating public channel '$channelName' ..."
                    New-TeamChannel -GroupId $processedTeam.GroupId -DisplayName "$channelName" -MembershipType "Standard" -Description "$channelName"
                    Write-Verbose "Public channel '$channelName' successfully created"

                } elseif ($teamChannelExists) {
                    Write-Warning "Channel '$channelName' already exists"       
                }  
            }     
        }
    }

    #####################
    # Process migration #
    #####################
    if ($ProcessMigration.IsPresent) {

        # Switch "Migration To"
        switch ($itemToProcess.MigrateTo) {

            'ModernTeamSite (SPO only)' {

                # Migrate SharePoint List as is without any target modification
                migrateSharePointList   -sourceSiteURL $itemToProcess.SourceSiteUrl `
                                        -targetSiteURL $itemToProcess.TargetSiteUrl `
                                        -sourceListName $itemToProcess.SourceListName `
                                        -targetListName $itemToProcess.SourceListName
            }
            'ModernTeamSite (Teams)' {

                $processedTeam = MicrosoftTeams\Get-Team -DisplayName $teamsName
                $teamMailNickname = $processedTeam.MailNickName
                $targetSiteURL = $null

                # Get the SharePoint site URL associated to the Teams
                $targetSiteURL = Get-PnPTenantSite -Connection $pnpAdminConnection -Template 'GROUP#0' -Filter "Url -eq '$baseTeamsSiteurl/$teamMailNickname'" | Select-Object -ExpandProperty Url

                if ($targetSiteURL) {

                    # Migrate SharePoint List
                    migrateSharePointList   -sourceSiteURL $itemToProcess.SourceSiteUrl `
                                            -targetSiteURL $targetSiteURL `
                                            -sourceListName $itemToProcess.SourceListName `
                                            -targetListName $itemToProcess.SourceListName
                } else {
                    Write-Verbose "The SharePoint site associated to the Team hasn't been found"
                }
            }
            'Teams' {

                # Get target site URL
                $teamsName = $itemToProcess.TargetTeamsName
                $channelName = $itemToProcess.TargetChannelName
                $channelIsPrivate = $itemToProcess.TargetChannelIsPrivate

                $processedTeam = MicrosoftTeams\Get-Team -DisplayName $teamsName
                $teamMailNickname = $processedTeam.MailNickName
                $targetSiteURL = $null

                # Need to perform an exact match
                $teamSiteUrl = Get-PnPTenantSite -Connection $pnpAdminConnection -Template 'GROUP#0' -Filter "Url -eq '$baseTeamsSiteurl/$teamMailNickname'" | Select-Object -ExpandProperty Url

                if ($channelIsPrivate) {

                    # Make sure the channel exists
                    $channel = (Get-TeamChannel -GroupId $processedTeam.GroupId -MembershipType Private | Where-Object { $_.DisplayName -eq "$channelName" })

                    if ($channel -and $teamSiteUrl) {
                        # Get the corresponding SharePoint site URL               
                        $privateChannelSites = Get-PnPTenantSite -Connection $pnpAdminConnection -Template 'TEAMCHANNEL#0' -Filter "Url -like '$teamSiteUrl'" | Where-Object {$_.Title -match $channelName}
                        
                        # It should retrieve only one site
                        if ($privateChannelSites.Length -eq 1) {
                            $targetSiteURL = $privateChannelSites[0].Url
                        }
                    }
                } else {
                    $targetSiteURL = $teamSiteUrl
                }
                
                if ($targetSiteURL) {

                    # Document libaries
                    if ($itemToProcess.SourceContainerType -eq "DocumentLibrary") {

                        # Migrate SharePoint source library to a folder inside the default library of the Team
                        migrateSharePointList   -sourceSiteURL $itemToProcess.SourceSiteUrl `
                                                -targetSiteURL $targetSiteURL `
                                                -sourceListName $itemToProcess.SourceListName `
                                                -targetListName "Documents" `
                                                -targetFolder $itemToProcess.TargetChannelName `
                                                -migrateToTeams $true
                    }

                    # Lists
                    if ($itemToProcess.SourceContainerType -eq "GenericList") {

                        # Migrate the list as is because list items can't be merged in the default document library
                        migrateSharePointList   -sourceSiteURL $itemToProcess.SourceSiteUrl `
                                                -targetSiteURL $targetSiteURL `
                                                -sourceListName $itemToProcess.SourceListName `
                                                -targetListName $itemToProcess.SourceListName
                    }


                } else {
                    Write-Verbose "The SharePoint site associated to the Team hasn't been found"
                }
            }
        }
    }
}

# Disconnect from Teams
if ($ProcessProvisioning.IsPresent){ Disconnect-MicrosoftTeams }

# Disconnect from Exchange Online
if ($ProcessMigration.IsPresent){Disconnect-ExchangeOnline -Confirm:$false}
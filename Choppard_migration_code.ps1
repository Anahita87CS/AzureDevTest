
#############
# Variables #
#############
# Connection variables
$migrationSourceUserName = "svc-sharegate-ww"
$migrationSourcePassword = "P@ss4Migr@teF1le"

#$migrationDestinationUserName = "ext-adm-fco@choptest.onmicrosoft.com"
#$migrationDestinationPassword ='"Fgtu67uu3ki!!'

$migrationDestinationUserName = "ext-adm-fco@choptest.onmicrosoft.com"
$migrationDestinationPassword = '"Fgtu67uu3ki!!'

$targetTenantURL = "https://choptest-admin.sharepoint.com"
$baseTeamsSiteurl = "https://choptest.sharepoint.com/sites"

######################
# Load PS Modules    #
######################
#if (!(Get-Module ExchangeOnlineManagement)) {
#    Install-Module ExchangeOnlineManagement -Force -AllowClobber
#}

#if (!(Get-Module MicrosoftTeams)) {
#    Install-Module MicrosoftTeams -Force -AllowClobber
#}

# Global variables
$reportPath = Join-Path (Get-Location) -ChildPath "Reports"

# Build credentials
[SecureString]$migrationSourceSecurePass = ConvertTo-SecureString $migrationSourcePassword -AsPlainText -Force 
[System.Management.Automation.PSCredential]$sourceMigrationCredentials = New-Object System.Management.Automation.PSCredential($migrationSourceUserName, $migrationSourceSecurePass)

[SecureString]$migrationDestinationSecurePass = ConvertTo-SecureString $migrationDestinationPassword -AsPlainText -Force 
[System.Management.Automation.PSCredential]$destinationMigrationCredentials = New-Object System.Management.Automation.PSCredential($migrationDestinationUserName, $migrationDestinationSecurePass)

######################
# ShareGate settings #
######################
# Set mapping settings
$mappingSettingsFilePath = Join-Path (Get-Location) -ChildPath ".\config\Chopard.sgum"
$mappingSettings = Import-UserAndGroupMapping -Path $mappingSettingsFilePath

# Set copy settings
$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate

# Set property template
$propertyTemplate = New-PropertyTemplate -AuthorsAndTimestamps -VersionHistory -WebParts -VersionLimit 40 -CheckInAs SameAsCurrent -ContentApproval SameAsCurrent

#################
# Retry-Command #
#################
function Retry-Command {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Position = 1, Mandatory = $false)]
        [int]$Maximum = 15,

        [Parameter(Position = 2, Mandatory = $false)]
        [int]$Delay = 10
    )

    Begin {
        $cnt = 0
    }

    Process {
        do {
            $cnt++
            try {
                $ScriptBlock.Invoke()
                return
            }
            catch {
                Write-Host "An error occured, retrying in 20 seconds ..." -ForegroundColor Yellow -ErrorAction Continue
                Write-Host $_.exception.Message -ForegroundColor Yellow 
                Start-Sleep -Seconds $Delay
            }
        } while ($cnt -lt $Maximum)

        # Info
        write-host "Execution failed" -ForegroundColor Red
    }
}

#########################
# migrateSharePointList #
#########################
function migrateSharePointList {
    param( $sourceSiteURL = $null,
        $targetSiteURL = $null,
        $sourceListName = $null,
        $targetListName = $null,
        $targetFolder = $null,
        $migrateToTeams = $null )

    # Clear result
    $result = $null

    # Clear source variables
    $sourceSite = $null
    $sourceList = $null

    # Clear target variables
    $targetSite = $null
    $targetList = $null

    # Info
    Write-Host "Performing migration from $sourceSiteURL/$sourceListName to $targetSiteURL/$targetListName/$targetFolder"
    write-host "-- Connecting to sites"

    # Connect to source site
    $sourceSite = Connect-Site -Url $sourceSiteURL -Credential $sourceMigrationCredentials

    # Connect to target site
    $targetSite = Connect-Site -Url $targetSiteURL -Credential $destinationMigrationCredentials
    
    write-host "-- Getting lists"
    # Get source list
    $sourceList = Get-List -Site $sourceSite -Name $sourceListName
    
    # If the source list doesn't exist
    if (!$sourceList) { 
        Write-Host ("The source list " + $listName + " doesn't exist") -ForegroundColor Yellow
    }
    else {
        # Info
        write-host "-- Launching migration"
        
        # Copy list
        if ($migrateToTeams) {
            # Copy list structure
            # 'NoCustomPermissions' is important here to not break default permissions inheritance made by Office 365 when creating channel folders
            Copy-List -NoCustomPermissions -NoCustomizedListForms -List $sourceList -DestinationSite $targetSite -ListTitleUrlSegment "Shared Documents" -ListTitle "Documents" -NoContent -MappingSettings $mappingSettings -CopySettings $copysettings -InsaneMode -WaitForImportCompletion -TaskName "Copy list to Teams (No content) - $($sourceListName)"

            # Get target list (should be only one)
            $targetList = Get-List -Site $targetSite -Name $targetListName | Where-Object { $_.BaseType -eq 'Document library' } | Select-Object -Unique

            if ($targetFolder) {
                # Copy content
                $result = Copy-Content -SourceList $sourceList -DestinationList $targetList -DestinationFolder $targetFolder -Template $propertyTemplate -CopySettings $copysettings -MappingSettings $mappingSettings -InsaneMode -WaitForImportCompletion -TaskName "Copy to channel - $($sourceListName) to $($targetFolder)"
            }
            else {
                # Copy content
                $result = Copy-Content -SourceList $sourceList -DestinationList $targetList -Template $propertyTemplate -CopySettings $copysettings -MappingSettings $mappingSettings -InsaneMode -WaitForImportCompletion -TaskName "Copy to channel - $($sourceListName) to $($targetList)"
            }
        }
        else {
            # Copy list
            $result = Copy-List -NoCustomizedListForms -List $sourceList -DestinationSite $targetSite -CopySettings $copysettings -MappingSettings $mappingSettings -InsaneMode -WaitForImportCompletion -TaskName "Migrate list - $($sourceListName)"
        }
    }


    # Export Report
    Export-Report $result -Path ($reportPath + $srcsite.Title) -Overwrite

    # Info
    Write-Host "Migration completed"
}

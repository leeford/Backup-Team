<#

.SYNOPSIS
 
    Backup-Team.ps1 - Backup and recreate a Team in Microsoft teams
    https://www.lee-ford.co.uk/backup-team
 
.DESCRIPTION
    Author: Lee Ford

    This tool allows you to backup certain aspects of a Team for safe keeping to be used to recreate (not restore) the Team. See https://www.github.com/leeford/Backup-Team for latest details.

    This tool has been written for PowerShell "Core" on Windows, Mac and Linux - it will not work with "Windows PowerShell".

    Note: A predefined Azure AD application has been used, but you can use your own with the same scope.

.LINK
    Blog: https://www.lee-ford.co.uk
    Twitter: http://www.twitter.com/lee_ford
    LinkedIn: https://www.linkedin.com/in/lee-ford/
 
.EXAMPLE 
    
    To backup a Team:
    Backup-Team.ps1 -Action Backup -Path <folder to store backup>

    To backup all Teams:
    Backup-Team.ps1 -Action Backup -Path <folder to store backups> -All

    To recreate a Team:
    Backup-Team.ps1 -Action Recreate -Path <path to backup file>

    To recreate all Teams from backups in a folder:
    Backup-Team.ps1 -Action Recreate -Path <path to backups> -All

    To backup a Team and accept (yes) to all actions:
    Backup-Team.ps1 -Action Backup -Path <folder to store backup> -YesToAll

    To backup a Team and exclude files as part of backup:
    Backup-Team.ps1 -Action Backup -Path <folder to store backup> -ExcludeFiles

    To recreate a Team and change the UPN suffix (e.g. recreating a different tenant with different UPN suffix):
    Backup-Team.ps1 -Action Recreate -Path <path to backup file> -ChangeUPNSuffix <e.g. domain.com>

#>

Param (

    [Parameter(mandatory = $true)][ValidateSet('Backup', 'Recreate')][string]$Action,
    [Parameter(mandatory = $false)][switch]$All,
    [Parameter(mandatory = $true)][string]$Path,
    [Parameter(mandatory = $false)][switch]$YesToAll,
    [Parameter(mandatory = $false)][string]$ChangeUPNSuffix,
    [Parameter(mandatory = $false)][switch]$ExcludeFiles

)

# Application (client) ID, resource and scope
$script:clientId = "95765b28-e4ee-40df-919a-10f1481538a8"
$script:scope = "Group.ReadWrite.All, User.ReadBasic.All"
$script:tenantId = "common"

function Invoke-GraphAPICall {

    param (

        [Parameter(mandatory = $true)][uri]$URI,
        [Parameter(mandatory = $false)][switch]$WriteStatus,
        [Parameter(mandatory = $false)][string]$Method,
        [Parameter(mandatory = $false)][string]$Body

    )

    # Is method speficied (if not assume GET)
    if ([string]::IsNullOrEmpty($method)) { $method = 'GET' }

    # Access token still valid?
    $currentEpoch = [int][double]::Parse((Get-Date (get-date).ToUniversalTime() -UFormat %s))

    if ($currentEpoch -gt [int]$script:token.expires_on) {

        Refresh-UserToken

    }

    $Headers = @{"Authorization" = "Bearer $($script:token.access_token)" }

    $currentUri = $URI

    $content = while (-not [string]::IsNullOrEmpty($currentUri)) {

        # API Call
        $apiCall = try {
            
            Invoke-RestMethod -Method $method -Uri $currentUri -ContentType "application/json; charset=UTF-8" -Headers $Headers -Body $body -ResponseHeadersVariable script:responseHeaders

        }
        catch {
            
            $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

        }
        
        $currentUri = $null
    
        if ($apiCall) {
    
            # Check if any data is left
            $currentUri = $apiCall.'@odata.nextLink'
    
            $apiCall
    
        }
    
    }

    if ($WriteStatus) {

        # If error returned
        if ($errorMessage) {

            Write-Host "FAILED $($errormessage.error.message)" -ForegroundColor Red

        }
        else {

            Write-Host "SUCCESS" -ForegroundColor Green

        }
        
    }

    return $content
    
}

function Invoke-FileUpload {

    param (

        [Parameter(mandatory = $true)][string]$filePath,
        [Parameter(mandatory = $true)][string]$destinationPath,
        [Parameter(mandatory = $true)][string]$teamId

    )

    try {

        Write-Host "    - Uploading file $filePath to $destinationPath... " -NoNewline

        # Get upload session
        while ([string]::IsNullOrEmpty($uploadSession.uploadUrl)) {

            $uploadSession = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$teamId/drive/root:/$($destinationPath):/createUploadSession" -Method "POST"

        }

        # Solution inspired by https://stackoverflow.com/questions/57563160/sharepoint-large-upload-using-ms-graph-api-erroring-on-second-chunk
        $chunkSize = 8192000 # Roughly 8 MB chunks

        # File information
        $fileInfo = New-Object System.IO.FileInfo($filePath)

        # Load file in to memory
        $reader = [System.IO.File]::OpenRead($filePath)

        # Buffer Array
        $buffer = New-Object -TypeName Byte[] -ArgumentList $chunkSize

        # Start at beginning of file
        $position = 0

        # First upload, so data is required
        $moreData = $true
        
        while ($moreData) {

            # Progress
            Write-Progress -Activity "Uploading File:" -Status "$($fileInfo.Name)" -CurrentOperation "$position/$($fileInfo.Length) bytes" -PercentComplete (($position / $fileInfo.Length) * 100)

            # Read chunk of data using buffer as an offset
            $bytesRead = $reader.Read($buffer, 0, $buffer.Length)
            $output = $buffer

            # If chunk is smaller than buffer length - no more data is needed
            if ($bytesRead -ne $buffer.Length) {

                $moreData = $false

                # Shrink the output array to the number of bytes
                $output = New-Object -TypeName Byte[] -ArgumentList $bytesRead
                [Array]::Copy($buffer, $output, $bytesRead)

            }

            # Upload chunk
            $Headers = @{

                #"Content-Length" = $output.Length # Not required in PS Core - it is automatically added to Headers!
                "Content-Range" = "bytes $position-$($position + $output.Length - 1)/$($fileInfo.Length)"

            }

            Invoke-WebRequest -Uri $uploadSession.uploadUrl -Method "PUT" -Headers $Headers -Body $output -SkipHeaderValidation | Out-Null

            # Set new position
            $position = $position + $output.Length

        }

        $reader.Close()

        Write-Host "SUCCESS" -ForegroundColor Green

    }
    catch {

        Write-Host "FAILED" -ForegroundColor Red
        $_

    }

}

function Get-UserToken {

    $script:token = $null

    $resource = "https://graph.microsoft.com/"

    $codeBody = @{ 

        resource  = $resource
        client_id = $script:clientId
        scope     = $script:scope
        

    }

    # Get OAuth Code
    $codeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$script:tenantId/oauth2/devicecode" -Body $codeBody

    # Print Code to host
    Write-Host "`n$($codeRequest.message)"

    $tokenBody = @{

        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        code       = $codeRequest.device_code
        client_id  = $clientId

    }

    # Get OAuth Token
    while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {

        $tokenRequest = try {

            Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$script:tenantId/oauth2/token" -Body $tokenBody

        }
        catch {

            $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

            # If not waiting for auth, throw error
            if ($errorMessage.error -ne "authorization_pending") {

                ThrowError

            }

        }

    }

    $script:token = $tokenRequest

}

function Refresh-UserToken {
    param (
        
    )

    $refreshBody = @{

        client_id     = $script:clientId
        scope         = "$script:scope offline_access" # Add offline_access to scope to ensure refresh_token is issued
        grant_type    = "refresh_token"
        refresh_token = $script:token.refresh_token

    }

    $tokenRequest = try {

        Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$script:tenantId/oauth2/token" -Body $refreshBody

    }
    catch {

        ThrowError
    
    }

    $script:token = $tokenRequest

}

function Backup-Team {
    param (

        [Parameter(mandatory = $true)][System.Object]$chosenTeam

    )

    if ($chosenTeam.id) {

        # Create Temp Backup Folder
        New-Item -Path "$Path/_Backup_Team_Temp_/" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

        # Start Transcript
        $date = Get-Date -UFormat "%Y-%m-%d %H%M"
        Start-Transcript -Path "$Path/_Backup_Team_Temp_/transcript.txt" | Out-Null

        Write-Host "`n Backing up Team '$($chosenTeam.displayName)'
            `r----------------------------------------------------------------------------------------------" -ForegroundColor Yellow

        # Group
        Write-Host " - Backing up Group Settings... " -NoNewline
        $groupSettings = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)" -WriteStatus

        # Team
        Write-Host " - Backing up Team Settings... " -NoNewline
        $teamSettings = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/teams/$($chosenTeam.id)" -WriteStatus

        # Owners
        Write-Host " - Backing up Owners... " -NoNewline
        $owners = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)/owners?`$select=id,displayName,userPrincipalName,mail" -WriteStatus 

        # Members
        Write-Host " - Backing up Members... " -NoNewline
        $members = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)/members?`$select=id,displayName,userPrincipalName,mail" -WriteStatus

        # Channels
        Write-Host " - Backing up Channels... " -NoNewline
        $channels = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/teams/$($chosenTeam.id)/channels" -WriteStatus

        # Tabs
        Write-Host " - Backing up Tabs... "
        $tabs = @()
        $channels.value | ForEach-Object {

            $channelId = $_.id

            Write-Host "    - Backing up Tabs for Channel '$($_.displayName)'... " -NoNewline
            $channelTabs = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/teams/$($chosenTeam.id)/channels/$($_.id)/tabs" -WriteStatus

            $channelTabs.value | ForEach-Object {

                $tab = @{

                    tab       = $_
                    channelId = $channelId

                }

                $tabs += New-Object PSObject -Property $tab

            }

        }

        # Build config in to single file
        $config = @{

            groupSettings = $groupSettings
            teamSettings  = $teamSettings
                    
            owners        = $owners
            members       = $members
        
            channels      = $channels
            tabs          = $tabs
        
        }
        
        Write-Host " - Saving configuration to teamConfig.json... " -NoNewline
        try {
                    
            $config | ConvertTo-Json -Depth 20 | Out-File "$Path/_Backup_Team_Temp_/teamConfig.json"
            Write-Host "SUCCESS" -ForegroundColor Green
        
        }
        catch {
        
            Write-Host "FAILED" -ForegroundColor Red
        
        }

        # Check Membership (required for files and conversations)
        Write-Host " - Checking Membership..."
        if ($config.members.value.id -notcontains $me.id -and $config.owners.value.id -notcontains $me.id) {

            $notOriginalMember = $true

            while ($addMember -notmatch "([Y|N])" -and -not $YesToAll) {

                ($addMember = Read-Host "`n*** ONLY AGREE TO THE BELOW IF YOU HAVE BEEN GRANTED PERMISSION TO READ THE CONTENTS OF THIS TEAM ***`nYou are not a member or owner of this Team. To read Conversations or Files, would you like to be like to become a member of the Team temporarily? [Y/N]").ToUpper() | Out-Null
        
            }

            if ($addMember -eq "Y" -or $YesToAll) {

                Write-Host " - Adding $($script:me.displayName) as a temporary Team member (to read conversations and files)... " -NoNewline

                $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($script:me.id)" } | ConvertTo-Json

                Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)/members/`$ref" -WriteStatus -Body $body -Method "POST"

                $membership = $true

                # Wait to be provisioned correctly
                Start-Sleep -Seconds 5
                
            }
            else {

                $membership = $false

            }

        }
        else {

            $membership = $true

        }

        # Conversations

        # Member of Team, so backup conversations
        if ($membership) {

            Write-Host " - Backing up Conversations... "

            # Create Conversations folder
            New-Item -ItemType Directory -Force -Path "$Path/_Backup_Team_Temp_/Conversations/" | Out-Null

            $channels.value | ForEach-Object {

                $Conversations = @()
                $channelId = $_.id
                $channelName = $_.displayName

                Write-Host "    - Backing up Conversations for Channel '$($_.displayName)'... " -NoNewline
                $channelConversations = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/teams/$($chosenTeam.id)/channels/$($_.id)/messages" -WriteStatus

                $channelConversations.value | Sort-Object -Property createdDateTime -Unique | ForEach-Object {

                    $replies = Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/teams/$($chosenTeam.id)/channels/$($channelId)/messages/$($_.id)/replies"

                    $message = @{

                        message   = $_
                        replies   = $replies.value
                        channelId = $channelId

                    }

                    $conversations += New-Object PSObject -Property $message

                }

                Write-Host "        - Saving Conversations to JSON $channelName.json... " -NoNewline
                try {
                        
                    $conversations | ConvertTo-Json -Depth 20 | Out-File "$Path/_Backup_Team_Temp_/Conversations/$channelName.json"
                    Write-Host "SUCCESS" -ForegroundColor Green
            
                }
                catch {
            
                    Write-Host "FAILED" -ForegroundColor Red
            
                }

                # HTML output of Conversations - inspired by https://github.com/veskunopanen/Teams-Graph-API/blob/master/WriteToOneNote.ps1
                $html = "<p>Below is a backup of the conversation history of this channel:</p>"

                $conversations | ForEach-Object {

                    $important = Check-MessageImportance $_.message
                    $attachments = Check-MessageAttachments $_.message.attachments
                    $edited = Check-MessageEdited $_.message
                    $deleted = Check-MessageDeleted $_.message

                    $html += "
                            <div class='card'>
                                <div class='card-header bg-light'><b>$($_.message.from.user.displayName)</b> $($_.message.createdDateTime) $edited</div>
                                $important
                                <div class='card-body'>
                                    <h4 class='card-title'>$($_.message.subject)</h4>
                                    <p>$($_.message.body.content)</p>
                                    $deleted
                                    $attachments
                                </div>"

                    if ($_.replies) {

                        $replyCount = ($_.replies).Count

                        $sortedReplies = $_.replies | Sort-Object -Property createdDateTime

                        $html += "
                                <ul class='list-group list-group-flush'>
                                <li class='list-group-item list-group-item-light'><span class='badge badge-success badge-pill'>$replyCount replies:</span></li>
                                "

                        $sortedReplies | ForEach-Object {

                            $important = Check-MessageImportance $_
                            $attachments = Check-MessageAttachments $_.attachments
                            $edited = Check-MessageEdited $_
                            $deleted = Check-MessageDeleted $_


                            $html += "
                                <li class='list-group-item'>
                                    <p class='card-title'><b>$($_.from.user.displayName)</b> $($_.createdDateTime) $edited</p>
                                    $important
                                    <p>$($_.body.content)</p>
                                    $deleted
                                    $attachments
                                </li>
                                "

                        }

                        $html += "</ul>"

                    }

                    $html += "</div><br />"

                }

                Write-Host "        - Saving Conversations to HTML $channelName.htm... " -NoNewline
                Create-HTMLPage -Content $html -PageTitle "$($chosenTeam.displayName) - $channelName - Conversations" -Path "$Path/_Backup_Team_Temp_/Conversations/$channelName.htm"

            }

        }
        else {

            Write-Host " - Excluding conversations..."

        }

        # Files
        if (-not $ExcludeFiles -and $membership) {
            
            Write-Host " - Backing up Files..."

            # List all items in drive
            $itemList = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)/drive/list/items?`$expand=DriveItem"

            # Loop through items
            $itemList.value | ForEach-Object {

                $item = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)/drive/items/$($_.DriveItem.id)"

                # If item can be downloaded
                if ($item."@microsoft.graph.downloadUrl") {

                    # Get path in relation to drive

                    $itemPath = $item.parentReference.path -replace "/drive/root:", ""
                    $fullFolderPath = "$Path/_Backup_Team_Temp_/Files/$itemPath" -replace "//", "/"
                    $fullPath = "$Path/_Backup_Team_Temp_/Files/$itemPath/$($item.name)" -replace "//", "/"

                    # Create folder to maintain structure
                    New-Item -ItemType Directory -Force -Path $fullFolderPath | Out-Null

                    # Download file
                    Write-Host "    - Saving $($item.name)... " -NoNewline
                    try {

                        Invoke-WebRequest -Uri $item."@microsoft.graph.downloadUrl" -OutFile $fullPath
                        Write-Host "SUCCESS" -ForegroundColor Green

                    }
                    catch {

                        Write-Host "FAILED" -ForegroundColor Red

                
                    }
            
                }

            }

        }
        else {

            Write-Host " - Excluding Files..."

        }

        # If temporarily added to Team as a member for backup, remove afterwards
        if ($notOriginalMember -and $membership) {

            Write-Host " - Removing $($script:me.displayName) as a temporary Team member (to read conversations and files)... " -NoNewline
            Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups/$($chosenTeam.id)/members/$($script:me.id)/`$ref" -WriteStatus -Method "DELETE"
        
        }

        # Stop Transcript
        Stop-Transcript | Out-Null

        # Create HTML Report

        # Channels
        $config.channels.value | ForEach-Object {

            $channelsReport += "<tr>
                                    <td>$($_.displayName)</td>
                                    <td>$($_.description)</td>
                                    <td><a class='btn btn-primary btn-sm m-1' role='button' href='./Files/$($_.displayName)' >Files</a></td>
                                    <td><a class='btn btn-primary btn-sm m-1' role='button' href='./Conversations/$($_.displayName).htm' >Conversations</a></td>
                                </tr>"

        }

        # Owners
        $teamMembersReport = $null

        $config.owners.value | ForEach-Object {

            $teamMembersReport += "<tr>
                                    <td>$($_.displayName)</td>
                                    <td>$($_.mail)</td>
                                    <td>Owner</td>
                                </tr>"
    
        }

        # Members
        $config.members.value | ForEach-Object {

            # Don't add if already an owner
            if ($config.owners.value.id -notcontains $_.id) {
    
                $teamMembersReport += "<tr>
                                        <td>$($_.displayName)</td>
                                        <td>$($_.mail)</td>
                                        <td>Member</td>
                                    </tr>"
    
            }
    
        }

        # Team Settings - Member
        $teamMemberSettingsReport = "<table class='table table-borderless'><tbody>"

        $config.teamSettings.memberSettings.PSObject.Properties | ForEach-Object {

            $teamMemberSettingsReport += "<th scope='row'>$($_.Name)</th><td>$($_.Value)</td></tr>"

        }

        $teamMemberSettingsReport += "</tbody></table>"

        # Team Settings - Guest
        $teamGuestSettingsReport = "<table class='table table-borderless'><tbody>"

        $config.teamSettings.guestSettings.PSObject.Properties | ForEach-Object {
        
            $teamGuestSettingsReport += "<th scope='row'>$($_.Name)</th><td>$($_.Value)</td></tr>"
        
        }
        
        $teamGuestSettingsReport += "</tbody></table>"

        # Team Settings - Messaging
        $teamMessagingSettingsReport = "<table class='table table-borderless'><tbody>"

        $config.teamSettings.messagingSettings.PSObject.Properties | ForEach-Object {
        
            $teamMessagingSettingsReport += "<th scope='row'>$($_.Name)</th><td>$($_.Value)</td></tr>"
        
        }
        
        $teamMessagingSettingsReport += "</tbody></table>"

        # Team Settings - Fun
        $teamFunSettingsReport = "<table class='table table-borderless'><tbody>"

        $config.teamSettings.funSettings.PSObject.Properties | ForEach-Object {
        
            $teamFunSettingsReport += "<th scope='row'>$($_.Name)</th><td>$($_.Value)</td></tr>"
        
        }
        
        $teamFunSettingsReport += "</tbody></table>"

        # Team Settings - Discovery
        $teamDiscoverySettingsReport = "<table class='table table-borderless'><tbody>"

        $config.teamSettings.DiscoverySettings.PSObject.Properties | ForEach-Object {
        
            $teamDiscoverySettingsReport += "<th scope='row'>$($_.Name)</th><td>$($_.Value)</td></tr>"
        
        }
        
        $teamDiscoverySettingsReport += "</tbody></table>"

        # Transcript
        Get-Content -Path "$Path/_Backup_Team_Temp_/transcript.txt" | ForEach-Object {

            $transcript += "$_<br />"

        }

        # Overall Backup Status 
        if ([string]$transcript -like "*FAILED*" -or [string]$transcript -like "*ERROR*") {

            $backupStatus = "FAILED - See transcript for errors"

        }
        else {

            $backupStatus = "SUCCESS"

        }

        # Build HTML
        $html = "
        <br /><div class='card'>
            <h5 class='card-header bg-light'>Team Overview</h5>
            <div class='card-body'>
            <table class='table table-borderless'>
            <tbody>
                <tr>
                    <th scope='row'>Backup Date:</th>
                    <td>$date</td>
                </tr>
                <tr>
                    <th scope='row'>Backup Status:</th>
                    <td>$backupStatus</td>
                </tr>
                <tr>
                    <th scope='row'>Name</th>
                    <td>$($config.groupSettings.displayName)</td>
                </tr>
                <tr>
                    <th scope='row'>Description</th>
                    <td>$($config.groupSettings.description)</td>
                </tr>
                <tr>
                    <th scope='row'>Mail</th>
                    <td>$($config.groupSettings.mail)</td>
                </tr>
                <tr>
                    <th scope='row'>Visibility</th>
                    <td>$($config.groupSettings.visibility)</td>
                </tr>
            </tbody>
          </table>
        </div>
        </div>
        <br /><div class='card'>
            <h5 class='card-header bg-light'>Team Membership</h5>
            <div class='card-body'>
            <table class='table table-borderless'>
            <thead>
                <tr>
                    <th scope='col'>Name</th>
                    <th scope='col'>Mail</th>
                    <th scope='col'>Role</th>
                </tr>
            </thead>
            <tbody>
                $teamMembersReport
            </tbody>
          </table>
        </div>
        </div>
        <br />
        <div class='card'>
            <h5 class='card-header bg-light'>Channel Overview</h5>
            <div class='card-body'>
          <table class='table table-borderless'>
          <thead>
                <th scope='col'>Name</th>
                <th scope='col'>Description</th>
                <th scope='col'>Files</th>
                <th scope='col'>Conversations</th>
          </thead>
          <tbody>
                $channelsReport
          </tbody>
        </table>
        </div>
        </div>
        <br />
        <div class='card'>
            <h5 class='card-header bg-light'>Team Settings</h5>
            <div class='card-body'>

                <div class='card-deck'>

                    <div class='card'>
                    <div class='card-header'>Member Settings</div>
                    <div class='card-body'>
                        $teamMemberSettingsReport
                    </div>
                    </div>

                    <div class='card'>
                    <div class='card-header'>Messaging Settings</div>
                    <div class='card-body'>
                        $teamMessagingSettingsReport
                    </div>
                    </div>

                    </div>

                    <br />
                    <div class='card-deck'>

                    <div class='card'>
                    <div class='card-header'>Fun Settings</div>
                    <div class='card-body'>
                        $teamFunSettingsReport
                    </div>
                    </div>

                    <div class='card'>
                    <div class='card-header'>Guest Settings</div>
                    <div class='card-body'>
                        $teamGuestSettingsReport
                    </div>
                    </div>

                    </div>

                    <br />
                    <div class='card-deck'>

                    <div class='card'>
                    <div class='card-header'>Discovery Settings</div>
                    <div class='card-body'>
                        $teamDiscoverySettingsReport
                    </div>
                    </div>

        </div>
        </div>
        </div>
                    <br />
                    <div class='card'>
                    <h5 class='card-header bg-light'>Backup Transcript</h5>
                    <div class='card-body'>
                    <p>$transcript</p>
                </div>
                </div>

"

        Write-Host " - Saving Backup Report Report.htm... " -NoNewline
        Create-HTMLPage -Content $html -PageTitle "$($chosenTeam.displayName) - Backup Report" -Path "$Path/_Backup_Team_Temp_/Report.htm"

        # Add Temp Backup Folder in to Zip
        $BackupFile = "BACKUP_TEAM_$($chosenTeam.displayName)_$date.zip" -replace '([\\/:*?"<>|\s])+', "_"

        Write-Host " - Adding files to zip file $BackupFile... " -ForegroundColor Yellow -NoNewline

        # Add all files to Zip
        $SaveBackupFile = try {
        
            Compress-Archive -Path "$Path/_Backup_Team_Temp_/*" -DestinationPath "$Path/$BackupFile" -CompressionLevel Optimal -Force
            Write-Host "SUCCESS" -ForegroundColor Green

        }
        catch {

            Write-Host "FAILED" -ForegroundColor Red
            Write-Host $SaveBackupFile -ForegroundColor Red

        }

        # Backup Status
        Write-Host "`nBackup Status of Team: $backupStatus"

        # Delete Temp Backup Folder
        Remove-Item -Path "$Path/_Backup_Team_Temp_/" -Force -Recurse | Out-Null


    }

}

function Recreate-Team {
    param (
        
        [Parameter(mandatory = $true)][string]$Path

    )

    try {

        # Extract Backup File
        Write-Host "`n - Extracting $Path... " -NoNewline

        $fileInfo = New-Object System.IO.FileInfo($path)
        $tempPath = "$($fileInfo.Directory)/_Recreate_Team_Temp_"

        Expand-Archive -Path $Path -DestinationPath $tempPath -Force
        Write-Host "SUCCESS" -ForegroundColor Green

    }
    catch {

        Write-Host "FAILED" -ForegroundColor Red
        break

    }
    
    # Load Config JSON
    Write-Host " - Loading teamConfig.json... " -NoNewline
    try {

        $config = Get-Content -Path "$tempPath/teamConfig.json" | ConvertFrom-Json
        Write-Host "SUCCESS" -ForegroundColor Green

    }
    catch {

        Write-Host "FAILED" -ForegroundColor Red
        break

    }

    # Prompt to build Team
    while ($createTeam -notmatch "([Y|N])" -and -not $YesToAll) {

        ($createTeam = Read-Host "`nWould you like to recreate Team '$($config.teamSettings.displayName)'? [Y/N]").ToUpper() | Out-Null

    }

    if ($createTeam -eq "Y" -or $YesToAll) {

        
        Write-Host "`n Creating Team '$($config.teamSettings.displayName)'
            `r----------------------------------------------------------------------------------------------" -ForegroundColor Yellow

        # Build Channels
        $channels = @()
        $config.channels.value | ForEach-Object {

            $channel = @{

                displayName = $_.displayName
                description = $_.description

            }

            $channels += New-Object PSObject -Property $channel

        }

        $body = @{

            "template@odata.bind" = "https://graph.microsoft.com/beta/teamsTemplates('standard')"
            displayName           = $config.teamSettings.displayName
            description           = $config.teamSettings.description
            visibility            = $config.groupSettings.visibility

            memberSettings        = $config.teamSettings.memberSettings
            guestSettings         = $config.teamSettings.guestSettings
            funSettings           = $config.teamSettings.funSettings
            messagingSettings     = $config.teamSettings.messagingSettings
            discoverySettings     = $config.teamSettings.discoverySettings

            channels              = $channels

        }

        $bodyJSON = $body | ConvertTo-Json

        # Create Team
        Write-Host " - Creating Team '$($config.teamSettings.displayName)'... " -NoNewline
        Invoke-GraphAPICall -URI "https://graph.microsoft.com/beta/teams" -WriteStatus -Method "POST" -Body $bodyJSON

        # Get created Team ID
        $matches = $null
        "$($script:responseHeaders.Location)" -match "\/teams\('([a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12})'\)\/operations\('([a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12})'\)" | Out-Null

        # If new Team ID exists
        if ($matches[1]) {
        
            $newTeamId = $matches[1]

            # Get New Team/Group, may need to keep trying as there can be a small delay in creating a Team
            while ([string]::IsNullOrEmpty($newTeam)) {

                $newTeam = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$newTeamId"

                Start-Sleep -Seconds 1

            }

            # Wait to ensure Team is fully provisioned
            Write-Host " - Waiting 10 seconds to allow Team to be provisioned fully..."
            Start-Sleep -Seconds 10

            # Members
            Write-Host " - Adding Members... "

            $existingMembers = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$($newTeam.id)/members"

            $config.members.value | ForEach-Object {

                # Change UPN Suffix Check
                $userId = Get-UserId $_

                Write-Host "    - Adding $($_.DisplayName)... " -NoNewline

                # Only add if not already an member
                if ($existingMembers.value.id -notcontains $userId) {

                    $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($userId)" } | ConvertTo-Json

                    Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$($newTeam.id)/members/`$ref" -WriteStatus -Body $body -Method "POST"

                }
                else {

                    Write-Host "SUCCESS" -ForegroundColor Green

                }

            }
            
            # Owners
            Write-Host " - Adding Owners... "

            $existingOwners = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$($newTeam.id)/owners"

            $config.owners.value | ForEach-Object {
    
                # Change UPN Suffix Check
                $userId = Get-UserId $_

                Write-Host "    - Adding $($_.DisplayName)... " -NoNewline

                # Only add if not already an owner
                if ($existingOwners.value.id -notcontains $userId) {

                    $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($userId)" } | ConvertTo-Json

                    Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$($newTeam.id)/owners/`$ref" -WriteStatus -Body $body -Method "POST"

                }
                else {

                    Write-Host "SUCCESS" -ForegroundColor Green

                }

            }

            # Files
            if (-not $ExcludeFiles) {

                if (Test-Path "$tempPath/Files") {

                    Write-Host " - Files directory found, uploading files..."

                    $files = Get-ChildItem -Recurse "$tempPath/Files" | Where-Object { ! $_.PSIsContainer }

                    # Current working directory (and change to forward slashes if not already)
                    $forwardSlashed = $tempPath -replace "\\", "/"
                    $regex = [Regex]::Escape("$forwardSlashed/Files/")

                    $files | ForEach-Object {

                        # Create destination path for file (based on current folder structure)
                        $destinationPath = $_.FullName -replace "\\", "/"
                        $destinationPath = $destinationPath -replace $regex, "/"

                        Invoke-FileUpload -filePath $_.FullName -destinationPath $destinationPath -teamId $newTeam.id

                    }

                }

            }

            # Previous Chats
            if (Test-Path "$tempPath/Conversations") {

                Write-Host " - Conversations directory found, uploading files..."

                $chatfiles = Get-ChildItem -Recurse "$tempPath/Conversations/"

                $forwardSlashed = $tempPath -replace "\\", "/"
                $regex = [Regex]::Escape("$forwardSlashed/Conversations/")

                $chatfiles | Where-Object { $_.name -like '*.htm' } | ForEach-Object {

                    # Find the file name to base folderpath on
                    $ChannelFolder = $_.name -replace ".htm", ""

                    # Create destination path for file (based on current folder structure)
                    $destinationPath = $_.FullName -replace "\\", "/"
                    $destinationPath = $destinationPath -replace $regex, "/"
                    $destinationPath = "/" + $ChannelFolder + "/Previous Conversations" + $destinationPath

                    Invoke-FileUpload -filePath $_.FullName -destinationPath $destinationPath -teamId $newTeam.id

                }

            }

            # Remove 'me' as a owner as part of creation (if not originally part of Team)
            if ($config.owners.value.id -notcontains $script:me.id) {

                Write-Host " - Removing 'me' as an owner (not part of owners in backup)... " -NoNewline
                Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/groups/$($newTeam.id)/owners/$($script:me.id)/`$ref" -Method "DELETE" -WriteStatus

            }

        }

    }

    # Delete Temp Backup Folder
    Remove-Item -Path $tempPath -Force -Recurse | Out-Null

}

function ChooseTeam {
    param (
        
    )

    $teams = $null

    # Ask for Team Name
    while ([string]::IsNullOrEmpty($teams)) {

        # Ask for Team name
        $teamName = Read-Host "`nWhat is the full or partial name of the Team? (Leave blank for all Teams)"

        # Get Groups
        $groups = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups?`$orderby=displayName" 

        # Filter down
        $teams = $groups.value | Where-Object { $_.DisplayName -like "*$teamName*" -and $_.ResourceProvisioningOptions -contains "Team" }

        if ([string]::IsNullOrEmpty($teams)) {

            Write-Host "No Teams found matching '$teamName'" -ForegroundColor Yellow

        }

    }

    # List Teams
    $teamsList = @()
    $i = 0

    $teams | ForEach-Object {

        $Team = @{

            Number      = $i
            DisplayName = $_.DisplayName
            Mail        = $_.mail
            id          = $_.id

        }

        $teamsList += New-Object PSObject -Property $Team

        $i++

    }

    # Count Teams
    Write-Host "`n$($teamsList.Count) Team(s) Found matching '$teamName':" -ForegroundColor Green

    $teamsList | Format-Table -Property Number, DisplayName, Mail | Out-Host

    # Prompt to choose Team
    $lastNumber = [int]$teamsList[-1].Number
    $chosenNumber = $null

    while ($chosenNumber -isnot [int] -or [int]$chosenNumber -gt [int]$lastNumber -or [int]$chosenNumber -lt [int]0) {

        $chosenNumber = Read-Host "Please choose a Team Number to backup [0-$lastNumber]"

        # Try to convert to int
        try {

            $chosenNumber = [int]$chosenNumber
            
        }
        catch {
           
        }

    }

    $chosenTeam = $teamsList[$chosenNumber]

    Write-Host "`n$($chosenTeam.Number) - '$($chosenTeam.DisplayName)' chosen..." -ForegroundColor Green

    return $chosenTeam
    
}

function Check-MessageImportance {

    param (

        [Parameter(mandatory = $true)][System.Object]$message

    )

    if ($message.importance -eq "high") {

        return "<div class='alert alert-danger m-2' role='alert'>IMPORTANT!</div>"
        
    }
    else {
        
        return $null

    }

}

function Check-MessageEdited {

    param (

        [Parameter(mandatory = $true)][System.Object]$message

    )

    if ($message.lastModifiedDateTime) {

        return "<b>Edited</b>"
        
    }
    else {
        
        return $null

    }

}

function Check-MessageDeleted {

    param (

        [Parameter(mandatory = $true)][System.Object]$message

    )

    if ($message.deletedDateTime) {

        return "<b>Message Deleted</b>"
        
    }
    else {
        
        return $null

    }

}

function Check-MessageAttachments {

    param (

        [Parameter(mandatory = $true)][System.Object]$messageAttachments

    )

    $attachments = $null

    if ($messageAttachments) {

        $messageAttachments | ForEach-Object {

            if ($_.contentType -eq "reference" -and $_.contentUrl -match "(https:\/\/.+\.sharepoint\.com\/sites\/.+\/Shared Documents\/)") {

                $link = $_.contentUrl -replace "(https:\/\/.+\.sharepoint\.com\/sites\/.+\/Shared Documents\/)", "../Files/"
                $attachments += "<a class='btn btn-primary btn-sm m-1' href='$link' role='button'>$($_.name)</a>"

            }

        }

        return $attachments
        
    }
    else {
        
        return $null

    }

}

function Get-UserId {
    param (
        
        [Parameter(mandatory = $true)][System.Object]$user

    )

    # Change UPN Suffix - Used for creating in New Tenant (Thanks Alexander!)
    if ($ChangeUPNSuffix) {

        # If a Guest user
        if ($user.userPrincipalName -like "*#EXT#*") {

            Write-Host "    - Getting 'New' User ID for $($_.DisplayName) (Guest User)... " -NoNewline
            $newUser = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/users/?`$filter=mail eq '$($user.mail)'" -WriteStatus
            return $newUser.value.id

        }
        else {
    
            # Get user ID in new tenant
            $newUPN = $user.userPrincipalName -replace '@(.*)', "@$ChangeUPNSuffix"
            Write-Host "    - Changing UPN Suffix - Getting 'New' User ID for $($_.DisplayName)... " -NoNewline
            $newUser = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/users/$($newUPN)" -WriteStatus
            return $newUser.id

        }
    
    }
    else {
    
        # No changes needed
        return $user.id
    
    }
    
}

function Create-HTMLPage {
    param (

        [Parameter(mandatory = $true)][string]$Content,
        [Parameter(mandatory = $true)][string]$PageTitle,
        [Parameter(mandatory = $true)][string]$Path

    )

    $html = "
    <div class='p-0 m-0' style='background-color: #F3F2F1'>
        <div class='container m-3'>
            <div class='page-header'>
                <h1>$pageTitle</h1>
                <h5>Created with <a href='https://www.lee-ford.co.uk/Backup-Team'>Backup-Team</a></h5>
            </div>

            $Content

            </div>
    </div>"

    try {
            
        ConvertTo-Html -CssUri "https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css" -Body $html -Title $PageTitle | Out-File $Path
        Write-Host "SUCCESS" -ForegroundColor Green

    }
    catch {

        Write-Host "FAILED" -ForegroundColor Red

    }

}

Write-Host "`n----------------------------------------------------------------------------------------------
            `n Backup-Team.ps1 - Lee Ford - https://www.lee-ford.co.uk
            `n----------------------------------------------------------------------------------------------" -ForegroundColor Yellow

# Get Azure AD User Token
Get-UserToken | Out-Null

# Get logged in User
$script:me = Invoke-GraphAPICall "https://graph.microsoft.com/v1.0/me"

if ($script:me.id) {

    Write-Host "`nSIGNED-IN as $($me.mail)" -ForegroundColor Green

    switch ($Action) {
        Backup {

            # Backup All Teams?
            if ($All) {

                # Get Groups
                $groups = Invoke-GraphAPICall -URI "https://graph.microsoft.com/v1.0/groups?`$orderby=displayName" 

                # Filter Teams
                $teams = $groups.value | Where-Object { $_.ResourceProvisioningOptions -contains "Team" }

                $counter = 0
                $teamsCount = ($teams).Count

                $teams | ForEach-Object {

                    $counter ++

                    # Progress
                    Write-Progress -Activity "Backing Up Team:" -Status "$counter out of $teamsCount" -CurrentOperation "$($_.displayName)"  -PercentComplete (($counter / $teamsCount) * 100)

                    Backup-Team $_

                }

            }
            else {
            
                # Initially Y on first run
                $anotherTeam = "Y"

                while ($anotherTeam -ne "N") {

                    if ($anotherTeam -eq "Y") {

                        $chosenTeam = ChooseTeam

                        Backup-Team $chosenTeam

                    }

                    ($anotherTeam = Read-Host "`nWould you like to backup another Team? [Y/N]").ToUpper() | Out-Null

                }

            }

        }

        Recreate {

            # Recreate All Teams?
            if ($All) {

                $backupFiles = Get-ChildItem -Path $path | Where-Object { $_.Name -like "BACKUP_TEAM_*.zip" }

                $counter = 0
                $fileCount = ($backupFiles).Count

                $backupFiles | ForEach-Object {

                    $counter ++

                    # Progress
                    Write-Progress -Activity "Opening Backup File:" -Status "$counter out of $fileCount" -CurrentOperation "$($_.Name)"  -PercentComplete (($counter / $fileCount) * 100)

                    Recreate-Team $_.FullName

                }

            }
            else {

                if (Test-Path $Path) {

                    Recreate-Team $Path

                }
                else {

                    Write-Host "Failed to open file, does this file exist?" -ForegroundColor Red

                }

            }

        }

    }

}
else {

    Write-Host "`nFAILED TO SIGNIN" -ForegroundColor Red

}
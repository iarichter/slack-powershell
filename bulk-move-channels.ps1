# Slack API URIs
$setTeamsURI = "https://slack.com/api/admin.conversations.setTeams"
$archiveURI = "https://slack.com/api/admin.conversations.archive"
$unArchiveURI = "https://slack.com/api/admin.conversations.unarchive"
$postMessageURI = "https://slack.com/api/chat.postMessage"
$joinConversationURI = "https://slack.com/api/conversations.join"

# Prompt for user token
$token = Read-Host -Prompt 'Input full user token (xoxp-)'

# Load CSV file containing channel information
$csvFile = "$PSScriptRoot\channelsToMove.csv"
$table = Import-Csv $csvFile -Delimiter ","

# Start logging the script output
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -Path "$PSScriptRoot\movingChannelsOutputLog.txt" -Append

foreach ($row in $table) {

    # Extract channel information from the current row
    $channel_id = $row.channel_id
    $team_id = $row.team_id
    $channel_name = $row.channel_name
    $target_team_ids = $row.target_team_ids
    $is_archived = $row.is_archived

    # Notification message to be posted in the channel
    $messageText = "Pardon our dust! IT is currently moving this channel into the new Product workspace where it will then be archived. You will probably get a number of notifications from this, sorry! Feel free to ignore anything on this channel for the time being and please reach out on the #slack-migration-help channel if you have any questions."

    if ($channel_id -ne "") {
        # Prepare request bodies for Slack API calls
        $moveBody = @{
            token = $token
            channel_id = $channel_id
            target_team_ids = $target_team_ids
            org_channel = $true
            team_id = $team_id
        }

        $removeOrgBody = @{
            target_team_ids = $target_team_ids
            org_channel = $false
            channel_id = $channel_id
            token = $token
        }

        # Used for both unarchive and archive POST requests
        $archiveBody = @{
            token = $token
            channel_id = $channel_id
        }

        # Used to alert the channel
        $postMessageBody = @{
            token = $token
            channel = $channel_id
            text = $messageText
        }

        # Used to add admin to conversations
        $joinConversationBody = @{
            token = $token
            channel = $channel_id
        }

        # Unarchive the channel if it is archived
        if ($is_archived -eq "TRUE") {
            Write-Host "Now UN-archiving:" $channel_name "("$channel_id")" -BackgroundColor Blue
            Invoke-RestMethod -Uri $unArchiveURI -Method POST -Body $archiveBody | Write-Host
            Start-Sleep -Milliseconds 3000
        } else {
            Write-Host "No need to Archive"
        }

        # Move the channel to the new workspace
        Write-Host "Moving and making org-channel:" $channel_name "("$channel_id")" -BackgroundColor Black -ForegroundColor Yellow
        Invoke-RestMethod -Uri $setTeamsURI -Method POST -Body $moveBody | Write-Host
        Start-Sleep -Milliseconds 3000

        # Add admin to the channel
        Write-Host "Adding IT.Accounts to channel:" $channel_name "("$channel_id")" -BackgroundColor Black -ForegroundColor Red
        Invoke-RestMethod -Uri $joinConversationURI -Method POST -Body $joinConversationBody | Write-Host
        Start-Sleep -Milliseconds 300

        # Notify the channel
        Write-Host "Post notification to:" $channel_name "("$channel_id")" -BackgroundColor Yellow -ForegroundColor Black
        Invoke-RestMethod -Uri $postMessageURI -Method POST -Body $postMessageBody | Write-Host
        Start-Sleep -Milliseconds 300

        # Remove the channel from organization-wide access
        Write-Host "Removing org-channel:" $channel_name "("$channel_id")" -BackgroundColor Red
        Invoke-RestMethod -Uri $setTeamsURI -Method POST -Body $removeOrgBody | Write-Host
        Start-Sleep -Milliseconds 3000

        # Archive the channel
        Write-Host "Now archiving:" $channel_name "("$channel_id")" -BackgroundColor Black
        Invoke-RestMethod -Uri $archiveURI -Method POST -Body $archiveBody | Write-Host
        Start-Sleep -Milliseconds 3000
    } else {
        Write-Host "Skipping row"
    }
}

# Stop logging
Stop-Transcript

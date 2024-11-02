# Slack API URI for archiving channels
$uri = "https://slack.com/api/admin.conversations.archive"

# Prompt for the Slack user token
$token = Read-Host -Prompt 'Input full user token (xoxp-)'

# Load CSV file containing channels to archive
$csvFile = "$PSScriptRoot\channelsToArchive.csv"
$table = Import-Csv $csvFile -Delimiter ","

# Start logging the script output
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -Path "$PSScriptRoot\archiveOutputLog.txt" -Append

foreach ($row in $table) {
    # Extract channel information from the current row
    $channel_id = $row.channel_id
    $channel_name = $row.channel_name

    # Display the channel being archived
    Write-Host "Now Archiving:" $channel_name "("$channel_id")" -BackgroundColor Black

    # Prepare the request body for the Slack API call
    $body = @{
        token = $token
        channel_id = $channel_id
    }

    # Archive the channel
    Invoke-RestMethod -Uri $uri -Method POST -Body $body

    # Pause briefly before processing the next channel
    Start-Sleep -Milliseconds 500
}

# Stop logging
Stop-Transcript

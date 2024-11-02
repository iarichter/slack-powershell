# Slack API URI for renaming channels
$uri = "https://slack.com/api/conversations.rename"

# Prompt for the Slack user token
$token = Read-Host -Prompt 'Input full user token (xoxp-)'

# Load CSV file containing channels to rename
$csvFile = "$PSScriptRoot\channelsToRename.csv"
$table = Import-Csv $csvFile -Delimiter ","

# Start logging the script output
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -Path "$PSScriptRoot\renameOutputLog.txt" -Append

foreach ($row in $table) {
    # Extract channel information from the current row
    $channel_id = $row.channel_id
    $channel_name = $row.channel_name
    $channel_new_name = $row.new_name

    # Display the channel being renamed
    Write-Host "Now renaming:" $channel_name "("$channel_id") to " $channel_new_name -BackgroundColor Black

    # Pause briefly before making the API call
    Start-Sleep -Milliseconds 500

    # Prepare the request body for the Slack API call
    $body = @{
        token = $token
        channel = $channel_id
        name = $channel_new_name
    }

    # Rename the channel
    Invoke-RestMethod -Uri $uri -Method POST -Body $body
}

# Stop logging
Stop-Transcript

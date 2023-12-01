<#
Copyright 2023 Jake Gwynn

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

# Set the security protocol to TLS 1.2 to prevent errors when connecting to the Graph API
[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

# Define the client secret, tenant ID, and application ID for API authentication
# It is recommended to secure your client secret by storing it in an encrypted file.  https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.secretmanagement/?view=ps-modules
$ClientSecret = "YOUR_CLIENT_SECRET"
$TenantId = "YOUR_TENANT_ID"
$AppId = "YOUR_APPLICATION_ID"

# Can be between 1 and 1000
$NumberOfEmailsToExport = 5
# Path to export emails to. Each mailbox will have its own folder.
$EmailExportPath = "C:\Temp\ExportedEmails"
# Path to CSV file containing shared mailboxes. The CSV file should have a column named UserPrincipalName. 
$SharedMailboxList = "C:\Temp\SharedMailboxes.csv"

# Initialize variables for token timer and access token
$TokenTimer = $null
$AccessToken = $null

# Function to handle REST API errors
function Get-RestApiError ($RestError) {
    # If the error is a WebException, return the response body from the exception
    if ($RestError.Exception.GetType().FullName -eq "System.Net.WebException") {
        $ResponseStream = $null
        $Reader = $null
        $ResponseStream = $RestError.Exception.Response.GetResponseStream()
        $Reader = New-Object System.IO.StreamReader($ResponseStream)
        $Reader.BaseStream.Position = 0
        $Reader.DiscardBufferedData()
        return $Reader.ReadToEnd();
    }
}

# Function to authenticate with the Microsoft Graph API using a client secret
function Connect-GraphApiWithClientSecret ($TenantId, $AppId, $ClientSecret) {
    if($global:TokenTimer -eq $null -or $global:TokenTimer.elapsed.minutes -gt '55'){
        try{
            Write-Host "Authenticating to Graph API"
            # Prepare request body for acquiring an access token
            $Body = @{    
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                client_Id     = $AppId
                Client_Secret = $ClientSecret
                } 
            # Request access token from Microsoft online login endpoint
            $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $Body
            # Start a timer to track token validity
            $global:TokenTimer =  [system.diagnostics.stopwatch]::StartNew()	
            return $ConnectGraph.access_token
        }
        catch {
            $RestError = $null
            $RestError = Get-RestApiError -RestError $_
            Write-Host $_ -ForegroundColor Red
            return Write-Host $RestError -ForegroundColor Red 
        }
    }
    else {
        return $global:AccessToken
    }
}

# Function to get recent emails from a mailbox using the Graph API
Function Get-RecentEmails {
    Param (
        [string]$AccessToken,
        [string]$MailboxId
    )

    # Construct the Graph API endpoint URL for retrieving emails
    $EmailsEndpoint = "https://graph.microsoft.com/v1.0/users/$MailboxId/messages?`$orderby=receivedDateTime DESC&`$top=$NumberOfEmailsToExport"
    
    # Prepare the authorization headers for the API request
    $Headers = @{
        Authorization = "Bearer $AccessToken"
    }

    # Call the Graph API to retrieve the most recent emails from the mailbox
    $Emails = Invoke-RestMethod -Uri $EmailsEndpoint -Headers $Headers -Method Get
    return $Emails.value
}

# Function to export emails to HTML files
Function Export-Emails {
    Param (
        [string]$MailboxUPN,
        [object]$Emails
    )

    # Create a directory for storing emails from the specified mailbox
    $DirectoryPath = "$EmailExportPath\$MailboxUPN"
    $NewItem = New-Item -Path $DirectoryPath -ItemType Directory -Force

    # Iterate over each email and save its content to a file
    foreach ($Email in $Emails) {
        $EmailContent = $Email.body.content
        $EmailSubject = $Email.subject -replace '[\\/*?:"<>|]'
        $FilePath = "$DirectoryPath/$EmailSubject.html"

        # Convert SentDateTime to local time, including the name of the local time zone
        $SentDateTime = [DateTime]::Parse($Email.sentdatetime).ToLocalTime()

        # Extract additional properties
        $Sender = $Email.sender.emailaddress.name + " (" + $Email.sender.emailaddress.address + ")"
        $ToRecipients = $Email.torecipients | ForEach-Object { $_.emailaddress.name + " (" + $_.emailaddress.address + ")" }
        $ToRecipients = $ToRecipients -join ', '
        $CcRecipients = $Email.ccrecipients | ForEach-Object { $_.emailaddress.name + " (" + $_.emailaddress.address + ")" } 
        $CcRecipients = $CcRecipients -join ', '

        # Format additional properties for inclusion in the HTML file
        $AdditionalProperties = @"
        <p><strong>Received Date:</strong> $SentDateTime</p>
        <p><strong>From:</strong> $Sender</p>
        <p><strong>To:</strong> $ToRecipients</p>
        <p><strong>CC:</strong> $CcRecipients</p>
        <br>
        <p><strong>Subject:</strong> $EmailSubject</p>
"@

        # Append additional properties to the email content
        $EmailContent = $AdditionalProperties + $EmailContent

        # Write email content to an HTML file
        Set-Content -Path $FilePath -Value $EmailContent
    }
}

####################################### Main Script Execution #######################################

# Import shared mailboxe list from a CSV file
$SharedMailboxes = Import-Csv -Path $SharedMailboxList

# Process each shared mailbox listed in the CSV file
foreach ($Mailbox in $SharedMailboxes) {
    # Retrieve an access token for Microsoft Graph API
    $AccessToken = Connect-GraphApiWithClientSecret -TenantId $TenantId -AppId $AppId -ClientSecret $ClientSecret
    
    Write-Host "Processing mailbox: $($Mailbox.UserPrincipalName)"

    # Retrieve the most recent emails from the mailbox
    $Emails = Get-RecentEmails -AccessToken $AccessToken -MailboxId $Mailbox.UserPrincipalName

    # Export the retrieved emails to the specified directory
    Export-Emails -MailboxUPN $Mailbox.UserPrincipalName -Emails $Emails
}
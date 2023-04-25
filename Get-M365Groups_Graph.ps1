<#
  Copyright 2023 Jake Gwynn
  Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files
  (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge,
  publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
  subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
  FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

# Variables to be populated 
$tenantId = '<your-tenant-id>'
$clientID = '<your-client-id>'
$clientSecret = '<your-client-secret>'
$CsvExportPath = 'c:\temp\AllGroups.csv'

# Function to get an access token using client credentials
function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $body = @{
        'client_id'     = $ClientId
        'client_secret' = $ClientSecret
        'grant_type'    = 'client_credentials'
        'scope'         = 'https://graph.microsoft.com/.default'
    }

    $response = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
    return $response.access_token
}

# Function to list all groups in a tenant using the Microsoft Graph API
function Get-AllGroups {
    param (
        [string]$AccessToken
    )

    $url = "https://graph.microsoft.com/v1.0/groups"
    $groups = @()

    do {
        $response = Invoke-RestMethod -Method GET -Uri $url -Headers @{Authorization = "Bearer $AccessToken"}
        $groups += $response.value
        $url = $response.'@odata.nextLink'
    } while ($url)

    return $groups
}

##### Main Script

$accessToken = Get-AccessToken -TenantId $tenantId -ClientId $AppId -ClientSecret $clientSecret
$groups = Get-AllGroups -AccessToken $accessToken

# Output the results
$groups | Export-CSV -Path $CsvExportPath -NoTypeInformation

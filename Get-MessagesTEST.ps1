$ClientSecret = ""
$TenantId = ""
$AppId = ""

function Connect-GraphApiWithClientSecret {
    Write-Host "Authenticating to Graph API"
    $Body = @{    
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        client_Id     = $AppId
        Client_Secret = $ClientSecret
    } 
    $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $Body
    return $ConnectGraph.access_token
}
$Token = Connect-GraphApiWithClientSecret

$SecureStringToken = ConvertTo-SecureString -String $Token -AsPlainText -Force

Connect-MgGraph -AccessToken $SecureStringToken

Get-MgUserMessage -UserId "jakegwynn@jakegwynndemo.com" -All

function Get-GraphUserMessages {
    param (
        [string]$AccessToken,
        [string]$UserPrincipalName
    )

    $url = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/messages"
    $messages = @()

    <#
    do {
        $response = Invoke-RestMethod -Method GET -Uri $url -Headers @{Authorization = "Bearer $AccessToken"}
        $messages += $response.value
        $url = $response.'@odata.nextLink'
    } while ($url)
    #>
    $response = Invoke-RestMethod -Method GET -Uri $url -Headers @{Authorization = "Bearer $AccessToken"}
    $messages = $response.value

    return $messages
}

Get-GraphUserMessages -AccessToken $Token -UserPrincipalName "jakegwynn@jakegwynndemo.com"

$url = "https://graph.microsoft.com/v1.0/users/jakegwynn@jakegwynndemo.com/messages"
$response = Invoke-RestMethod -Method GET -Uri $url -Headers @{Authorization = "Bearer $Token"}

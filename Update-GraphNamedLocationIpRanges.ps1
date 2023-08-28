[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

$ClientSecret = ""
$TenantId = ""
$AppId = ""

$global:TokenTimer = $null
$global:Token = $null

function Get-RestApiError ($RestError) {
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

function Connect-GraphApiWithClientSecret ($TenantId, $AppId, $ClientSecret) {
    if($global:TokenTimer -eq $null -or $global:TokenTimer.elapsed.minutes -gt '55'){
        try{
            Write-Host "Authenticating to Graph API"
            $Body = @{    
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                client_Id     = $AppId
                Client_Secret = $ClientSecret
                } 
            $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $Body
            $Global:TokenTimer =  [system.diagnostics.stopwatch]::StartNew()	
            $global:GraphToken = $ConnectGraph.access_token
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
        return $global:GraphToken
    }
}

function Get-GraphNamedLocationPolicies {
    $Uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
    try{
        $NamedLocationPolicies = (Invoke-RestMethod -Headers $Headers -Uri $Uri -Method Get).value

        return $NamedLocationPolicies
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}

function Update-NamedLocationPolicy {
    param (
        [Parameter(Mandatory = $true)]
        $NamedLocationPolicyId,
        [Parameter(Mandatory = $true)]
        $UpdatedIpRangeList
    )
    $Uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/$NamedLocationPolicyId"
    try{
        $body = @{
            "@odata.type" = "#microsoft.graph.ipNamedLocation"
            "ipRanges" = $UpdatedIpRangeList
        } | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Headers $headers -Uri $Uri -Method Patch -Body $body -ContentType "application/json"
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}
    

$global:Token = Connect-GraphApiWithClientSecret -TenantId $TenantId -AppId $AppId -ClientSecret $ClientSecret

$headers = @{
    "Authorization" = "Bearer $Token"
    "Content-type" = "application/json"
}


# Get Named Location Policy that matches the DisplayName being targeted
$NamedLocationPolicy = Get-GraphNamedLocationPolicies | Where-Object {$_.displayName -eq "TestLocations"}

# IP Ranges to Add to Named Location Policy
$NewIpRanges = @("6.1.1.1/32", "7.1.2.2/32")

# Object that will be used to update Named Location Policy
[array]$UpdatedIpRangeList = $NamedLocationPolicy.ipRanges

foreach ($IpRange in $NewIpRanges) {
    # Add each IP Range from the $NewIpRanges variable to the $UpdatedIpRangeList variable
    $UpdatedIpRangeList += @{
        "@odata.type" = "#microsoft.graph.iPv4CidrRange"
        "cidrAddress" = $IpRange
    }
}

# Updates the Named Location Policy with the UpdateIpRangeList variable that contains the original and added IP Ranges.
Update-NamedLocationPolicy -NamedLocationPolicyId $NamedLocationPolicy.id -UpdatedIpRangeList $UpdatedIpRangeList
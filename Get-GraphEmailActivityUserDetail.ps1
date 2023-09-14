[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

$ClientSecret = ""
$TenantId = ""
$AppId = ""
$FileName = "C:\Temp\EmailActivityUserDetail_$(Get-Date -format "MM-dd-yyyy").csv"

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

$global:Token = Connect-GraphApiWithClientSecret -TenantId $TenantId -AppId $AppId -ClientSecret $ClientSecret

$headers = @{
    "Authorization" = "Bearer $Token"
    "Content-type" = "application/json"
}

#Length options for report: D7, D30, D90, and D180
$Report = Invoke-WebRequest -uri "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D7')" -Headers $headers -Method Get 
[System.IO.File]::WriteAllBytes($FileName, $Report.content)

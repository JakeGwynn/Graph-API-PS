[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

$ClientSecret = ""
$TenantId = ""
$AppId = ""

$global:Stopwatch = $null
$global:Token = $null
function Connect-GraphApiWithClientSecret {
    if($global:Stopwatch -eq $null -or $global:Stopwatch.elapsed.minutes -gt '55'){
        Write-Host "Authenticating to Graph API"
        $Body = @{    
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $AppId
            Client_Secret = $ClientSecret
            } 
        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $Body
        $Global:Stopwatch =  [system.diagnostics.stopwatch]::StartNew()	
        return $ConnectGraph.access_token
    }
}

$Token = Connect-GraphApiWithClientSecret
Connect-MgGraph -AccessToken $Token # -Scopes "User.Read.All","Calendars.Read"

$CsvExportPath = "C:\temp\SharedMailboxEvents.csv"
$Users = Get-MgUser -All | Where-Object {$_.Mail -ne $null} #Add whatever filtering you need here

$Meetings = [System.Collections.Generic.List[psobject]]@()

Foreach ($User in $Users) {
    Try{
        $Events = Get-MgUserEvent -UserId $User.UserPrincipalName -Filter "isOrganizer eq true" -All -ErrorAction Stop 
        foreach ($Event in $Events) {
            Write-Host "User: $($User.UserPrincipalName)"
            Write-Host "Event Subject: $($Event.Subject)"
            $Attendees = ""
            foreach ($Attendee in $Event.Attendees) {
                $Attendees += $Attendee.EmailAddress.Address
                $Attendees += ";"
            }
            $Attendees
            $Meetings.Add(@{
                User = $User.UserPrincipalName
                EventSubject = $Event.Subject
                Attendees = $Attendees
            })
        }
    }
    Catch {}
}
$Meetings | Export-Csv -Path $CsvExportPath -NoTypeInformation

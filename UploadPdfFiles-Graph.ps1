<#
Copyright 2022 Jake Gwynn
DISCLAIMER:
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>
[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

$TenantId = ""
$AppId = ""
$ClientSecret = ""

$SharePointBaseHostname = "jakegwynndemo.sharepoint.com"
$PdfFileFolderPath = "C:\users\jakegwynn\Downloads"
$CSVFilePath = "C:\users\jakegwynn\Downloads\SharePointLocations.csv"

$global:Stopwatch = $null
$global:Token = $null
function Connect-GraphApiWithClientSecret {
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
function Get-GraphSiteId ($SiteName){
    $ApiUrl = "https://graph.microsoft.com/v1.0/sites/$($SharePointBaseHostname):/sites/$SiteName/"
    try {
        $ApiCallResponse = (Invoke-RestMethod -Headers @{Authorization = "Bearer $Token"} -Uri $ApiUrl -Method Get)
        return $ApiCallResponse.id
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}
function Get-GraphDocumentLibraries ($SiteId) {
    $ApiUrl = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/"
    try {
        $ApiCallResponse = (Invoke-RestMethod -Headers @{Authorization = "Bearer $Token"} -Uri $ApiUrl -Method Get)
        return $ApiCallResponse.value
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}
function Get-GraphDocLib-Folders ($DocumentLibraryId,$FolderName) {
    [System.Collections.Generic.List[string]]$FolderNameArray = $FolderName -split "/"
    if ($FolderNameArray.Count -gt 1) {
        $FolderNameArray.RemoveAt($FolderNameArray.Count - 1)
        $NewFolderName = $FolderNameArray -join "/"
        $ApiUrl = "https://graph.microsoft.com/v1.0/drives/$($DocumentLibraryId)/root:/$($NewFolderName):/children"
    }
    else {
        $ApiUrl = "https://graph.microsoft.com/v1.0/drives/$DocumentLibraryId/root/children"
    }
    try {
        $ApiCallResponse = (Invoke-RestMethod -Headers @{Authorization = "Bearer $Token"} -Uri $ApiUrl -Method Get)
        return $ApiCallResponse.value
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}
function Get-GraphDocLibFolder-FromLink ($SharePointLink) {
    $ShareLinkInBytes = [System.Text.Encoding]::UTF8.GetBytes($SharePointLink)
    $ShareLinkEncoded = [System.Convert]::ToBase64String($ShareLinkInBytes)
    $ApiUrl = "https://graph.microsoft.com/v1.0/shares/" + "u!" + $ShareLinkEncoded + "/driveItem"
    try {
        $ApiCallResponse = (Invoke-RestMethod -Headers @{Authorization = "Bearer $Token"} -Uri $ApiUrl -Method Get)
        return $ApiCallResponse
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}
function Upload-FileToGraphDocLib ($DocumentLibraryId, $FolderId, $FileName) {
    $ApiUrl = "https://graph.microsoft.com/v1.0/drives/$DocumentLibraryId/items/$($FolderId):/$($FileName):/content"
    try {
        $Params = @{
            Headers = @{Authorization = "Bearer $Token"} 
            Uri = "$ApiUrl"
            Method = "Put"
            InFile = "$PdfFileFolderPath\$FileName"
            ContentType = "application/pdf"
        }
        $ApiCallResponse = Invoke-RestMethod @Params
        Write-Host "File `"$FileName`" successfully uploaded to:" -ForegroundColor Yellow
        Write-Host "URL: https://$SharePointBaseHostname/sites/$($File.'Site Name')/$($File.'Document Library')/" -ForegroundColor Yellow
        Write-Host "Folder Name: $($File.'Folder Name')`r`n" -ForegroundColor Yellow
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}
function Upload-LargeFileToGraphDocLib ($DocumentLibraryId, $FolderId, $FileName) {
    $ApiUrl = "https://graph.microsoft.com/v1.0/drives/$DocumentLibraryId/items/$($FolderId):/$($FileName):/createUploadSession"
    try {
        $Params = @{
            Headers = @{Authorization = "Bearer $Token"} 
            Uri = "$ApiUrl"
            Method = "Post"
            ContentType = "application/json"
        }
        $SessionCallResponse = Invoke-RestMethod @Params
        $LastByte = $FileLength - 1
        $Params = @{
            Headers = @{
                Authorization = "Bearer $Token"
                "Content-Length" = $FileLength
                "Content-Range" = "bytes 0-$LastByte/$FileLength"} 
            Uri = $SessionCallResponse.uploadUrl
            Method = "Put"
            InFile = "$PdfFileFolderPath\$FileName"
            ContentType = "application/pdf"
        }
        $UploadCallResponse = Invoke-RestMethod @Params
        Write-Host "File $FileName successfully uploaded to:" -ForegroundColor Yellow
        Write-Host "URL: https://$SharePointBaseHostname/sites/$($File.'Site Name')/$($File.'Document Library')/" -ForegroundColor Yellow
        Write-Host "Folder Name: $($File.'Folder Name')`r`n" -ForegroundColor Yellow
    }
    catch {
        $RestError = $null
        $RestError = Get-RestApiError -RestError $_
        Write-Host $_ -ForegroundColor Red
        return Write-Host $RestError -ForegroundColor Red 
    }
}

############################ MAIN SCRIPT ############################

$CSV = Import-Csv -Path $CSVFilePath 

# Using separated Site Name, Document Library Name, and Folder Name
foreach ($File in $CSV) {
    if($Stopwatch -eq $null -or $Stopwatch.elapsed.minutes -gt '55'){
        $Token = Connect-GraphApiWithClientSecret
    }
    $ChildFolderName = ($File.'Folder Name' -split "/")[-1]
    $SiteId = Get-GraphSiteId -SiteName $File.'Site Name'
    $DocumentLibraryId = (Get-GraphDocumentLibraries -SiteId $SiteId | Where-Object {$_.Name -eq $File.'Document Library'}).Id
    $FolderId = (Get-GraphDocLib-Folders -DocumentLibraryId $DocumentLibraryId -FolderName $File.'Folder Name' | Where-Object {$_.name -eq $ChildFolderName}).Id
    $FileLength = (Get-Item $PdfFileFolderPath\$FileName).Length
    if ($FileLength -le 4194300) {
        Upload-FileToGraphDocLib -DocumentLibraryId $DocumentLibraryId -FolderId $FolderId -FileName $File.'File Name' 
    }
    else {
        Upload-LargeFileToGraphDocLib -DocumentLibraryId $DocumentLibraryId -FolderId $FolderId -FileName $File.'File Name' 
    }
}

# Using SharePoint full folder URL
foreach ($File in $CSV) {
    if($Stopwatch -eq $null -or $Stopwatch.elapsed.minutes -gt '55'){
        $Token = Connect-GraphApiWithClientSecret
    }
    $FileLength = (Get-Item $PdfFileFolderPath\$FileName).Length
    $Folder = Get-GraphDocLibFolder-FromLink -SharePointLink $File.'SharePoint Directory'
    if ($FileLength -le 4194300) {
        Upload-FileToGraphDocLib -DocumentLibraryId $Folder.parentReference.driveId -FolderId $Folder.id -FileName $File.'File Name' 
    }
    else {
        Upload-LargeFileToGraphDocLib -DocumentLibraryId $Folder.parentReference.driveId -FolderId $Folder.id -FileName $File.'File Name' 
    }
}
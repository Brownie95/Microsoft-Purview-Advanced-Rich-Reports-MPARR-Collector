#Check folder path and construct file names
function GetFileName($Date, $Subscription, $OutputPath)
{
    if ($UseCustomParameters)
    {
        Write-Verbose " using custom parameter for filename"
        $JSONfilename = ($Subscription + "_" + $pFilenameCode + ".json")
    }
    else {
        Write-Verbose " using default for filename"
        $JSONfilename = ($Subscription + "_" + $Date + ".json")
       
    }

    Write-Verbose " filename: $jsonfilename"
    return $OutputPath + $JSONfilename
}

function GetAuthToken
{
    $body = @{grant_type="client_credentials";resource=$APIResource;client_id=$AppClientID;client_secret=$ClientSecretValue}
    Write-Host -ForegroundColor Blue -BackgroundColor white "Obtaining authentication token..." -NoNewline
    try{
        $oauth = Invoke-RestMethod -Method Post -Uri "$loginURL/$TenantDomain/oauth2/token?api-version=1.0" -Body $body -ErrorAction Stop
        $script:tokenExpiresOn = ([DateTime]('1970,1,1')).AddSeconds($oauth.expires_on).ToLocalTime()
        $script:OfficeToken = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
        Write-Host -ForegroundColor Green "Authentication token obtained"
    } catch {
        write-host -ForegroundColor Red "FAILED"
        write-host -ForegroundColor Red "Invoke-RestMethod failed."
        Write-host -ForegroundColor Red $error[0]
        exit
    }
}

function buildLog($BaseURI, $Subscription, $tenantGUID, $OfficeToken) {
    try {
        #
        # if using custom value for start/end 
        #
            $strt = $startTime
            $end  = $endTime

        Write-Verbose " Start = $strt"
        Write-Verbose " End   = $end"

        $URIstring = "$BaseURI/content?contentType=$Subscription&startTime=$strt&endTime=$end&PublisherIdentifier=$TenantGUID"
        Write-Host " "
        Write-Verbose " URI    : $uristring"

        $Log = Invoke-WebRequest -Method GET -Headers $OfficeToken `
               -Uri "$BaseURI/content?contentType=$Subscription&startTime=$strt&endTime=$end&PublisherIdentifier=$TenantGUID" `
               -UseBasicParsing -ErrorAction Stop
        
    } 
    catch {
        write-host -ForegroundColor Red "Invoke-WebRequest command has failed"
        Write-host $error[0]
        return
    }

	$TotalContentPages = @()
    #Try to find if there is a NextPage in the returned URI
    if ($Log.Headers.NextPageUri) 
    {
        $NextContentPage = $true
        $NextContentPageURI = $Log.Headers.NextPageUri
		if ($NextContentPageURI -is [array])
		{
			$NextContentPageURI = $Log.Headers.NextPageUri[0]
		}
		$oldURI = $NextContentPageURI

        Write-Verbose " NextPage is present: $NextContentPageURI"

        while ($NextContentPage -ne $false)
        {
			Write-Verbose "Retrieving page nr $($TotalContentPages.Count + 1)"
            $ThisContentPage = Invoke-WebRequest -Headers $OfficeToken -Uri $NextContentPageURI -UseBasicParsing
            $TotalContentPages += $ThisContentPage

            if ($ThisContentPage.Headers.NextPageUri)
            {
                $NextContentPage = $true    
            }
            else
            {
                $NextContentPage = $false
            }
            $NextContentPageURI = $ThisContentPage.Headers.NextPageUri
			if ($NextContentPageURI -is [array])
			{
				$NextContentPageURI = $Log.Headers.NextPageUri[0]
			}
			if ($oldURI -eq $NextContentPageURI)
			{
				$NextContentPage = $false
			}
			$oldURI = $NextContentPageURI
        }
    } 
    $TotalContentPages += $Log

    Write-Host -ForegroundColor Green "OK"
    Write-Host "***"
    return $TotalContentPages
}

function FetchData($TotalContentPages, $Officetoken, $Subscription) {
    # Changed from "-gt 2" to "-gt 0"
    if ($TotalContentPages.content.length -gt 0)
    {
        $uris = @()
        $pages = $TotalContentPages.content.split(",")
        
        foreach($page in $pages)
        {
            if ($page -match "contenturi") {
                $uri = $page.split(":")[2] -replace """"
                $uri = "https:$uri"
                $uris += $uri
            }
        }

        $Logdata = @()
        $filterName = "Filter" + $Subscription.Replace('.', '')
        foreach($uri in $uris)
        {

            Write-Verbose " uri:$uri"

            try {

                # check for token expiration
                if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
                {
                    Write-Host "Refreshing access token..."
                    GetAuthToken
                }

                $result = Invoke-RestMethod -Uri $uri -Headers $Officetoken -Method Get
                if ($script:PSBoundParameters.ContainsKey($filterName))
                {
                    Write-Verbose "Applying filter '$($script:PSBoundParameters[$filterName])' on $($filterName)."
                    if ($schemas.$filterName -eq "NotContains")
                    {
                        $Logdata += $result | Where-Object {$_.Operation -notmatch $($script:PSBoundParameters[$filterName])}
                    }
                    else 
                    {
                        $Logdata += $result | Where-Object {$_.Operation -match $($script:PSBoundParameters[$filterName])}
                    }
                }
                else 
                {
                    $Logdata += $result
                }
            } 
            catch {
                write-host -ForegroundColor Red "ERROR"
                Write-host $error[0]
                return
            }      
        }
        $Logdata 
        write-host -ForegroundColor Green "OK"
    } 
    else {
        Write-Host -ForegroundColor Yellow "Nothing to output"
    }
}

function Export-Logs
{
    Write-Verbose " enter export-logs" 

    # Access token Request and Retrieval 
    GetAuthToken
    
    #create new Subscription (if needed)

    Write-Host -ForegroundColor Blue -BackgroundColor white "Creating Subscriptions...."

    foreach($Subscription in $Subscriptions){
        Write-Host -ForegroundColor Cyan "$Subscription : " -NoNewline
        try { 
            $response = Invoke-WebRequest -Method Post -Headers $OfficeToken `
                                          -Uri "$BaseURI/start?contentType=$Subscription" `
                                          -UseBasicParsing -ErrorAction Stop
        } catch {
            $_.Exception.Message
        }
    }

    #Check subscription status
    $CheckSubTemp = Invoke-WebRequest -Headers $OfficeToken -Uri "$BaseURI/list" -UseBasicParsing
    Write-Host -ForegroundColor Blue -BackgroundColor white "Subscription Content Status"
    $CheckSub = $CheckSubTemp.Content | convertfrom-json
    $CheckSub | ForEach-Object {write-host $_.contenttype "--->" -nonewline; write-host $_.status -ForegroundColor Green}

    #Collecting and Exporting Log data
    Write-Host -ForegroundColor Blue -BackgroundColor white "Collecting and Exporting Log data"
    foreach($Subscription in $Subscriptions)
    {    
        Write-Host -ForegroundColor Cyan "`n-> Collecting log data from '" -NoNewline
        Write-Host -ForegroundColor White -BackgroundColor DarkGray $Subscription -NoNewline
        Write-Host -ForegroundColor Cyan "': " -NoNewline

        # check for token expiration
        if ($tokenExpiresOn.AddMinutes(5) -lt (Get-Date))
        {
            Write-Host "Refreshing access token..."
            GetAuthToken
        }

        $logs = buildLog $BaseURI $Subscription $TenantGUID $OfficeToken
    
        
        $JSONfileName = getFileName $Date $Subscription $outputPath
    
        $output = FetchData $logs $OfficeToken $Subscription
        if ($ExportToFileOnly)
        {
            $output | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 $JSONfilename
            Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
            Write-Host -ForegroundColor White -BackgroundColor DarkGray $JSONfilename -NoNewline
            Write-Host -ForegroundColor Cyan "': " -NoNewline
    
        }
        elseif ($ExportWithFile)
        {
            $output | ConvertTo-Json -Depth 100 | Set-Content -Encoding UTF8 $JSONfilename
            Write-host -ForegroundColor Cyan "---> Exporting log data to '" -NoNewline
            Write-Host -ForegroundColor White -BackgroundColor DarkGray $JSONfilename -NoNewline
            Write-Host -ForegroundColor Cyan "': " -NoNewline
            Publish-LogAnalytics $output $Subscription
        }
        else 
        {
            Publish-LogAnalytics $output $Subscription
        }
    }
}

# Function to create the authorization signature
function BuildSignature 
{
    param(
        $customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource
    )

    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId,$encodedHash
    return $authorization
}

# Function to create and post the request
function PostLogAnalyticsData
{
    param(
        $customerId, $sharedKey, $json, $logType
    )

    $body = [System.Text.Encoding]::UTF8.GetBytes($json)
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $bsParams = @{
        customerId = $customerId
        sharedKey = $sharedKey
        date = $rfc1123date
        contentLength = $contentLength
        method = $method
        contentType = $contentType
        resource = $resource
    }
    $signature = BuildSignature @bsParams
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"

    $headers = @{
        "Authorization" = $signature;
        "Log-Type" = $logType;
        "x-ms-date" = $rfc1123date;
        "time-generated-field" = "CreationTime"
    }

    try
    {
        $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing -ErrorAction Stop
    }
    catch
    {
        $response = New-Object psobject
        $response | Add-Member -MemberType NoteProperty -Name "StatusCode" -Value 400
        $msg = $_.ErrorDetails.Message | ConvertFrom-Json
        $errString = $_.Exception.Message + "`n" + $msg.Error + ": " + $msg.Message
        $response | Add-Member -MemberType NoteProperty -Name "Exception" -Value $errString
    }
    $response

}

function Publish-LogAnalytics
{
    param (
        $objFromJson,
        $Subscription
    )

    Write-Host "Starting export to LA..."
    $list = New-Object System.Collections.ArrayList
    $LogName = $Subscription.Replace(".", "")

    $count = 0
    $elements = 0
    foreach ($item in $objFromJson)
    {
        $elements++
        $count++
        $item | Add-Member -MemberType NoteProperty -Name "EventCreationTime" -Value ($item.CreationTime)
        [void]$list.Add($item)
        if ($elements -ge $BatchSize)
        {
            $elements = 0
            $eventJSON = $list | ConvertTo-Json -Depth 100
            $result = PostLogAnalyticsData -customerId $CustomerID -sharedKey $SharedKey -json $eventJSON -logType $LogName 
            if ([int]$result.StatusCode -ne 200)
            {
                $count -= $BatchSize 
                Write-Host "Error exporting to the Log Analytics. Exception: $($result.Exception)" -ForegroundColor Red
                $errorFile = $OutputPath + "Error_" + $Subscription + "_" + $Date + ".json"
                $eventJSON | Set-Content -Encoding utf8 -Path $errorFile
                Write-Host "Failed records were saved to the $errorFile file. Please investigate them and import with ExportAIPData2LA script."
            }
            $list.Clear()
            $list.TrimToSize()            
        }
    }
    if ($list.Count -gt 0)
    {
        $eventJSON = $list | ConvertTo-Json -Depth 100
        $result = PostLogAnalyticsData -customerId $CustomerID -sharedKey $SharedKey -json $eventJSON -logType $LogName 
        if ([int]$result.StatusCode -ne 200)
        {
            $count -= $elements 
            Write-Host "Error exporting to the Log Analytics. Exception: $($result.Exception)" -ForegroundColor Red
            $errorFile = $OutputPath + "Error_" + $Subscription + "_" + $Date + ".json"
            $eventJSON | Set-Content -Encoding utf8 -Path $errorFile
            Write-Host "Failed records were saved to the $errorFile file. Please investigate them and import with ExportAIPData2LA script."
        }
    }
    Write-Host "$count elements exported for $Subscription."
}

# Script variables 01  --> Update everything in this section:
$containerPath = (Get-AutomationVariable -Name ContainerPath)
$OutputPath = $containerPath
$BatchSize = 500

#API Endpoint URLs ---> Don't Update anything here
$CLOUDVERSIONS = @{
    Commercial = "https://manage.office.com"
    GCC = "https://manage-gcc.office.com"
    GCCH = "https://manage.office365.us"
    DOD = "https://manage.protection.apps.mil"
}

$AppClientID = (Get-AutomationVariable -Name AppClientID)
$ClientSecretValue = (Get-AutomationVariable -Name ClientSecretValue)
$TenantGUID = (Get-AutomationVariable -Name TenantGUID)
$TenantDomain = (Get-AutomationVariable -Name TenantDomain)
$CustomerID = (Get-AutomationVariable -Name CustomerID)
$SharedKey = (Get-AutomationVariable -Name SharedKey)
$Cloud = "Commercial"

$APIResource = $CLOUDVERSIONS.Commercial
if ($Cloud -ne $null)
{
    $APIResource = $CLOUDVERSIONS["$Cloud"]
    Write-Host "Connecting to $Cloud cloud."
}

# Subscriptions
$Subscriptions = @()
$json = (Get-AutomationVariable -Name AuditSources)

[PSCustomObject]$schemas = ConvertFrom-Json -InputObject $json
foreach ($item in $schemas.psobject.Properties)
{
    if ($schemas."$($item.Name)" -eq "True")
    {
        $Subscriptions += $item.Name
    }
}
Write-Host "Subscriptions list: $Subscriptions"   

# Script variables 02  ---> Don't Update anything here:
$loginURL = "https://login.microsoftonline.com/"
$BaseURI = "$APIResource/api/v1.0/$TenantGUID/activity/feed/subscriptions"

$Date = (Get-date).AddDays(-1)
$Date = $Date.ToString('MM-dd-yyyy_hh-mm-ss')

#region Timestamp
$timestampFile = $OutputPath + "timestamp.json"
# read startTime from the file
if (-not (Test-Path -Path $timestampFile))
{
    # if file not present create new value
    $startTime = (Get-Date).AddHours(-23).ToString("yyyy-MM-ddTHH:mm:ss")
}
else 
{
    $json = Get-Content -Raw -Path $timestampFile
    [PSCustomObject]$timestamp = ConvertFrom-Json -InputObject $json
    $startTime = $timestamp.startTime.ToString("yyyy-MM-ddTHH:mm:ss")   
    # check if startTime greater than 7 days (7 days is max value)
    if ((New-TimeSpan -Start $startTime -End ([datetime]::Now)).TotalDays -gt 7)
    {
        $startTime = (Get-Date).AddDays(-7).AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss")
        Write-Host "StartTime is older than 7 days. Setting to the correct value: $startTime" -ForegroundColor Yellow
        Write-Host "Records with CreationTime older than two days will be ingested with current time for the TimeGenerated column!" -ForegroundColor Red
    }
}
$endTime = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
# check if difference between start and end times bigger than 24 hours 
if ((New-TimeSpan -Start $startTime -End $endTime).TotalHours -gt 24)
{
    $endTime = ([datetime]$startTime).AddHours(23).ToString("yyyy-MM-ddTHH:mm:ss")
    Write-Host "Timeframe based on StartTime is bigger than 24 hours. Setting to the correct value: $startTime" -ForegroundColor Yellow
    if ((New-TimeSpan -Start $startTime -End ([datetime]::Now)).TotalDays -gt 2)
    {
        Write-Host "Records with CreationTime older than two days will be ingested with current time for the TimeGenerated column!" -ForegroundColor Red
    }
}
$timestamp = @{"startTime" = $endTime}
ConvertTo-Json -InputObject $timestamp | Out-File -FilePath $timestampFile -Force


Export-Logs
# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format.
#$currentUTCtime = (Get-Date).ToUniversalTime()

$storageAccountName = $env:storageaccount_name
$monitoredQueues = $env:storageaccount_monitoredqueues
$accessKey = $env:storageaccount_accesskey
$appInsightsKey = $env:appinsights_key

# Load .dll assembly into PowerShell session
[Reflection.Assembly]::LoadFile("$PSScriptRoot\Microsoft.ApplicationInsights.dll")

# Instanciate a new TelemetryClient
$TelemetryClient = [Microsoft.ApplicationInsights.TelemetryClient]::new()

# Set the Application Insights Instrumentation Key
$TelemetryClient.InstrumentationKey = $appInsightsKey

ForEach ($queue in $monitoredQueues.split(',')) {

    $apiVersion = "2021-12-02"
    $queue_url = "https://$storageAccountName.queue.core.windows.net/$queue" + "?comp=metadata"
    $apiCallTimestamp = [DateTime]::UtcNow.ToString('r')

    # Generate signature string based on the request we are about to make to the Azure Storage Queue API
    $signatureString_canonicalizedHeaders = "x-ms-date:$apiCallTimestamp`nx-ms-version:$apiVersion`n"
    $signatureString_canonicalizedResource = "/$storageAccountName/$queue" + "`ncomp:metadata"
    $stringToSign = "GET" + "`n`n`n`n`n`n`n`n`n`n`n`n" + `
        $signatureString_canonicalizedHeaders + `
        $signatureString_canonicalizedResource

    # Hash the signature using the Storage Account Key and convert to Base64 to use in the Authorization header
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.key = [Convert]::FromBase64String($accessKey)
    $signatureUTF8 = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($stringToSign))
    $signatureBase64 = [Convert]::ToBase64String($signatureUTF8)

    # Build the headers for the request
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("x-ms-date", $apiCallTimestamp)             
    $headers.Add("x-ms-version", $apiVersion)      
    $headers.Add("Authorization","SharedKey $($storageAccountName):$signatureBase64")

    # Make the request to the Azure Storage Queue API and return the message count

    try {
        Invoke-RestMethod -Method "GET" -Uri $queue_url -Headers $headers -ResponseHeadersVariable 'ResponseHeaders' | Out-Null
        $messageCount = $responseHeaders['x-ms-approximate-messages-count']

        # Write an information log with the current time.
        #Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
        $TelemetryClient.TrackEvent("Checked message count for queue: $queue")
        Write-Host("Checked message count for '$queue' - $messageCount messages")

        $metric = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $metric.add("name", "queue-$queue MessageCount")
        $metric.add("value", $messageCount)
        $TelemetryClient.TrackMetric($metric)
        $TelemetryClient.Flush()
        
    } catch [Microsoft.PowerShell.Commands.HttpResponseException] {
        Write-Host "An error occurred calling the Queue Storage API for metadata for queue: $queue"
        Write-Host "Message: $($_.Exception.Message)" -ErrorAction Stop
    }
}

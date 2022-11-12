
param (
    [Parameter(Mandatory)]
    [string]
    $ApiKey = "",
    [Parameter(Mandatory)]
    [string]
    $AgentName = "",
    [Parameter(Mandatory)]
    [string]
    $DOMOTZ_AGENT_IP = "",
    [Parameter()]
    [string]
    $EndPoint = "https://api-us-east-1-cell-1.domotz.com/public-api/v1" 

)

$HEADERS = @{

    'x-api-key' = $ApiKey
}

$POSTData = @{

    name     = $AgentName
    endpoint = $EndPoint
}

$URI = 'http://'+$DOMOTZ_AGENT_IP+':3000/api/v1/agent'
Write-Host $URI

Invoke-WebRequest -Method Post -Headers $HEADERS -Body $POSTData -Uri $URI

# Define API Key and Base URL
$apiKey = "YOUR-API-KEY"
$baseURL = "https://api-us-east-1-cell-1.domotz.com/public-api/v1/"

# Set API Endpoint
$apiEndpoint = "$baseURL/custom-driver"

# Define headers
$headers = @{
    "Accept" = "application/json"
    "X-Api-Key" = $apiKey
}

# Perform API Request
$response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get

# Check if response contains data
if ($response -and $response.Count -gt 0) {
    # Extract relevant data
    $customDrivers = $response | Select-Object @{Name="Driver Name"; Expression={$_.name}}, @{Name="Driver ID"; Expression={$_.id}}
    
    # Export to CSV
    $csvPath = ".\CustomDrivers.csv"
    $customDrivers | Export-Csv -Path $csvPath -NoTypeInformation

    Write-Host "CSV file created at: $csvPath"
} else {
    Write-Host "No custom drivers found."
}
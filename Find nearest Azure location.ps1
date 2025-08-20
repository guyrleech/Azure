<#
.SYNOPSIS
    Find closest Azure locations

.PARAMETER subscription
    The Azure subscription name to use (only required if not already authenticated)

.PARAMETER tenant
    The Azure tenant to use (only required if not already authenticated)

.NOTES
    Modification History:

    2025/08/19  @guyrleech  Script born
    2025/08/20  @guyrleech  Added authentication
#>

[CmdletBinding()]

Param
(
    [string]$subscription ,
    [string]$tenant ,
    [int]$first = [int]::MaxValue ,
    [string]$geoApiUrl = "http://ip-api.com/json/?fields=lat,lon,city,country"
)

# Function to calculate Haversine distance (in kilometers)
function Get-HaversineDistance {
    param (
        [double]$lat1, [double]$lon1, [double]$lat2, [double]$lon2
    )
    $earthRadius = 6371  # Earth's radius in kilometers
    $dLat = [Math]::PI * ($lat2 - $lat1) / 180
    $dLon = [Math]::PI * ($lon2 - $lon1) / 180
    $a = [Math]::Sin($dLat / 2) * [Math]::Sin($dLat / 2) +
         [Math]::Cos([Math]::PI * $lat1 / 180) * [Math]::Cos([Math]::PI * $lat2 / 180) *
         [Math]::Sin($dLon / 2) * [Math]::Sin($dLon / 2)
    $c = 2 * [Math]::ATan2([Math]::Sqrt($a), [Math]::Sqrt(1 - $a))
    return $earthRadius * $c
}

Write-Verbose "Fetching current location from IP geolocation API"
try {
    $geoResponse = Invoke-RestMethod -Uri $geoApiUrl -Method Get
    $currentLat = $geoResponse.lat
    $currentLon = $geoResponse.lon
    $currentCity = $geoResponse.city
    $currentCountry = $geoResponse.country
    Write-Verbose "Current location: $currentCity, $currentCountry (Lat: $currentLat, Lon: $currentLon)"
} catch {
    Write-Error "Error fetching geolocation: $_"
    exit
}

try {
    Import-Module -Name Az.Accounts,Az.Resources -verbose:$false -ErrorAction Stop
    if( $null -eq (Get-AzContext) )
    {
        Write-Warning "Prompting for interactive Azure sign in"
        [hashtable]$connectParameters = @{}
        if( -Not [string]::IsNullOrEmpty( $subscription ) ) {
            $connectParameters.Add( 'Subscription' , $subscription )
        }
        if( -Not [string]::IsNullOrEmpty( $tenant ) ) {
            $connectParameters.Add( 'Tenant' , $tenant )
        }
        $connection = Connect-AzAccount @connectParameters
        if( -Not $? -or $null -eq $connection ) {
            Write-Error "Azure authentication issue"
        }
    }
    $azLocations = $null
    $azLocations = @( Get-AzLocation | Where-Object { $_.Latitude -and $_.Longitude } )
    if( $null -eq $azLocations -or $azLocations.Count -eq 0 ) {
        Write-Error "Failed to retrieve any Azure locations"
        exit
    }
    Write-Verbose "Retrieved $($azLocations.Count) Azure regions with coordinates"
} catch {
    Write-Error "Error fetching Azure locations: $_"
    exit
}

$closestRegion = $null
$minDistance = [double]::MaxValue

[array]$distances = @( foreach ($location in $azLocations) {
    $distance = Get-HaversineDistance -lat1 $currentLat -lon1 $currentLon -lat2 $location.Latitude -lon2 $location.Longitude
    [PSCustomObject]@{
        DisplayName = $location.DisplayName
        Location    = $location.Location
        DistanceKm  = [math]::Round($distance, 2)
        Latitude    = $location.Latitude
        Longitude   = $location.Longitude
    }
    if ($distance -lt $minDistance) {
        $minDistance = $distance
        $closestRegion = $location
    }
})

# Step 4: Output and log results
Write-Verbose "Closest Azure region: $($closestRegion.DisplayName) ($($closestRegion.Location)), Distance: $([math]::Round($minDistance, 2)) km"

# Display all distances (sorted)
$distances | Sort-Object -Property DistanceKm | Select-Object -First $first

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String]
    $OutFolder,

    [Parameter()]
    [String[]]
    $UserId
)

if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

$startTime = Get-Date

if (!($null = Get-Module Microsoft.Graph.Authentication -ErrorAction Stop)) {
    "Connect to Microsoft Graph First using the Connect-MgGraph cmdlet." | Out-Default
    return $null
}

if (!($null = Get-MgContext)) {
    "Connect to Microsoft Graph First using the Connect-MgGraph cmdlet." | Out-Default
    return $null
}

if (!(Test-Path $OutFolder -PathType Container)) {
    "The specified output folder does not exist or is not valid." | Out-Default
    return $null
}

# Define the export location
$photoLocation = (Resolve-Path $OutFolder).Path
$resultFile = "$($photoLocation)\$((Get-MgDomain | Where-Object {$_.IsInitial}).Id)_$($startTime.ToString("yyyy-MM-dd_HH-mm-ss")).csv"

try {
    $null = New-Item -ItemType File -Path $resultFile -ErrorAction Stop -Force -Confirm:$false
}
catch {
    "Cannot create the result file." | Out-Default
    $_.Exception.Message | Out-Default
    return $null
}

## Set progress bar visibility
$ProgressPreference = 'Continue'

## Set progress bar style if PowerShell Core
if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

$properties = "UserPrincipalName", "DisplayName", "AccountEnabled"
if (!$UserId) {
    $users = Get-MgUser -Filter "userType eq 'Member'" -Property $properties -All | Select-Object $properties | Sort-Object UserPrincipalName
}

if ($UserId) {
    $users = @(
        $UserId | ForEach-Object {
            Get-MgUser -UserId $_ -Property $properties | Select-Object $properties
        }
    )
}

$users | Add-Member -MemberType NoteProperty -Name PhotoFilename -Value ''
$users | Add-Member -MemberType NoteProperty -Name Notes -Value ''

for ($i = 0; $i -lt $($users.Count); $i++) {
    $percentComplete = (($i + 1) * 100) / $($users.Count)
    Write-Progress -Activity "User: $($users[$i].DisplayName)" -Status "Progress: $($i+1) of $($users.Count) ($([math]::round($percentComplete,2))%)" -PercentComplete $percentComplete -ErrorAction SilentlyContinue
    try {
        $photoFileName = "$($users[$i].UserPrincipalName)_photo.jpg"
        Get-MgUserPhotoContent -UserId $users[$i].UserPrincipalName -OutFile "$($photoLocation)\$($photoFileName)" -ErrorAction Stop -WarningAction SilentlyContinue
        $users[$i].PhotoFilename = $photoFileName
    }
    catch {
        if ($_.Exception.Message -like "*Microsoft.Fast.Profile.Core.Exception.ImageNotFoundException*") {
            $users[$i].Notes = "User doesn't have a photo."
        }
        else {
            $users[$i].Notes = $_.Exception.Message
        }
        $users[$i].PhotoFilename = ''
    }
    $users[$i]
    $users[$i] | Export-Csv -Append -Force -Path $resultFile
}

Write-Progress -Activity 'Done' -Completed -PercentComplete 100

$endTime = Get-Date

$timeSpan = New-TimeSpan -Start $startTime -End $endTime
"========================================" | Out-Default
"               SUMMARY" | Out-Default
"========================================" | Out-Default
"Total users      : $($users.Count)" | Out-Default
"Users with photo : $(($users | Where-Object {$_.PhotoFilename}).Count)" | Out-Default
"Photo dump       : $($photoLocation)" | Out-Default
"Total time       : $($timeSpan.ToString("dd\.hh\:mm\:ss"))" | Out-Default
"========================================" | Out-Default


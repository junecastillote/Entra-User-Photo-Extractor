# Define the export location
$photoLocation = "$((Resolve-Path .\).Path)\photos"
New-Item -Path $photoLocation -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

Connect-MgGraph -TenantId "contoso.onmicrosoft.com"

## Set progress bar visibility
$ProgressPreference = 'Continue'

## Set progress bar style if PowerShell Core
if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

$properties = "UserPrincipalName", "DisplayName", "AccountEnabled"
$users = Get-MgUser -Filter "userType eq 'Member'" -Property $properties -All | Select-Object $properties | Sort-Object UserPrincipalName
$users | Add-Member -MemberType NoteProperty -Name PhotoFilename -Value ''
$users | Add-Member -MemberType NoteProperty -Name Notes -Value ''

for ($i = 0; $i -lt $($users.Count); $i++) {
    $percentComplete = $(($($i + 1) * 100) / $($users.Count))
    Write-Progress -Activity "Getting photo of $($users[$i].UserPrincipalName)." -Status "Progress: $($i+1) of $($users.Count)" -PercentComplete $percentComplete
    try {
        $photoFileName = "$($users[$i].UserPrincipalName)_photo.jpg"
        Get-MgUserPhotoContent -UserId $users[$i].UserPrincipalName -OutFile "$($photoLocation)\$($photoFileName)" -ErrorAction Stop
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
}
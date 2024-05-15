Connect-MgGraph -TenantId "tenant.onmicrosoft.com"

## Export photo of all users (if available)
.\Export-EntraIdUserPhoto.ps1 -OutFolder "PHOTO DUMP FOLDER PATH"

## Export photo of specific users
$userId = @("user_a@contoso.com", "user_b@contoso.com", "user_c@contoso.com")
.\Export-EntraIdUserPhoto.ps1 -UserId $userId -OutFolder "PHOTO DUMP FOLDER PATH"
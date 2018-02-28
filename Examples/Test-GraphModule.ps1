Remove-Module MicrosoftGraphAPI
Import-Module "C:\Users\JGS\OneDrive\Git\MicrosoftGraphAPI\MicrosoftGraphAPI"
$clientId = "cdec3c46-b1cd-4ce7-859a-b6fac1ce0b3e"
$redirectUri = "http://www.runbook.guru"

$username = "jakob@runbook.guru"
$password = Read-Host -AsSecureString
$cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username, $password


$Token = $null
$Token = Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId $clientId -RedirectUri $redirectUri -Credential $cred
Invoke-GraphRequest -url "https://graph.microsoft.com/beta/me" -Token $Token -Method Get


<#
.Synopsis
    Get a oAuth token to access Microsoft Graph API
.DESCRIPTION
    Get a oAuth token to access Microsoft Graph API. Token will be valid for 2 hours.
.EXAMPLE
    Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId "cdec3c46-b1cd-4ce7-859a-b6fac1ceafee" -RedirectUri "http://www.runbook.guru" -Credential (get-credential)
.NOTES
This is a preview/beta version. Please send any comments to jgs@coretech.dk
Developed by MVP Jakob Gottlieb Svendsen - jakob@runbook.guru - jgs@coretech.dk
.LINK
    http://graph.microsoft.io/
#>
function Get-GraphAuthToken {
    [CmdletBinding(DefaultParameterSetName = 'Default', 
        SupportsShouldProcess = $false, 
        PositionalBinding = $false,
        HelpUri = 'http://www.runbook.guru/')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = ’Connetion’)]
        [Hashtable] $Connection, 
        #Connection
        #"Name":  "AADTenantName"
        #"Name":  "ClientId"
        #"Name":  "RedirectUri"
        #"Name":  "UserName"
        #"Name":  "Password"
        [Parameter(Mandatory = $true, ParameterSetName = ’Default’)] 
        [String] $AADTenant, 
        [Parameter(Mandatory = $true, ParameterSetName = ’Default’)]
        [String] $ClientId,
        [Parameter(Mandatory = $true, ParameterSetName = ’Default’)]
        [String]$RedirectUri,
        [Parameter(Mandatory = $true, ParameterSetName = ’Default’)]
        [PSCredential] $Credential
    )
   
    #On Connection
    if ($Connection) {   
        $AADTenant = $Connection.AADTenantName
        $ClientId = $Connection.ClientId
        $RedirectUri = $Connection.RedirectUri
    
        $username = $Connection.UserName
        $password = $Connection.Password | ConvertTo-SecureString -AsPlainText -Force
        $Credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
    }
    
    $resourceAppIdURI = “https://graph.microsoft.com”
   
    $authority = “https://login.windows.net/$aadTenant”
   
    <#
   $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
   $uc = new-object Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential -ArgumentList $Credential.Username,$Credential.Password

   $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$uc)
   #>
    try {
        $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority 
        $userCredentials = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential -ArgumentList $Credential.Username, $Credential.Password
        $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext, $resourceAppIdURI, $ClientId, $userCredentials);

        if ($authResult.Result.AccessToken) {
            return  $authResult.Result

        }
        elseif ($authResult.Exception) {
            throw "An error occured getting access token: $($authResult.Exception.InnerException)"
        }
    }
    catch { 
        throw $_.Exception.Message 
    }
   
}

<#
.Synopsis
    Invoke a request to the Microsoft Graph API
.DESCRIPTION
    Invoke a request to the Microsoft Graph API using the Token and setting content type to correct format (JSON)
.EXAMPLE
    $Token = Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId "cdec3c46-b1cd-4ce7-859a-b6fac1ceafee" -RedirectUri "http://www.runbook.guru" -Credential (get-credential)
    Invoke-GraphRequest -url "https://graph.microsoft.com/beta/subscriptions/303d5e85-d6c2-4c2d-9ed3-bd6b2fb5ecf1" -Token $Token -Method DELETE
.NOTES
    This is a preview/beta version. Please send any comments to jgs@coretech.dk
    Developed by MVP Jakob Gottlieb Svendsen - jakob@runbook.guru - jgs@coretech.dk
.LINK
    http://graph.microsoft.io/
#>
Function Invoke-GraphRequest {
    param($Token, $url, $Method, $Body)
    
    try {
        $headers = @{}
        $headers.Add('Authorization', 'Bearer ' + $Token.AccessToken)
        $headers.Add('Content-Type', "application/json")

        if ($Body) {
            $response = Invoke-WebRequest -Uri $url -Method $Method -Body $Body -Headers $headers -UseBasicParsing
        }
        else {
            $response = Invoke-WebRequest -Uri $url -Method $Method -Headers $headers -UseBasicParsing
        }

        return (ConvertFrom-Json $response.Content)
    }
    catch {
        $CurrentError = $error[0]
        #throw ($error[0].Exception.Response) 
        if ($_.Exception.Response) {
            $result = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($result)
            $responseBody = $reader.ReadToEnd();
            throw "Error: $($CurrentError.Exception.Message)`n $($CurrentError.InvocationInfo.PositionMessage) - Reponse:`n $responsebody"
        }
        else {
            throw $_
        }

    }
   
}

<#
.Synopsis
    Gets a subscription object from Microsoft Graph API
.DESCRIPTION
    Gets a subscription object from Microsoft Graph API
.EXAMPLE
    $Token = Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId "cdec3c46-b1cd-4ce7-859a-b6fac1ceafee" -RedirectUri "http://www.runbook.guru" -Credential (get-credential)
    Get-GraphSubscription -Token $Token -SubscriptionId "b539f640-7a5b-462e-960d-e7cb6a3460f6"
.NOTES
    This is a preview/beta version. Please send any comments to jgs@coretech.dk
    Developed by MVP Jakob Gottlieb Svendsen - jakob@runbook.guru - jgs@coretech.dk
.LINK
    http://graph.microsoft.io/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/resources/subscription/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/api/subscription_get/
#>
Function Get-GraphSubscription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] 
        [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationResult] $Token, 
        [Parameter(Mandatory = $true)] 
        [String] $SubscriptionId
    )
  
    $url = "https://graph.microsoft.com/beta/subscriptions"
    $responseBody = Invoke-GraphRequest -url "https://graph.microsoft.com/beta/subscriptions/$SubscriptionId" -Token $Token -Method Get
    return $responseBody

}

<#
.Synopsis
    Creates a new subscription object in the Microsoft Graph API
.DESCRIPTION
    Creates a new subscription object in the Microsoft Graph API
.EXAMPLE
    $Token = Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId "cdec3c46-b1cd-4ce7-859a-b6fac1ceafee" -RedirectUri "http://www.runbook.guru" -Credential (get-credential)
    $webhook = "https://demo.azurewebsites.net/api/webhooks?code=dffdsfdj6pqfrldb6pghzxrihse1zm7vutbj4i&token=VYa2bgSAPCt9NoIx8%2f%2fmG2HrVMvp46vta5Zq6%2bo468Q%3d"
    $resource = "me/mailFolders('Inbox')/messages"
    New-GraphSubscription -Token $Token -ResourceUri $resource -WebhookUri $webhook -ChangeType Created,Deleted,Updated
.NOTES
This is a preview/beta version. Please send any comments to jgs@coretech.dk
Developed by MVP Jakob Gottlieb Svendsen - jakob@runbook.guru - jgs@coretech.dk
.LINK
    http://graph.microsoft.io/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/resources/subscription/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/api/subscription_post_subscriptions/
#>
Function New-GraphSubscription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] 
        [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationResult] $Token, 
        [Parameter(Mandatory = $true)] 
        [String] $ResourceUri, 
        [Parameter(Mandatory = $true)] 
        [String] $WebhookUri, 
        [ValidateSet("Created", "Updated", "Deleted")]
        [String[]] $ChangeType = "Created", 
        [DateTime] $ExpiratetionDateTime = (get-date).ToUniversalTime().AddMinutes(4230), 
        [String] $ClientState = "DefaultClientState"
    )
  
    $url = "https://graph.microsoft.com/beta/subscriptions"

    $FormattedDate = $ExpiratetionDateTime.ToString("yyyy-MM-ddThh:mm:ss.FFFFFFFZ")

    $Request = @"
{
   "changeType": "$($ChangeType -join ",")",
   "notificationUrl": "$WebhookUri",
   "resource": "$ResourceUri",
   "expirationDateTime":"$FormattedDate",
   "clientState": "$ClientState"
}
"@

    $responseBody = Invoke-GraphRequest -Token $Token -url $Url -Method Post -Body $Request

    return $responseBody

}


<#
.Synopsis
    Removes a subscription object from Microsoft Graph API
.DESCRIPTION
    Removes a subscription object from Microsoft Graph API
.EXAMPLE
    $Token = Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId "cdec3c46-b1cd-4ce7-859a-b6fac1ceafee" -RedirectUri "http://www.runbook.guru" -Credential (get-credential)
    Remove-GraphSubscription -Token $Token -SubscriptionId "b539f640-7a5b-462e-960d-e7cb6a3460f6"
.NOTES
    This is a preview/beta version. Please send any comments to jgs@coretech.dk
    Developed by MVP Jakob Gottlieb Svendsen - jakob@runbook.guru - jgs@coretech.dk
.LINK
    http://graph.microsoft.io/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/resources/subscription/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/api/subscription_delete/
#>
Function Remove-GraphSubscription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] 
        [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationResult] $Token, 
        [Parameter(Mandatory = $true)] 
        [String] $SubscriptionId
    )
  
    $url = "https://graph.microsoft.com/beta/subscriptions"
    $responseBody = Invoke-GraphRequest -url "https://graph.microsoft.com/beta/subscriptions/$SubscriptionId" -Token $Token -Method DELETE
    return $responseBody

}

<#
.Synopsis
Updates a subscription object in Microsoft Graph API
.DESCRIPTION
Updates a subscription object in Microsoft Graph API
.EXAMPLE
$Token = Get-GraphAuthToken -AADTenant "runbookguru.onmicrosoft.com" -ClientId "cdec3c46-b1cd-4ce7-859a-b6fac1ceafee" -RedirectUri "http://www.runbook.guru" -Credential (get-credential)
Remove-GraphSubscription -Token $Token -SubscriptionId "b539f640-7a5b-462e-960d-e7cb6a3460f6"
.NOTES
This is a preview/beta version. Please send any comments to jgs@coretech.dk
Developed by MVP Jakob Gottlieb Svendsen - jakob@runbook.guru - jgs@coretech.dk
.LINK
    http://graph.microsoft.io/
.LINK
    http://graph.microsoft.io/en-us/docs/api-reference/beta/resources/subscription/
.LINK    
    http://graph.microsoft.io/en-us/docs/api-reference/beta/api/subscription_update/
#>
Function Update-GraphSubscription {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] 
        [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationResult] $Token, 
        [Parameter(Mandatory = $true)] 
        [String] $SubscriptionId,
        [DateTime] $ExpiratetionDateTime = (get-date).ToUniversalTime().AddMinutes(4230)
    )
  
    $url = "https://graph.microsoft.com/beta/subscriptions"
    $FormattedDate = $ExpiratetionDateTime.ToString("yyyy-MM-ddThh:mm:ss.FFFFFFFZ")

    
    $Request = @"
    {
       "expirationDateTime":"$FormattedDate",
    }
"@

    $responseBody = Invoke-GraphRequest -url "https://graph.microsoft.com/beta/subscriptions/$SubscriptionId" -Token $Token -Method PATCH -Body $Request
    return $responseBody

}

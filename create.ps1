#region Initialize default properties
$config = ConvertFrom-Json $configuration
$p = $person | ConvertFrom-Json
$pp = $previousPerson | ConvertFrom-Json
$pd = $personDifferences | ConvertFrom-Json
$m = $manager | ConvertFrom-Json

$success = $False
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];

# AzureAD Application Parameters #
$config = ConvertFrom-Json $configuration

$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

#endregion Initialize default properties

#region Change mapping here
# Change mapping here
$account = [PSCustomObject]@{
    userPrincipalName = $p.Accounts.MicrosoftActiveDirectory.userPrincipalName
};

#region Execute
try {
    #Find Azure AD ACcount by UserPrincipalName
    Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
    $baseAuthUri = "https://login.microsoftonline.com/"
    $authUri = $baseAuthUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$AADAppId"
        client_secret = "$AADAppSecret"
        resource      = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;

    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept         = "application/json";
    }

    $baseGraphUri = "https://graph.microsoft.com/"
    $searchUri = $baseGraphUri + "v1.0/users/$($account.userPrincipalName)"

    $response = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $azureUser = $response

    if ($azureUser.id -eq $null) { throw "Could not find Azure user $($account.userPrincipalName)" }

    Write-Information "Account correlated to $($azureUser.userPrincipalName)";
    $aRef = $azureUser.id

    $auditLogs.Add([PSCustomObject]@{
            Action  = "CreateAccount"
            Message = "Account correlated to $($azureUser.userPrincipalName)";
            IsError = $false;
        });
	
    $success = $true;
}
catch {
    $auditLogs.Add([PSCustomObject]@{
            Action  = "CreateAccount"
            Message = "Account failed to correlate to $($account.userPrincipalName): $_"
            IsError = $True
        });
    Write-Error $_;
}
#endregion Execute

#region build up result
$result = [PSCustomObject]@{
    Success          = $success;
    AccountReference = $aRef
    AuditLogs        = $auditLogs
    Account          = $account

    # Optionally update the data for use in other systems
    ExportData       = [PSCustomObject]@{
        objectID          = $aRef
        UserPrincipalName = $account.UserPrincipalName
    };
};

Write-Output $result | ConvertTo-Json -Depth 10
#endregion build up result
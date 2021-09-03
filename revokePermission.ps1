#region Initialize default properties
$config = ConvertFrom-Json $configuration
$p = $person | ConvertFrom-Json;
$m = $manager | ConvertFrom-Json;
$aRef = $accountReference | ConvertFrom-Json;
$mRef = $managerAccountReference | ConvertFrom-Json;

# The permissionReference object contains the Identification object provided in the retrieve permissions call
$pRef = $permissionReference | ConvertFrom-Json;

$success = $True
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];

# AzureAD Application Parameters #
$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Troubleshooting
<#
    $aRef = "[AzureObjectGuid]"
    $dryRun = $false
    #>

#Retrieve account information for notifications
$account = [PSCustomObject]@{
    id = $aRef
}

# The permissionReference contains the Identification object provided in the retrieve permissions call

try {
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

    Write-Information "Revoking permission to $($pRef.Name) ($($pRef.id)) for $($aRef)";
    $baseGraphUri = "https://graph.microsoft.com/"
    $removeGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($pRef.id)/members/$($aRef)" + '/$ref'
    if (-Not($dryRun -eq $True)) {
        $response = Invoke-RestMethod -Method DELETE -Uri $removeGroupMembershipUri -Headers $authorization -Verbose:$false
    }
    Write-Information "Successfully revoked Permission to Group $($pRef.Name) ($($pRef.id)) for $($aRef)";
}
catch {
    if ($_ -like "*Resource '$($pRef.id)' does not exist or one of its queried reference-property objects are not present*") {
        Write-Information "AzureAD user $($aRef) is already no longer a member or AzureAD group $($pRef.Name) ($($pRef.id)) does not exist anymore";
    }
    else {
        $success = $False
        # Log error for further analysis.  Contact Tools4ever Support to further troubleshoot
        Write-Error "Error revoking Permission to Group $($pRef.Name) ($($pRef.id)). Error: $_"
    }
}

#build up result
$result = [PSCustomObject]@{ 
    Success          = $success;
    AccountReference = $aRef;
    AuditLogs        = $auditLogs;
    Account          = $account;
};

Write-Output $result | ConvertTo-Json -Depth 10;
# AzureAD Application Parameters #
$config = ConvertFrom-Json $configuration

$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

try {
    Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

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

    Write-Verbose -Verbose "Searching for AzureAD groups.."

    #add optinal popertySelection (mandatory: id,displayName,onPremisesSyncEnabled)
    #comment out $properties to select all properties
    $properties = @("id", "displayName", "onPremisesSyncEnabled")

    if ($null -eq $properties) {
        $select = "&`$select=$($properties -join ",")"
    }
    else {
        $select = $null
    }

    $baseSearchUri = "https://graph.microsoft.com/"
    $searchUri = "$baseSearchUri/v1.0/groups?`$orderby=displayName$select"
    
    $azureADGroupsResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    $azureADGroups = $azureADGroupsResponse.value

    while (![string]::IsNullOrEmpty($azureADGroupsResponse.'@odata.nextLink')) {
        $azureADGroupsResponse = Invoke-RestMethod -Uri $azureADGroupsResponse.'@odata.nextLink' -Method Get -Headers $authorization -Verbose:$false
        $azureADGroups += $azureADGroupsResponse.value
    }    

    Write-Verbose -Verbose "Finished searching for AzureAD Groups. Found [$($azureADGroups.id.Count) groups]"
    
    #Filter for only Cloud groups, since synced groups can only be managed by the Sync
    Write-Verbose -Verbose "Filtering for only Cloud groups.."

    $azureADGroups = foreach ($azureADGroup in $azureADGroups) {
        if ($null -eq $azureADGroup.onPremisesSyncEnabled) {
            $azureADGroup
        }
    }
    Write-Verbose -Verbose "Successfully filtered for only Cloud groups. Filtered down to [$($azureADGroups.id.Count) groups]"
}
catch {
    throw "Could not gather Azure AD groups. Error: $_"
}

$permissions = @(foreach ($azureADGroup in $azureADGroups) {
        @{
            DisplayName    = $azureADGroup.displayName;
            Identification = @{
                Id   = $azureADGroup.id;
                Name = $azureADGroup.displayName;
            }
        }
    })

write-output $permissions | ConvertTo-Json -Depth 10;

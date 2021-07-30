#region Initialize default properties
$config = ConvertFrom-Json $configuration
$p = $person | ConvertFrom-Json
$pp = $previousPerson | ConvertFrom-Json
$pd = $personDifferences | ConvertFrom-Json
$m = $manager | ConvertFrom-Json
$aRef = $accountReference | ConvertFrom-Json
$mRef = $managerAccountReference | ConvertFrom-Json
$pRef = $entitlementContext | ConvertFrom-json

$success = $True
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];
$dynamicPermissions = New-Object Collections.Generic.List[PSCustomObject];

# AzureAD Application Parameters #
$config = ConvertFrom-Json $configuration

$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

$azureAdGroupNamePrefix = "Profession-"
$azureAdGroupNameSuffix = ""

#endregion Initialize default properties

#region Supporting Functions
function Get-ADSanitizeGroupName
{
    param(
        [parameter(Mandatory = $true)][String]$Name
    )
    $newName = $name.trim();
    $newName = $newName -replace ' - ','_'
    $newName = $newName -replace '[`,~,!,#,$,%,^,&,*,(,),+,=,<,>,?,/,'',",;,:,\,|,},{,.]',''
    $newName = $newName -replace '\[','';
    $newName = $newName -replace ']','';
    $newName = $newName -replace ' ','_';
    $newName = $newName -replace '\.\.\.\.\.','.';
    $newName = $newName -replace '\.\.\.\.','.';
    $newName = $newName -replace '\.\.\.','.';
    $newName = $newName -replace '\.\.','.';
    return $newName;
}
#endregion Supporting Functions

#region Change mapping here
$desiredPermissions = @{};
foreach($contract in $p.Contracts) {
    # M365-Functie-<functienaam>
    $group_name = "$azureAdGroupNamePrefix$($contract.Title.Name)$azureAdGroupNameSuffix"  
    $group_name = Get-ADSanitizeGroupName -Name $group_name

    if( ($contract.Context.InConditions) -or (-Not($dryRun -eq $True) ) )
    {        
        $desiredPermissions[$group_name] = $group_name
    }
}
Write-Verbose -Verbose ("Defined Permissions: {0}" -f ($desiredPermissions.keys | ConvertTo-Json))
#endregion Change mapping here

#region Execute
# Operation is a script parameter which contains the action HelloID wants to perform for this permission
# It has one of the following values: "grant", "revoke", "update"
$o = $operation | ConvertFrom-Json

if($dryRun -eq $True) {
    # Operation is empty for preview (dry run) mode, that's why we set it here.
    $o = "grant"
}

Write-Verbose -Verbose ("Existing Permissions: {0}" -f $entitlementContext)
$currentPermissions = @{}
foreach($permission in $pRef.CurrentPermissions) {
    $currentPermissions[$permission.Reference.Id] = $permission.DisplayName
}

# Compare desired with current permissions and grant permissions
foreach($permission in $desiredPermissions.GetEnumerator()) {
    $dynamicPermissions.Add([PSCustomObject]@{
            DisplayName = $permission.Value
            Reference = [PSCustomObject]@{ Id = $permission.Name }
    })

    if(-Not $currentPermissions.ContainsKey($permission.Name))
    {
        # Add user to Membership
        $permissionSuccess = $true
        
        if(-Not($dryRun -eq $True))
        {
            try
            {
                Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
                $baseAuthUri = "https://login.microsoftonline.com/"
                $authUri = $baseAuthUri + "$AADTenantID/oauth2/token"

                $body = @{
                    grant_type      = "client_credentials"
                    client_id       = "$AADAppId"
                    client_secret   = "$AADAppSecret"
                    resource        = "https://graph.microsoft.com"
                }

                $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
                $accessToken = $Response.access_token;

                #Add the authorization header to the request
                $authorization = @{
                    Authorization = "Bearer $accesstoken";
                    'Content-Type' = "application/json";
                    Accept = "application/json";
                }

                Write-Verbose -Verbose "Searching for Group displayName=$($permission.Name)"
                $baseSearchUri = "https://graph.microsoft.com/"
                $searchUri = $baseSearchUri + 'v1.0/groups?$filter=displayName+eq+' + "'$($permission.Name)'"

                $azureADGroupResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
                $azureADGroup = $azureADGroupResponse.value

                if(@($azureADGroup).count -eq 1) {
                    Write-Information "Found Group [$($permission.Name)]. Granting permission for [$($aRef)]";
                    $baseGraphUri = "https://graph.microsoft.com/"
                    $addGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($azureADGroup.id)/members" + '/$ref'
                    $body = @{ "@odata.id"= "https://graph.microsoft.com/v1.0/users/$($aRef)" } | ConvertTo-Json -Depth 10

                    $response = Invoke-RestMethod -Method POST -Uri $addGroupMembershipUri -Body $body -Headers $authorization -Verbose:$false
                
                    Write-Information "Successfully granted Permission for Group [$($permission.Name)] for [$($aRef)]";
                }elseif(@($azureADGroup).count -gt 1){
                    throw "Multiple groups found with displayName=$($permission.Name) "
                }
                else{
                    throw "Group displayName=$($permission.Name) not found"
                }
            }
            catch
            {
                if($_ -like "*One or more added object references already exist for the following modified properties*"){
                    Write-Information "AzureAD user [$($aRef)] is already a member of group";
                }else{
                    $permissionSuccess = $False
                    $success = $False
                    # Log error for further analysis.  Contact Tools4ever Support to further troubleshoot
                    Write-Error ("Error Granting Permission for Group [{0}]:  {1}" -f $permission.Name, $_)
                }
            }
        }

        $auditLogs.Add([PSCustomObject]@{
            Action = "GrantDynamicPermission"
            Message = "Granted membership: {0}" -f $permission.Name
            IsError = -NOT $permissionSuccess
        })
    }
}

# Compare current with desired permissions and revoke permissions
$newCurrentPermissions = @{}
foreach($permission in $currentPermissions.GetEnumerator()) {    
    if(-Not $desiredPermissions.ContainsKey($permission.Name))
    {
        # Revoke Membership
        if(-Not($dryRun -eq $True))
        
        {
            $permissionSuccess = $True
            try
            {
                Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
                $baseAuthUri = "https://login.microsoftonline.com/"
                $authUri = $baseAuthUri + "$AADTenantID/oauth2/token"

                $body = @{
                    grant_type      = "client_credentials"
                    client_id       = "$AADAppId"
                    client_secret   = "$AADAppSecret"
                    resource        = "https://graph.microsoft.com"
                }

                $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
                $accessToken = $Response.access_token;

                #Add the authorization header to the request
                $authorization = @{
                    Authorization = "Bearer $accesstoken";
                    'Content-Type' = "application/json";
                    Accept = "application/json";
                }

                Write-Verbose -Verbose "Searching for Group displayName=$($permission.Name)"
                $baseSearchUri = "https://graph.microsoft.com/"
                $searchUri = $baseSearchUri + 'v1.0/groups?$filter=displayName+eq+' + "'$($permission.Name)'"

                $azureADGroupResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
                $azureADGroup = $azureADGroupResponse.value

                if(@($azureADGroup).count -eq 1) {
                    Write-Information "Found Group [$($permission.Name)]. Revoking permission for [$($aRef)]";
                    $baseGraphUri = "https://graph.microsoft.com/"
                    $removeGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($azureADGroup.id)/members/$($aRef)" + '/$ref'

                    $response = Invoke-RestMethod -Method DELETE -Uri $removeGroupMembershipUri -Headers $authorization -Verbose:$false

                    Write-Information "Successfully revoked Permission for Group [$($permission.Name)] for [$($aRef)]";
                }elseif(@($azureADGroup).count -gt 1){
                    throw "Multiple groups found with displayName=$($permission.Name) "
                }
                else{
                    Write-Warning "Group displayName=$($permission.Name) not found"
                }
            }
            catch
            {
                if($_ -like "*Resource '$($azureADGroup.id)' does not exist or one of its queried reference-property objects are not present*"){
                    Write-Information "AzureAD user [$($aRef)] is already no longer a member or AzureAD group does not exist anymore";
                }else{
                    $permissionSuccess = $False
                    $success = $False
                    # Log error for further analysis.  Contact Tools4ever Support to further troubleshoot.
                    Write-Error ("Error Revoking Permission from Group [{0}]:  {1}" -f $permission.Name, $_)
                }
            }
        }
            
        $auditLogs.Add([PSCustomObject]@{
            Action = "RevokeDynamicPermission"
            Message = "Revoked membership: {0}" -f $permission.Name
            IsError = -Not $permissionSuccess
        })
    } else {
        $newCurrentPermissions[$permission.Name] = $permission.Value
    }
}

# Update current permissions
<# Updates not needed for Group Memberships.
if ($o -eq "update") {
    foreach($permission in $currentPermissions.GetEnumerator()) {    
        if($desiredPermissions.ContainsKey($permission.Name)){
            # Update user to Membership
            $permissionSuccess = $true
            
            # Custom check if Custom attribute 'FunctieMicrosoft365GroupExists' is true
            if($true -ne $contract.Custom.$sourceValidationField){
                Write-Information "Azure AD Group [$($permission.Name)] does not exist. Skipping update action";
            }else{
                if(-Not($dryRun -eq $True))
                {
                    try
                    {
                        Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
                        $baseAuthUri = "https://login.microsoftonline.com/"
                        $authUri = $baseAuthUri + "$AADTenantID/oauth2/token"

                        $body = @{
                            grant_type      = "client_credentials"
                            client_id       = "$AADAppId"
                            client_secret   = "$AADAppSecret"
                            resource        = "https://graph.microsoft.com"
                        }

                        $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
                        $accessToken = $Response.access_token;

                        #Add the authorization header to the request
                        $authorization = @{
                            Authorization = "Bearer $accesstoken";
                            'Content-Type' = "application/json";
                            Accept = "application/json";
                        }

                        Write-Verbose -Verbose "Searching for Group displayName=$($permission.Name)"
                        $baseSearchUri = "https://graph.microsoft.com/"
                        $searchUri = $baseSearchUri + 'v1.0/groups?$filter=displayName+eq+' + "'$($permission.Name)'"

                        $azureADGroupResponse = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
                        $azureADGroup = $azureADGroupResponse.value

                        if(@($azureADGroup).count -eq 1) {
                            Write-Information "Found Group [$($permission.Name)]. Granting permission for [$($aRef)]";
                            $baseGraphUri = "https://graph.microsoft.com/"
                            $addGroupMembershipUri = $baseGraphUri + "v1.0/groups/$($azureADGroup.id)/members" + '/$ref'
                            $body = @{ "@odata.id"= "https://graph.microsoft.com/v1.0/users/$($aRef)" } | ConvertTo-Json -Depth 10

                            $response = Invoke-RestMethod -Method POST -Uri $addGroupMembershipUri -Body $body -Headers $authorization -Verbose:$false
                        
                            Write-Information "Successfully granted Permission for Group [$($permission.Name)] for [$($aRef)]";
                        }elseif(@($azureADGroup).count -gt 1){
                            throw "Multiple groups found with displayName=$($permission.Name) "
                        }
                        else{
                            throw "Group displayName=$($permission.Name) not found"
                        }
                    }
                    catch
                    {
                        if($_ -like "*One or more added object references already exist for the following modified properties*"){
                            Write-Information "AzureAD user [$($aRef)] is already a member of group";
                        }else{
                            $permissionSuccess = $False
                            $success = $False
                            # Log error for further analysis.  Contact Tools4ever Support to further troubleshoot
                            Write-Error ("Error Granting Permission for Group [{0}]:  {1}" -f $permission.Name, $_)
                        }
                    }
                }

                $auditLogs.Add([PSCustomObject]@{
                    Action = "UpdateDynamicPermission"
                    Message = "Updated membership: {0}" -f $permission.Name
                    IsError = -NOT $permissionSuccess
                })
            }
        }

    }
}
#>
#endregion Execute

#region Build up result
$result = [PSCustomObject]@{
    Success = $success;
    DynamicPermissions = $dynamicPermissions;
    AuditLogs = $auditLogs;
};
Write-Output $result | ConvertTo-Json -Depth 10;
#endregion Build up result
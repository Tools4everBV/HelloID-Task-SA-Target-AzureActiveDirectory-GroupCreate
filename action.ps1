# HelloID-Task-SA-Target-AzureActiveDirectory-GroupCreate
#########################################################
# Form mapping
$formObject = @{
    description     = $form.description
    displayName     = $form.displayName
    mailNickname    = $form.mailNickname
    mailEnabled     = $form.mailEnabled
    securityEnabled = $form.securityEnabled
}

try {
    Write-Information "Executing AzureActiveDirectory action: [GroupCreate] for: [$($formObject.DisplayName)]"
    Write-Information "Retrieving Microsoft Graph AccessToken for tenant: [$AADTenantID]"
    $splatTokenParams = @{
        Uri         = "https://login.microsoftonline.com/$($AADTenantID)/oauth2/token"
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = @{                                                                                                                         
            grant_type    = 'client_credentials'
            client_id     = $AADAppID
            client_secret = $AADAppSecret
            resource      = 'https://graph.microsoft.com'
        }
    }
    $accessToken = (Invoke-RestMethod @splatTokenParams).access_token
    Write-Information "Creating AzureAD group [$($formObject.displayName)].."

    $headers = New-Object "[System.Collections.Generic.Dictionary[[String],[String]]]::new()"
    $headers.Add("Authorization", "Bearer $($accessToken)")
    $headers.Add("Content-Type", "application/json")

    $splatCreateGroupParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/groups"
        ContentType = 'application/json'
        Method      = 'POST'
        Headers     = $headers
        Body        = ($formObject | ConvertTo-Json)
    }

    $response = Invoke-RestMethod @splatCreateGroupParams
    
    Write-Information "AzureAD group [$($formObject.displayName)] created successfully"
    $auditLog = @{
        Action            = "CreateResource"
        System            = "AzureActiveDirectory"
        Message           = "AzureAD group [$($formObject.displayName)] created successfully"
        IsError           = $false
        TargetDisplayName = $formObject.displayName
        TargetIdentifier  = $([string]$response.id) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
       
}
catch {
    $ex = $_
    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'AzureActiveDirectory'
        TargetIdentifier  = ''
        TargetDisplayName = $formObject.displayName
        Message           = "Could not execute AzureActiveDirectory action: [GroupCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException')) {
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupCreate] for: [$($formObject.displayName)]"
        Write-Error "Could not execute AzureActiveDirectory action: [GroupCreate] for: [$($formObject.displayName)], error: $($ex.ErrorDetails)"
    }else{
        Write-Information -Tags "Audit" -MessageData $auditLog
        Write-Error "Could not execute AzureActiveDirectory action: [GroupCreate] for: [$($formObject.DisplayName)], error: $($ex.Exception.Message)"
    }    
}
#########################################################

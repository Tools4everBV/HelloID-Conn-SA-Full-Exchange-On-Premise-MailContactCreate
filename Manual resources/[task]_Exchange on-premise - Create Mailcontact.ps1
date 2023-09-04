# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

# Variables configured in form
$Alias = $form.alias
$ExternalEmailAddress = $form.externalEmailAddress
$initials = $form.initials
$FirstName = $form.firstname
$LastName = $form.lastname
$Name = $form.displayname
$HiddenFromAddressListsBoolean = $form.hidefromaddresslist

function Resolve-HTTPError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline
        )]
        [object]$ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            FullyQualifiedErrorId = $ErrorObject.FullyQualifiedErrorId
            MyCommand             = $ErrorObject.InvocationInfo.MyCommand
            RequestUri            = $ErrorObject.TargetObject.RequestUri
            ScriptStackTrace      = $ErrorObject.ScriptStackTrace
            ErrorMessage          = ''
        }
        if ($ErrorObject.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
            $httpErrorObj.ErrorMessage = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            $httpErrorObj.ErrorMessage = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
        }
        Write-Output $httpErrorObj
    }
}

function Remove-EmptyValuesFromHashtable {
    param(
        [parameter(Mandatory = $true)][Hashtable]$Hashtable
    )

    $newHashtable = @{}
    foreach ($Key in $Hashtable.Keys) {
        if (-not[String]::IsNullOrEmpty($Hashtable.$Key)) {
            $null = $newHashtable.Add($Key, $Hashtable.$Key)
        }
    }
    
    return $newHashtable
}

<#----- Exchange On-Premises: Start -----#>
# Connect to Exchange
try {
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
    
    $Log = @{
        Action            = "CreateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    Write-Error "Error connecting to Exchange using the URI [$exchangeConnectionUri]. Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "CreateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to connect to Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

try {
    # Check if Mail Contact already exists (should only occur on a retry of task)
    try {
        Write-Verbose "Querying Exchange Online mail contact with ExternalEmailAddress '$ExternalEmailAddress' OR Alias '$Alias' OR Name '$Name'"
        $mailContact = Get-Recipient -ResultSize unlimited | Where-Object { $_.EmailAddresses -match "$ExternalEmailAddress" }
        Write-Information "Successfully queried Exchange Online mailboxes with email address '$ExternalEmailAddress' of type $($mailboxes.RecipientTypeDetails). Result count: $($mailboxes.Identity.Count)"
    }
    catch {
        $ex = $PSItem
        if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
            $errorObject = Resolve-HTTPError -Error $ex    
            $verboseErrorMessage = $errorObject.ErrorMessage    
            $auditErrorMessage = $errorObject.ErrorMessage
        }
    
        # If error message empty, fall back on $ex.Exception.Message
        if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
            $verboseErrorMessage = $ex.Exception.Message
        }
        if ([String]::IsNullOrEmpty($auditErrorMessage)) {
            $auditErrorMessage = $ex.Exception.Message
        }

        Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"
        throw "Error creating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage"

        # Clean up error variables
        Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
        Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
    }
    if ($null -ne $mailContact.Identity) {
        Write-Warning "Found existing mail contact with ExternalEmailAddress '$ExternalEmailAddress' OR Alias '$Alias' OR Name '$Name'."
    }
    else {
        # Create Mail Contact
        try {
            Write-Verbose "Creating mail contact '$($Name)' with ExternalEmailAddress '$($ExternalEmailAddress)'"

            $mailContactParams = @{
                Name                 = $Name
                FirstName            = $FirstName
                Initials             = $Initials
                LastName             = $LastName
                Alias                = $Alias        
                ExternalEmailAddress = $ExternalEmailAddress
                OrganizationalUnit   = $ADContactsOU
            }
            $mailContactParams = Remove-EmptyValuesFromHashtable $mailContactParams
            $mailContact = New-MailContact @mailContactParams -ErrorAction Stop

            Write-Information "Successfully created mail contact with the following parameters: $($mailContactParams|ConvertTo-Json)"
            $Log = @{
                Action            = "CreateAccount" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Successfully created mail contact with the following parameters: $($mailContactParams|ConvertTo-Json)" # required (free format text) 
                IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $([string]$mailContact.DisplayName) # optional (free format text) 
                TargetIdentifier  = $([string]$mailContact.Guid) # optional (free format text) 
            }
            #send result back
            Write-Information -Tags "Audit" -MessageData $log

            
    
            if ($HiddenFromAddressListsBoolean -eq 'true') {
                $mailContactUpdateParams = @{            
                    Identity                      = $mailContact.Identity                
                    HiddenFromAddressListsEnabled = $true
                }   
                $null = Set-MailContact @mailContactUpdateParams -ErrorAction Stop        
            }
        
        }
        catch {
            $ex = $PSItem
            if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
                $errorObject = Resolve-HTTPError -Error $ex
                $verboseErrorMessage = $errorObject.ErrorMessage
                $auditErrorMessage = $errorObject.ErrorMessage
            }
        
            # If error message empty, fall back on $ex.Exception.Message
            if ([String]::IsNullOrEmpty($verboseErrorMessage)) {
                $verboseErrorMessage = $ex.Exception.Message
            }
            if ([String]::IsNullOrEmpty($auditErrorMessage)) {
                $auditErrorMessage = $ex.Exception.Message
            }
    
            $Log = @{
                Action            = "CreateAccount" # optional. ENUM (undefined = default) 
                System            = "Exchange On-Premise" # optional (free format text) 
                Message           = "Error creating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage" # required (free format text) 
                IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
                TargetDisplayName = $([string]$mailContactParams.Name) # optional (free format text) 
                TargetIdentifier  = $([string]$mailContactParams.ExternalEmailAddress) # optional (free format text) 
            }
            #send result back  
            Write-Information -Tags "Audit" -MessageData $log
    
            Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"
            throw "Error creating mail contact with the following parameters: $($mailContactParams|ConvertTo-Json). Error Message: $auditErrorMessage"
    
            # Clean up error variables
            Remove-Variable 'verboseErrorMessage' -ErrorAction SilentlyContinue
            Remove-Variable 'auditErrorMessage' -ErrorAction SilentlyContinue
        }
    }
}
catch {
    Write-Error "Error creating mail contact.  Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "CreateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Error creating mail contact." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $([string]$mailContactParams.Name) # optional (free format text) 
        TargetIdentifier  = $([string]$mailContactParams.ExternalEmailAddress) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}

# Disconnect from Exchange
try {
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
    $Log = @{
        Action            = "CreateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]" # required (free format text) 
        IsError           = $false # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"
    $Log = @{
        Action            = "CreateAccount" # optional. ENUM (undefined = default) 
        System            = "Exchange On-Premise" # optional (free format text) 
        Message           = "Failed to disconnect from Exchange using the URI [$exchangeConnectionUri]." # required (free format text) 
        IsError           = $true # optional. Elastic reporting purposes only. (default = $false. $true = Executed action returned an error) 
        TargetDisplayName = $exchangeConnectionUri # optional (free format text) 
        TargetIdentifier  = $([string]$session.GUID) # optional (free format text) 
    }
    #send result back  
    Write-Information -Tags "Audit" -MessageData $log
}
<#----- Exchange On-Premises: End -----#>

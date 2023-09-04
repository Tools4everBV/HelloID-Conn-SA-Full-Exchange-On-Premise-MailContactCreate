try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $searchValue = ($dataSource.externalemailaddress).trim()
    $searchQuery = "*$searchValue*"  

    $sessionOptionParams = @{
        SkipCACheck = $true
        SkipCNCheck = $true
        SkipRevocationCheck = $true
    }

    $sessionOption = New-PSSessionOption  @SessionOptionParams 

    $sessionParams = @{        
        Authentication = 'Kerberos' 
        ConfigurationName = 'Microsoft.Exchange' 
        ConnectionUri = $ExchangeConnectionUri 
        Credential = $adminCredential        
        SessionOption = $sessionOption       
    }

    $exchangeSession = New-PSSession @SessionParams

    Write-Information "Search query is '$searchQuery'" 
    
    $getMailcontactsParams = @{
        ResultSize = "Unlimited"
    }
   
    
     $invokecommandParams = @{
        Session = $exchangeSession
        Scriptblock = [scriptblock] { Param ($Params)Get-Recipient @Params}
        ArgumentList = $getMailcontactsParams
    } 

    Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"  
    
    $mailBoxes =  Invoke-Command @invokeCommandParams | Where-Object { $_.EmailAddresses -match "$searchValue" } 

    $resultCount = $mailboxes.Identity.Count

        Write-Information "Successfully queried Exchange for email address '$searchValue'. Result count: $resultCount"


        # Return info message if mailaddress is available or already in use
        if($resultCount -eq 0){
            $result = $true
        }else{
            #$result = "Email address $emailAddress is already in use by another mailbox: $($mailboxes.DisplayName) ($($mailboxes.PrimarySmtpAddress))"
            $result = $false
        }

        $returnObject = @{Result=$result}
        Write-Output $returnObject
    
    Remove-PSSession($exchangeSession)
  
} catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}


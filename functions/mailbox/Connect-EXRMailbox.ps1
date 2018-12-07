function Connect-EXRMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $ClientId,
		
        [Parameter(Position = 2, Mandatory = $false)]
        [string]
        $redirectUrl,
		
        [Parameter(Position = 3, Mandatory = $false)]
        [string]
        $ClientSecret,
		
        [Parameter(Position = 4, Mandatory = $false)]
        [string]
        $ResourceURL,
		
        [Parameter(Position = 5, Mandatory = $false)]
        [switch]
        $Beta,
		
        [Parameter(Position = 6, Mandatory = $false)]
        [String]
        $Prompt,

        [Parameter(Position = 7, Mandatory = $false)]
        [switch]
        $CacheCredentials,

        [Parameter(Position = 8, Mandatory = $false)]
        [switch]
        $Outlook,

        [Parameter(Position = 9, Mandatory = $false)]
        [switch]
        $ShowMenu,

        [Parameter(Position = 10, Mandatory = $false)]
        [switch]
        $EnableTracing,

        [Parameter(Position = 11, Mandatory = $false)]
        [switch]
        $ManagementAPI,
		
        [Parameter(Position = 10, Mandatory = $false)]
        [pscredential]
        $Credential,

        [Parameter(Position = 11, Mandatory = $false)]
        [psobject]
        $AdalToken,

        [Parameter(Position = 12, Mandatory = $false)]
        [string]
        $certificateFileName,

        [Parameter(Position = 13, Mandatory = $false)]
        [SecureString]
        $certificateFilePassword,

        [Parameter(Position = 14, Mandatory = $false)]
        [string]
        $GuestDomain,

        [Parameter(Position = 15, Mandatory = $false)]
        [switch]
        $useADAL,

        [Parameter(Position = 16, Mandatory = $false)]
        [switch]
        $useLoggedOnCredentials,

        [Parameter(Position = 17, Mandatory = $false)]
        [String]
        $AADUserName

		
    )
    Begin {
        if($useADAL.IsPresent){
            $adal = Join-Path $script:ModuleRoot "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $script:ModuleRoot "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
            if([System.IO.File]::Exists($adal)){ 
                Import-Module $adal -Force
            }
            if([System.IO.File]::Exists($adalforms)){ 
               Import-Module $adalforms -Force
            }           
            
        }
        if (![String]::IsNullOrEmpty($GuestDomain)) {
            $TenantId = Get-EXRTenantId -DomainName $GuestDomain
        }
        if ($ManagementAPI.IsPresent) {
            if ([String]::IsNullOrEmpty($ResourceURL)) {
                $ResourceURL = "manage.office.com"
            }
        }
        if (!$ResourceURL) {
            $ResourceURL = "Graph.Microsoft.com"
        }
        if ($AdalToken) {
            $Resource = "graph.microsoft.com"			 
            if ([bool]($AdalToken.PSobject.Properties.name -match "AccessToken")) {
                #$AdalToken.access_token = 
                Add-Member -InputObject $AdalToken -NotePropertyName access_token -NotePropertyValue (Get-ProtectedToken -PlainToken $AdalToken.AccessToken) -Force
            }
            Add-Member -InputObject $AdalToken -NotePropertyName mailbox -NotePropertyValue $MailboxName -Force
            if ($Beta.IsPresent) {
                Add-Member -InputObject $AdalToken -NotePropertyName Beta -NotePropertyValue $True
            }
            if (!$Script:TokenCache.ContainsKey($Resource)) {	
                $ResourceTokens = @{}		
                $Script:TokenCache.Add($Resource, $ResourceTokens)
            }
            Add-Member -InputObject $AdalToken -NotePropertyName Cached -NotePropertyValue $true -Force			
            Add-Member -InputObject $AdalToken -NotePropertyName expires_on -NotePropertyValue (New-TimeSpan -Start (Get-Date "01/01/1970") -End $AdalToken.ExpiresOn.DateTime).TotalSeconds -Force	
            Add-Member -InputObject $AdalToken -NotePropertyName resource -NotePropertyValue ("https://" + $Resource) -Force							
            $HostDomain = (New-Object system.net.Mail.MailAddress($MailboxName)).Host.ToLower()
            if (!$Script:TokenCache[$Resource].ContainsKey($HostDomain)) {			
                $Script:TokenCache[$Resource].Add($HostDomain, $AdalToken)
            }
            else {
                $Script:TokenCache[$Resource][$HostDomain] = $AdalToken
            }
            write-host ("Cached Token for " + $Resource + " " + $HostDomain)
        }
        else {
            if ($certificateFileName) {
                $Resource = "graph.microsoft.com"
                $TenantId = Get-EXRTenantId -DomainName $MailboxName.Split('@')[1]
                if (!$certificateFilePassword) {
                    $certificateFilePassword = Read-Host -AsSecureString -Prompt "Enter password for certificate file"
                }
                $Token = Get-EXRAppOnlyToken -CertFileName $certificateFileName -TenantId $TenantId -ClientId $ClientId  -ResourceURL $Resource -MailboxName $MailboxName -password $certificateFilePassword
                if (!$Token.access_token) {
                    throw "Error getting Access Token"
                }
                else {

                }
            }
            else {
                if ([String]::IsNullOrEmpty($ClientId)) {
                    $redirectUrl = "urn:ietf:wg:oauth:2.0:oob"
                    $defaultAppReg = Get-EXRDefaultAppRegistration
                    if($PSVersionTable.PSEdition -eq "Core"){
                        if(!$Credential){
                            $Credential = Get-Credential
                        }
                    }
                    if($GuestDomain){
			            $defaultAppReg = "" | Select ClientId,RedirectUrl
                        $defaultAppReg.ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
                        $defaultAppReg.RedirectUrl = "urn:ietf:wg:oauth:2.0:oob"
                    }
                    if ($defaultAppReg -eq $null -bor $ShowMenu.IsPresent) {
                        $ProceedOkay = $false
                        Do {
                            Write-Host "
                            ---Default ClientId Selection ----------
                            1 = Mailbox Access Only
                            2 = Mailbox Contacts Access Only
                            3 = Full Access to all Graph API functions
                            4 = Reporting Access Only
                            5 = Management API Access Only
                            6 = Default Office AppId
                            7 = Set Default Application Registration
                            8 = Delete Default Application Registration
                            9 = Exit
                            Or just hit enter to use the default Office AppId
                            --------------------------"
                            $choice1 = read-host -prompt "Select number & press enter"
                            switch ($choice1) {
                                "1" {
                                    $ProceedOkay = $true
                                    $ClientId = "1d236c67-7e0b-42bc-88fd-d0b70a3df50a"
                                }
                                "2" {
                                    $ProceedOkay = $true
                                    $ClientId = "9149e700-47a9-4ba6-b01e-20716509fac7"
							
                                }
                                "3" {
                                    $ProceedOkay = $true
                                    $ClientId = "5471030d-f311-4c5d-91ef-74ca885463a7"
                                }
                                "4" {
                                    $ProceedOkay = $true
                                    $ClientId = "e9a8cb7e-9630-4313-8705-9d6f3181bf01"
                                }
                                "5" {
                                    $ProceedOkay = $true
                                    $ClientId = "2eba6dfc-2962-4242-acdc-acd6c4f5dea8"
                                }	
                                "6" {
                                    $ProceedOkay = $true
                                    $ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
                                }						
                                "7" {
                                    New-EXRDefaultAppRegistration
                                    $ProceedOkay = $true
                                    $defaultAppReg = Get-EXRDefaultAppRegistration
                                    $ClientId = $defaultAppReg.ClientId
                                    $redirectUrl = $defaultAppReg.RedirectUrl 
                                }
                                "8" {
                                    Remove-EXRDefaultAppRegistration
                                    Write-Host "Removed Default Registration"
                                    $ProceedOkay = $true
                                }
                                "9" {return}
                                default {
                                    $ProceedOkay = $true
                                    $ClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
                                }
							

                            }
                        } until ($ProceedOkay)
                    }
                    else {
                        $ClientId = $defaultAppReg.ClientId
                        $redirectUrl = $defaultAppReg.RedirectUrl 
                    }
                    if ([String]::IsNullOrEmpty($ResourceURL)) {
                        $Resource = "graph.microsoft.com"
                    }
                    else {
                        $Resource = $ResourceURL
                    }			
                    if ($Outlook.IsPresent) {
                        $Resource = ""
                    }
                    if ($EnableTracing.IsPresent) {
                        $Script:TraceRequest = $true
                    }                    
                    if ($beta.IsPresent) {
                        $tkn = Get-EXRAccessToken -MailboxName $MailboxName -ClientId $ClientId  -redirectUrl $redirectUrl   -ResourceURL $Resource -beta -Prompt $Prompt -CacheCredentials -TenantId $TenantId                
                    }
                    else {
                        if ($Credential) {
                            $tkn = Get-EXRAccessTokenUserAndPass -ClientId $ClientId -MailboxName $MailboxName -ResourceURL $ResourceURL -CacheCredentials -Credentials $Credential  -TenantId $TenantId
                        }
                        else {
                            if($useADAL.IsPresent){
                                $tkn = Get-EXRAccessTokenADAL -MailboxName $MailboxName -ClientId $ClientId -redirectUrl $redirectUrl  -ResourceURL $ResourceURL -Prompt $Prompt -CacheCredentials -TenantId $TenantId -useLoggedOnCredentials:$useLoggedOnCredentials.IsPresent -AADUserName $AADUserName
                            }else{
                                $tkn = Get-EXRAccessToken -MailboxName $MailboxName -ClientId $ClientId -redirectUrl $redirectUrl  -ResourceURL $Resource -Prompt $Prompt -CacheCredentials -TenantId $TenantId

                            }
                        }
				  
                    }
                }
                else {
                    if ($Credential) {
                        $tkn = Get-EXRAccessTokenUserAndPass -ClientId $ClientId -MailboxName $MailboxName  -ResourceURL $ResourceURL -CacheCredentials -Credentials $Credential  -TenantId $TenantId
                    }
                    else {
                        if($useADAL.IsPresent){
                            $tkn = Get-EXRAccessTokenADAL -MailboxName $MailboxName -ClientId $ClientId -redirectUrl $redirectUrl  -ResourceURL $ResourceURL -Prompt $Prompt -CacheCredentials -TenantId $TenantId -useLoggedOnCredentials:$useLoggedOnCredentials.IsPresent -AADUserName $AADUserName
                        }else{
                            $tkn = Get-EXRAccessToken -ClientId $ClientId -MailboxName $MailboxName -redirectUrl $redirectUrl -ClientSecret $ClientSecret -ResourceURL $ResourceURL -Beta:$beta.IsPresent -prompt $Prompt -CacheCredentials -TenantId $TenantId
                        }
                    }
			
                }
            }
        }
        if ($tkn.Mailbox -ne $null) {
            if ([bool]($tkn.PSobject.Properties.name -match "error")) {   
                throw ("Error connecting to Mailbox " + $tkn) 
                
            }
            else {
                write-host "connected to mailbox"
            }
            
        }
    }
}

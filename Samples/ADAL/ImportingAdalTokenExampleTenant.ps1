$DomainName = "datarumble.com"
Import-Module .\Microsoft.IdentityModel.Clients.ActiveDirectory.dll -Force
$PromptBehavior = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters -ArgumentList Never
$EndpointUri = 'https://login.windows.net/' + (Get-EXRTenantId -DomainName $DomainName)
$Context = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext($EndpointUri)
$token = ($Context.AcquireTokenAsync("https://graph.microsoft.com","d3590ed6-52b3-4102-aeff-aad2292ab01c","urn:ietf:wg:oauth:2.0:oob",$PromptBehavior)).Result
Connect-EXRMailbox -MailboxName gscales@datarumble.com -AdalToken $token
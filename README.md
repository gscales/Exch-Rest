# Exch-Rest
The purpose of this library is to help automate using the REST API for Exchange on Office365
and Exchange 2016. 

Authentication 

To use any of the library cmdlets you need to first Authenticate to create an Access Token using Oauth.
For oAuth Authenticate you need to create an application registration in Azure for the script to use

https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-integrating-applications

There are two options for authentication which is authenticating as a user and athenticating as an application.

To Authenticate as an user use (with the ClientId and redirectUrl you created)

$Token = Get-AccessToken -MailboxName mailbox@domain.com -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 -redirectUrl urn:ietf:wg:oauth:2.0:oob

or if you have a ClientSecret use

$Token = Get-AccessToken -MailboxName mailbox@domain.com -ClientId 1bdbfb41-f690-4f93-b0bb-002004bbca79 -redirectUrl 'http://localhost:8000/authorize' -ClientSecret 1rwq9MmrSMu4SGhMEfGb9ggktWjzPYtW5lcAxXLzEtU=

To Authenticate as an Application see http://gsexdev.blogspot.com.au/2017/03/using-office365exchange-2016-rest-api.html

To create an App Only Token using a certificate file created above use

$Token = Get-AppOnlyToken -CertFile "c:\temp\drCert.pfx" -ClientId 1bdbfb41-f690-4f93-b0bb-002004bbca79 -redirectUrl 'http://localhost:8000/authorize' -TenantId cbdbfb41-f690-4f93-b0bb-002004bbca79 

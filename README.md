# Exch-Rest
A PowerShell module for the Office 365 and Exchange 2016 REST API.

## Setup
#### Application registration
The Office 365 / Exchange 2016 REST API uses OAuth 2.0 to authenticate users. This is good because it means that people using your app do not need to give your their password. Instead, they authenticate against a central authentication system (e.g. Azure AD, Active Directory) and get back a token. They then give your application permission to use that token to do a very limited number of things for a limited period of time.

However, this means that you must do a little bit of work to register an application in Azure before you can use the Exch-Rest functions. A good walk through of the application registration process is provided by Jason Johnston at [https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md](https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md).

The following is a brief overview of the steps you can take to create a native application:
  * Browse to [http://dev.office.com/app-registration](http://dev.office.com/app-registration) and login into your Azure tenant
  * Click `+ New application registration` and fill out the options, and click `Create`
    * Name: `<Name-of-app-users-will-see>`
    * Application type: `Native`
    * Sign-on URL: `http://localhost`
  * Click your newly created native application and note the `Application ID`. You will need this later as your `Client ID`.
  * Click `Redirect URIs`, you should see `http://localhost`. Replace that entry with `urn:ietf:wg:oauth:2.0:oob`
  * Click `Required permissions` and then click `+ Add`
  * Click `1 Select an API`, click `Office 365 Exchange Online (Microsoft.Exchange)`, and then click `Select`
  * Check off all the permissions that you wish to use, and then click `Select`. (Note: there seems to be a bug with the CheckAll button so I had to individually check each permission)
  * Click `Done`

#### Module installation
This module has not yet been published to the PowerShell gallery. The following steps can be used to install the module:
  * Download a zip of the source code `Invoke-WebRequest -Uri "https://codeload.github.com/gscales/Exch-Rest/zip/master" -OutFile "~\Exch-Rest-master.zip"`
  * Unblock the downloaded file `Unblock-File "~\Exch-Rest-master.zip"`
  * Extract the zip, `Expand-Archive "~\Exch-Rest-master.zip" -DestinationPath "~\Documents\WindowsPowerShell\Modules"`
  * Remove "-master" from the name `Move-Item "~\Documents\WindowsPowerShell\Modules\Exch-Rest-master" "~\Documents\WindowsPowerShell\Modules\Exch-Rest"`
  * Delete the downloaded source code `Remove-Item "~\Exch-Rest-master.zip"`
  * Import the module `Import-Module -Name Exch-Rest`


## Authentication
You can either authenticate as a user or as an application.

#### Example 1: authenticating as a user (supplying the ClientId and redirectUrl you created during application registration)
```
$Token = Get-AccessToken -MailboxName mailbox@domain.com `
                         -ClientId 5471030d-f311-4c5d-91ef-74ca885463a7 `
                         -redirectUrl urn:ietf:wg:oauth:2.0:oob
```
#### Example 2: authenticating as a user can and supplying a ClientSecret
```
$Token = Get-AccessToken -MailboxName mailbox@domain.com `
                         -ClientId 1bdbfb41-f690-4f93-b0bb-002004bbca79 `
                         -redirectUrl 'http://localhost:8000/authorize' `
                         -ClientSecret 1rwq9MmrSMu4SGhMEfGb9ggktWjzPYtW5lcAxXLzEtU=
```
#### Example 3: authenticating as an application using a certificate
```
$Token = Get-AppOnlyToken -CertFile "c:\temp\drCert.pfx" `
                          -ClientId 1bdbfb41-f690-4f93-b0bb-002004bbca79 `
                          -redirectUrl 'http://localhost:8000/authorize' `
                          -TenantId cbdbfb41-f690-4f93-b0bb-002004bbca79
```
Note that example 3 is typically used for administrative purposes to manage mulitple mailboxes. This type of authentication requires a different setup steps during the application registration. Please see [http://gsexdev.blogspot.com.au/2017/03/using-office365exchange-2016-rest-api.html](http://gsexdev.blogspot.com.au/2017/03/using-office365exchange-2016-rest-api.html) for more information.

## Usage
After you have authenticated and received a token you can use that token with the Exch-Rest functions to access the Office 365/Exchange REST API.
#### Example 1: get information about mailbox's Inbox use
```
Get-Inbox -MailboxName mailbox@domain.com -AccessToken $Token
```

## Available functions
  * Get-AllMailFolders
  * Get-AllChildFolders
  * Get-AllCalendarFolders
  * Get-AllContactFolders
  * Get-AllTaskfolders
  * Get-AccessToken
  * Get-AppOnlyToken
  * Get-MailboxSettings
  * Get-AutomaticRepliesSettings
  * Get-MailboxTimeZone
  * Get-FolderFromPath
  * Get-Inbox
  * Get-InboxItems
  * Get-FocusedInboxItems
  * Get-CalendarItems
  * Get-FolderItems
  * New-ContactFolder
  * New-CalendarFolder
  * Set-FolderRetentionTag
  * Get-AllMailboxItems
  * Get-TaggedProperty
  * Get-NamedProperty
  * Get-FolderPath
  * Get-ArchiveFolder
  * Get-MailboxSettingsReport
  * Get-People
  * Get-UserPhotoMetaData
  * Get-UserPhoto
  * Get-MailboxUser
  * Get-CalendarGroups
  * Invoke-EnumCalendarGroups
  * New-Folder
  * Rename-Folder
  * Update-Folder
  * Invoke-DeleteFolder
  * Update-FolderClass
  * Get-FolderClass
  * GetExtendedPropList
  * GetFolderRetentionTags

## Internal functions
  * Get-AppSettings
  * Get-HTTPClient
  * Convert-FromBase64StringWithNoPadding
  * Invoke-DecodeToken
  * New-JWTToken
  * Invoke-CreateSelfSignedCert
  * Show-OAuthWindow
  * Invoke-RestGet
  * Invoke-RestPOST
  * Invoke-RestPatch
  * Invoke-RestDELETE
  * Invoke-RefreshAccessToken
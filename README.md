# Exch-Rest Getting Started Guide
The Exch-Rest module is a PowerShell module for the Office 365 and Exchange 2016 REST API that allows you to access the functionality provided by the [Graph API](https://developer.microsoft.com/en-us/graph) 

## Setup

#### Module installation

The Module is availble from the PowerShell Gallery at https://www.powershellgallery.com/packages/Exch-Rest and can be installed on Windows 10 and Windows 8 using 

Install-Module Exch-Rest

Import-Module Exch-Rest

Or you can use the following to download and use the following steps can be used to install the module from the GitHub repo
    # Set constants
	$SourceCodeURL = "https://codeload.github.com/gscales/Exch-Rest/zip/master"
	$UserModuleHome = "~\Documents\WindowsPowerShell\Modules"

	# Download a zip of the source code
	Invoke-WebRequest -Uri $SourceCodeURL -OutFile "~\Exch-Rest-master.zip"

	# Unblock the downloaded file
	Unblock-File "~\Exch-Rest-master.zip"

	# Extract the zip
	Expand-Archive "~\Exch-Rest-master.zip" -DestinationPath $UserModuleHome

	# Remove "-master" from the name
	Move-Item "$UserModuleHome\Exch-Rest-master" "$UserModuleHome\Exch-Rest"

	# Delete the downloaded source code
	Remove-Item "~\Exch-Rest-master.zip"

	# Import the module
	Import-Module -Name Exch-Rest
    

## Connecting and Authenticating ##

To connect to a Mailbox which will start the authentication proces that will allow you to then use the cmdlets defined in the module use the following 

    Connect-EXRMailbox -Mailbox gscales@datarumble.com


#### Application registration
The Office 365 / Exchange 2016 REST API uses OAuth 2.0 to authenticate users. This means that people using an application that use this API do not need to give you their username/password. Instead, they authenticate against a central authentication system (e.g. Azure AD, Active Directory) and you get back a token which is then passed to the API endpoint for authentication. You can then give your application permission to use that token to do a limited number of things for a specific period of time.

However, to use OAuth tokens you must register an application in Azure before you can use the Exch-Rest functions.

You have two options when it comes to doing this, the most secure option is to register your own Application and assign just the permissions (or Permission Grants) you want the Module cmdlets to have based on what data you want to access using the module. The other option is to use one of the default Application registrations that have been registed for use in a Tenant that is owned by the modules Author Glen Scales, if no ClientId is specified when using Connect-exrMailbox a menu will be presented with the different Application Registration options and the permissions that thoses registration will need if used. eg

![](https://gscales.github.io/Exch-Rest/GetttingStarted/MenuCaptureGettingStarted.PNG)

If you select one of these Id's the first time you run this in a Tennant it will prompt for [administrative consent](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-devhowto-multi-tenant-overview#understanding-user-and-admin-consent) for the permissions that the cmdlet will need to access the Mailbox Items (or OneDrive, Directory,Report etc) eg you should see a screen simular to 

![](https://gscales.github.io/Exch-Rest/GetttingStarted/AppRegCapture.PNG)

If you want to create your own Application registration and permission grants (which is recommended) there is a good walk through of the application registration process is provided by Jason Johnston at [https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md](https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md).

The following is an overview of the steps you can take to create an application registration:
  
* Browse to [http://dev.office.com/app-registration](http://dev.office.com/app-registration) and login into your Azure tenant
  * Click `+ New application registration`, fill out the options, and click `Create`
    * Name: `<Name-of-app-users-will-see>`
    * Application type: `Native`
    * Sign-on URL: `http://localhost`
  * Click your newly created application and note the `Application ID`. You will need this later as your `Client ID`.
  * Click `Redirect URIs`, you should see `http://localhost`. Replace that entry with `urn:ietf:wg:oauth:2.0:oob`
  * Click `Required permissions` and then click `+ Add`
  * Click `1 Select an API`, click `Office 365 Exchange Online (Microsoft.Exchange)`, and then click `Select`
  * Check off all the permissions that you wish to use, and then click `Select`. (Note: there seems to be a bug with the CheckAll button so you may have to individually check off each permission)
  * Click `Done`

Once you have done this you can set the Id you created to be the default Appilcation registration everytime you use Connect-ExrMailbox (saving your from need to enter it again), to do this select the Number 5 Option from the below menu 

It will the prompt you to enter the clentId that was created when you create you application registration and the redirectURI (generally this will be urn:ietf:wg:oauth:2.0:oob if you have used a native app) eg

![](https://gscales.github.io/Exch-Rest/GetttingStarted/defaultappset.PNG)

once the default Application has been set the console menu will no longer show when you use the Connect-EXRMailbox cmdlet, if you do what to show the menu you can use the -ShowMenu switch with this cmdlet eg

    Connect-EXRMailbox -Mailbox gscales@datarumble.com -ShowMenu

#Using the Module

Once you have sucesfully authenticated and your token has been cached locally you can start using the cmdlets defined in the module.

## The -MailboxName parameter ##

 Most cmdlets have a -MailboxName switch which will control which mailbox a cmdlet is run against eg lets look an example

![](https://gscales.github.io/Exch-Rest/GetttingStarted/getInboxExampeNoMailbox.PNG)

In the above example No MailboxName is used so the Mailbox that was used in the orginal Connect-ExrMailbox cmd will be used. This is because that MailboxName is cached in the AccessToken.

If you want to connect to a partcular Mailbox you should use the -MailboxName parameter as follows

![](https://gscales.github.io/Exch-Rest/GetttingStarted/getInboxExampeMailbox.PNG)

Cmdlets that don't connect to a specific Mailbox don't need the Mailboxname pass in eg like Get-EXRUsers which will retreive all the user objects in the Azure Directory eg

![](https://gscales.github.io/Exch-Rest/GetttingStarted/get-users.PNG)

# Other Useful Examples #

### Show the last email from the Focused Inbox ###

    Get-EXRLastInboxEmail -MailboxName gscales@datarumble.com -Focused

### Show the last email from the Other (Focused Inbox) ###

    Get-EXRLastInboxEmail -MailboxName gscales@datarumble.com -Other

### Exporting the Contacts Folder of a Mailbox to CSV ###

To Export the contacts in a user Contacts folder you can use the Export-EXRContactFolderToCSV cmdlet

    Export-EXRContactFolderToCSV -mailboxname mec@datarumble.com -FileName c:\temp\MailboxContacts.csv


### Showing the Meeting Rooms in your Tennant ###

    Find-EXRRooms

### Create a new user created folder in a Mailbox`s Inbox Folder ###

    New-EXRFolder -MailboxName gscales@datarumble.com -ParentFolderPath '\Inbox' -DisplayName "My New Folder for Processing"

### Show the Folder Retention Tags applied to a folder ###

    Get-EXRFolderFromPath -MailboxName gscales@datarumble.com -FolderPath \Inbox -PropList (Get-EXRItemRetentionTags)


## Reporting ##

The Microsoft graph API provides access to the Office 365 usage reports for tennants and the module allows access to these reports eg getting the Mailbox Sizes and usage for the last 7 days


![](https://gscales.github.io/Exch-Rest/GetttingStarted/get-MailboxUsages.PNG)

To vary the report duration you can pass in a different duration in the -PeriodType parameter eg to use 30 days instead of the default 7 use

![](https://gscales.github.io/Exch-Rest/GetttingStarted/get-MailboxUsages30.PNG)













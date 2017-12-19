# Exch-Rest
The Exch-Rest module is a PowerShell module for the Office 365 and Exchange 2016 REST API that allows you to access the functionality provided by the [Graph API](https://developer.microsoft.com/en-us/graph) 

## Setup
#### Application registration
The Office 365 / Exchange 2016 REST API uses OAuth 2.0 to authenticate users. This means that people using your app do not need to give you their username/password. Instead, they authenticate against a central authentication system (e.g. Azure AD, Active Directory) and get back a token. They can then give your application permission to use that token to do a limited number of things for a specific period of time.

However, to use OAuth tokens you must register an application in Azure before you can use the Exch-Rest functions.

You have two options when it comes to doing this, the most secure option is to register your own ApplicationId and assign just the permission you want the App to have based on what data you want to access. The other option is to use one of the default ApplicationId's that have been registed for use in the App, if no ApplicationId is specified when using Connect-exrMailbox a menu will be presented with the different ApplicationId's and the permission that these Id's will need.

If you want to create you own ApplicationId there is a good walk through of the application registration process is provided by Jason Johnston at [https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md](https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md).

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

#### Module installation

The Module is availble from the PowerShell Gallery at https://www.powershellgallery.com/packages/Exch-Rest and can be installed on Windows 10 using 

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
    

## Connecting and Authenticatng ##

To connect to a Mailbox using the following 

    Connect-EXRMailbox -Mailbox gscales@datarumble.com

The will start the OAuth Autentication process, if you have

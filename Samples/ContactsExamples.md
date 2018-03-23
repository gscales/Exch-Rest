# Using Contacts in Exch-REST Module #

Within an Office365 Mailbox there are two different types of Contacts you may encounter, Mailbox contacts that exist as Exchange Store Items stored in a Mailbox Folder (or Group,Shared Mailbox) or Directory contacts which the two main types are Azure (AD) Contacts or Mail-Enabled Users but other outliers such as Mail Enabled PublicFolders and other Mail-Enabled objects may also exists. 


## Pre Reqs#

Before being able to use any of the contact cmdlets in the Exch-REST module you need to have established a connection to  Office365 as outlined in Getting Started Guide https://github.com/gscales/Exch-Rest

## Creating a new Contact in a Mailbox (New-EXRContact) ##

The New-EXRContact cmdlet can be used to create a contact in any mailbox specifying the most common properties

    New-EXRContact -MailboxName mec@datarumble.com -FirstName "FirstName" -LastName "Surname of Contact" -EmailAddress "EmailAddress@domain.com" -MobilePhone 1111-222-333

If you want to also upload a photo as part of the contact you can use -photo switch to specify a file-name that contains the contact photo you want to be used for that contact eg


    New-EXRContact -MailboxName mec@datarumble.com -FirstName "FirstName" -LastName "Surname of Contact" -EmailAddress "EmailAddress@domain.com" -MobilePhone 1111-222-333 -photo 'c:\photo\Johnsmith.jpg'

To create a contact in a contact Folder other then the default in the Mailbox use

    New-EXRContact -MailboxName mec@datarumble.com -FirstName "FirstName" -LastName "Surname of Contact" -EmailAddress "EmailAddress@domain.com" -MobilePhone 1111-222-333 -ContactFolder SubContactFolder

The return type of the New-EXRContact is the new Contact object which contains the Id of the new contact you just created that can be used to make further modifications . Eg if you wanted to set particular Extended Properties on the new contact you just created for example the Gender of the contact which there is currently no Strongly typed property for you could use the following snippet.

    $NewContact = New-EXRContact -MailboxName mec@datarumble.com -FirstName "FirstName" -LastName "Surname of Contact" -EmailAddress "EmailAddress@domain.com" -MobilePhone 1111-222-333
    $PropList = @()
    $Gender = Get-EXRTaggedProperty -DataType Short -Id 0x3A4D -Value 2
    $PropList += $Gender
    Set-EXRContact -id $NewContact.id -ExPropList $PropList

Strongly Type properties available in New-EXRContact

	.PARAMETER MailboxName
		A description of the MailboxName parameter.
	
	.PARAMETER DisplayName
		A description of the DisplayName parameter.
	
	.PARAMETER FirstName
		A description of the FirstName parameter.
	
	.PARAMETER LastName
		A description of the LastName parameter.
	
	.PARAMETER EmailAddress
		A description of the EmailAddress parameter.
	
	.PARAMETER CompanyName
		A description of the CompanyName parameter.
	
	.PARAMETER Credentials
		A description of the Credentials parameter.
	
	.PARAMETER Department
		A description of the Department parameter.
	
	.PARAMETER Office
		A description of the Office parameter.
	
	.PARAMETER BusinssPhone
		A description of the BusinssPhone parameter.
	
	.PARAMETER MobilePhone
		A description of the MobilePhone parameter.
	
	.PARAMETER HomePhone
		A description of the HomePhone parameter.
	
	.PARAMETER IMAddress
		A description of the IMAddress parameter.
	
	.PARAMETER Street
		A description of the Street parameter.
	
	.PARAMETER City
		A description of the City parameter.
	
	.PARAMETER State
		A description of the State parameter.
	
	.PARAMETER PostalCode
		A description of the PostalCode parameter.
	
	.PARAMETER Country
		A description of the Country parameter.
	
	.PARAMETER JobTitle
		A description of the JobTitle parameter.
	
	.PARAMETER Notes
		A description of the Notes parameter.
	
	.PARAMETER Photo
		A description of the Photo parameter.
	
	.PARAMETER FileAs
		A description of the FileAs parameter.
	
	.PARAMETER WebSite
		A description of the WebSite parameter.
	
	.PARAMETER Title
		A description of the Title parameter.
	
	.PARAMETER ContactFolder
		A description of the Folder parameter.
	
	.PARAMETER EmailAddressDisplayAs
		A description of the EmailAddressDisplayAs parameter.




## Modifying an Existing Mailbox Contact (Set-EXRContact) ##

Set-EXRContact can be used to modify an existing contact using any of the strongly typed properties listed above or any Extended properties that you pass in. To use this cmdlet you need to know the Id of the contact that you want to modify this can come from either the return object from New-EXRContac or  enumerating the contacts using Get-EXRContacts or finding a Contact using Search-EXRContact. Once you have the Id of a Contact to use this cmdlet you do something like
    
    Set-EXRContact -Id $ExistingContact.id -FirstName "New First Name" -Department "New Department"

## Enumerating Contact Folders in a Mailbox (Get-EXRContactFolders) ##

The Get-EXRContactFolders cmdlet is used to enumerate all the available contact folders in a Mailbox it will enumerate childfolders of any folder that has a childfolder count greater then 0 (it does this by getting the extended property for the ChildFolderCount) eg


![](https://gscales.github.io/Exch-Rest/Contacts/GetContactsFolder.PNG)

## Getting a Contacts Folder in a Mailbox if you know the name of the folder (Get-EXRContactsFolder) ##

The Get-EXRContactsFolder cmdlet is used to get a Contact Folder that you know the name of eg

    Get-EXRContactsFolder -MailboxName mailbox@domain.com -FolderName newcont

## Creating a New Contact Folder in a Mailbox (New-EXRContactFolder) ##

The New-EXRContactFolder cmdlet is used for creating a Contact Folder to do this you just specify the Mailbox and the name of the Folder you want create.

    New-EXRContactsFolder -MailboxName mailbox@domain.com -DisplayName newcont


## Enumerating Contacts in a Contacts Folder in a Mailbox (Get-EXRContacts) ##

The Get-EXRContacts cmdlets enumerate all the contacts from a particular Contacts Folder eg to enumerate the contacts from the default contacts folder in a Mailbox use

    Get-EXRContacts -mailbox mailbox@domain.com

To enumerate contacts from a folder other then the default ContactsFolder include the name of the contacts Folder in the -ContactsFolderName parameter eg

    Get-EXRContacts -mailbox mailbox@domain.com -ContactsFolderName n12

## Enumerating Directory Contacts from Azure AD (Get-EXRDirectoryContacts) ##

Directory Contacts are those contacts available in the Global Address Lists and created through the Exchange Administration Center (or they may have been synced from an onpremise Exchange server). Currently the Graph endpoint to get these Contacts is in beta but still should work okay. To enumerate these contacts use
	
	Get-EXRDirectoryContacts


## Finding a Contact in a Mailbox (Search-EXRContactFolder) ##

**Find a Contact with a Particular DisplayName (Exact Match)**

using a Filter

    Search-EXRContacts -MailboxName user@domain -displayName 'And the displayName is'

using KQL

    Search-EXRContacts -MailboxName user@domain -displayNameKQL 'And the displayName is' 

**Find a Contact where the DisplayName contains a phrase**

using a Filter

    Search-EXRContacts -MailboxName user@domain -displayNameContains 'And the displayName' 

using KQL

    Search-EXRContacts  -MailboxName user@domain -KQL "displayName:'And the displayName'" 


**Find a Contact where the DisplayName starts with a word or letters**
using a Filter

    Search-EXRContacts -MailboxName user@domain -DisplayNameStartsWith 'And the displayName' 

**Find a Contact with a Particular Email Address **

using a Filter to search based on email domain

    Search-EXRContacts -MailboxName user@domain -emailaddress '@domain.com'

using KQL

    Search-EXRContacts -MailboxName user@domain -emailaddressKQL 'myaddress@domain.com' 

## Get the Photo associated with a particular Contact (Get-EXRContactPhoto) ##

The Get-EXRContactPhoto gets the photo that is used on a particular Mailbox contact. This cmdlet need to the Id of the contact in question to work and filename you want to save the Contact Photo image as for example you could use Search-EXRContact to find a contact and then Get-EXRContactPhoto get download the photo for that contact eg

    $Contact = Search-EXRContacts -MailboxName user@domain -emailaddressKQL 'gscales@datarumble.com'
	Get-EXRContactPhoto -id $Contact.id -SaveASFileName c:\temp\contactphoto.jpg 

## Update the Photo associated with a particular Contact (Set-EXRContactPhoto) ##

The Set-EXRContactPhoto sets the users contact photo for a existing Mailbox contact, to use this cmdlet you need to pass in the Id of the contact and the filename of the jpg you wish to use as a contact photo eg

    $Contact = Search-EXRContacts -MailboxName user@domain -emailaddressKQL 'gscales@datarumble.com'
	Set-EXRContactPhoto -id $Contact.id -FileName c:\temp\contactphoto.jpg 

## Exporting Contacts to Vcards (Export-EXRContactToVcard and Export-EXRDirectoryContactToVcard) ##

The module has the ability to export contacts to VCard files (vcard version 2.1), to support the two different type of Mailbox contacts there are two different Export cmdlets.

The Export-EXRContactToVcard exports a Mailbox Contact to Vcard to use this you need to pass in the Id of the contact you wish to export the Filename to export to and if you want to include the contactphoto or not. For example to Search and export a contact to Vcard with contact photo use

    $Contact = Search-EXRContacts -MailboxName user@domain -emailaddressKQL 'gscales@datarumble.com'
	Export-EXRContactToVcard -id $Contact.id -FileName c:\temp\mvcard.vcf -IncludePhoto

The Export-EXRDirectoryContactToVcard exports a Directory Contact to Vcard to use this you need to pass in the Id of the contact you wish to export the Filename to export to. eg

    Export-EXRDirectoryContactToVcard -id $Contact.id -FileName c:\temp\directoryvcard.vcf 

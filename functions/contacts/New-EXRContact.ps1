function New-EXRContact
{
<#
	.SYNOPSIS
		Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
	
	.DESCRIPTION
		Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API
		
		Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
	
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
	
	.PARAMETER Folder
		A description of the Folder parameter.
	
	.PARAMETER EmailAddressDisplayAs
		A description of the EmailAddressDisplayAs parameter.
	
	.PARAMETER useImpersonation
		A description of the useImpersonation parameter.
	
	.EXAMPLE
		Example 1 To create a contact in the default contacts folder
		New-EXCContact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe"
		
	.EXAMPLE
		Example 2 To create a contact and add a contact picture
		New-EXCContact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -photo 'c:\photo\Jdoe.jpg'
		
	.EXAMPLE
		Example 3 To create a contact in a user created subfolder
		New-EXCContact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -Folder "\MyCustomContacts"
		
		This cmdlet uses the EmailAddress as unique key so it wont let you create a contact with that email address if one already exists.
#>
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $false)]
		[string]
		$MailboxName,
		
		[Parameter(Position = 1, Mandatory = $false)]
		[string]
		$DisplayName,
		
		[Parameter(Position = 2, Mandatory = $true)]
		[string]
		$FirstName,
		
		[Parameter(Position = 3, Mandatory = $true)]
		[string]
		$LastName,
		
		[Parameter(Position = 4, Mandatory = $true)]
		[string]
		$EmailAddress,
		
		[Parameter(Position = 5, Mandatory = $false)]
		[string]
		$CompanyName,
		
		
		[Parameter(Position = 7, Mandatory = $false)]
		[string]
		$Department,
		
		[Parameter(Position = 8, Mandatory = $false)]
		[string]
		$Office,
		
		[Parameter(Position = 9, Mandatory = $false)]
		[string]
		$BusinssPhone,
		
		[Parameter(Position = 10, Mandatory = $false)]
		[string]
		$MobilePhone,
		
		[Parameter(Position = 11, Mandatory = $false)]
		[string]
		$HomePhone,
		
		[Parameter(Position = 12, Mandatory = $false)]
		[string]
		$IMAddress,
		
		[Parameter(Position = 13, Mandatory = $false)]
		[string]
		$Street,
		
		[Parameter(Position = 14, Mandatory = $false)]
		[string]
		$City,
		
		[Parameter(Position = 15, Mandatory = $false)]
		[string]
		$State,
		
		[Parameter(Position = 16, Mandatory = $false)]
		[string]
		$PostalCode,
		
		[Parameter(Position = 17, Mandatory = $false)]
		[string]
		$Country,
		
		[Parameter(Position = 18, Mandatory = $false)]
		[string]
		$JobTitle,
		
		[Parameter(Position = 19, Mandatory = $false)]
		[string]
		$Notes,
		
		[Parameter(Position = 20, Mandatory = $false)]
		[string]
		$Photo,
		
		[Parameter(Position = 21, Mandatory = $false)]
		[string]
		$FileAs,
		
		[Parameter(Position = 22, Mandatory = $false)]
		[string]
		$WebSite,
		
		[Parameter(Position = 23, Mandatory = $false)]
		[string]
		$Title,
		
		[Parameter(Position = 24, Mandatory = $false)]
		[string]
		$Folder,
		
		[Parameter(Position = 25, Mandatory = $false)]
		[string]
		$EmailAddressDisplayAs
		
	
		
		
	)
	Begin
	{

		if($AccessToken -eq $null)
        {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if($AccessToken -eq $null){
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
         if([String]::IsNullOrEmpty($MailboxName)){
            $MailboxName = $AccessToken.mailbox
        }  
		$HttpClient = Get-HTTPClient($MailboxName)
		$EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users"

		if($ContactsFolder -eq $null){
			$RequestURL = $EndPoint + "('$MailboxName')/Contacts/"
		}
		else{
			$RequestURL = $EndPoint + "('$MailboxName')/Contacts/" + $ContactsFolder.Id + "/"
		}
		if(![String]::IsNullOrEmpty($DisplayName)){
			$Subject = $DisplayName
		}
		$NewMessage = "{" + "`r`n"
		if(![String]::IsNullOrEmpty($FirstName)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"givenName`": `"" + $FirstName + "`"" + "`r`n"
		}
		if(![String]::IsNullOrEmpty($LastName)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"surname`": `"" + $LastName + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($DisplayName)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"displayName`": `"" + $DisplayName + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($CompanyName)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"companyName`": `"" + $CompanyName + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($Department)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"department`": `"" + $Department + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($Office)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"officeLocation`": `"" + $Office + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($BusinssPhone)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"businessphones`": [`"" + $BusinssPhone + "`"]" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($MobilePhone)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"mobilephone`": `"" + $MobilePhone + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($HomePhone)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"homephones`": [`"" + $HomePhone + "`"]" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($Street)){
		if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"businessAddress`": {" + "`r`n"       
			$NewMessage += "`"street`": `"" + $Street + "`"" + "`r`n"	
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"state`": `"" + $State + "`"" + "`r`n"	
            if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"city`": `"" + $City + "`"" + "`r`n"	
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"countryOrRegion`": `"" + $Country + "`"" + "`r`n"	
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"postalCode`": `"" + $PostalCode + "`"" + "`r`n"	
			$NewMessage += "}" + "`r`n"
		}
		if(![String]::IsNullOrEmpty($EmailAddress)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"emailAddresses`":[ {" + "`r`n"       
			if(![String]::IsNullOrEmpty($EmailAddressDisplayAs)){
				$NewMessage += "`"name`": `"" + $EmailAddressDisplayAs + "`"" + "`r`n"	
			}
			else{
				$NewMessage += "`"name`": `"" + $EmailAddress + "`"" + "`r`n"	
			}
			$NewMessage += ",`"address`": `"" + $EmailAddress + "`"" + "`r`n"
			$NewMessage += "}]" + "`r`n"
		}
		if(![String]::IsNullOrEmpty($IMAddress)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"imAddresses`": [`"" + $IMAddress + "`"]" + "`r`n"				
		}
		if(![String]::IsNullOrEmpty($WebSite)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"businessHomePage`": `"" + $WebSite + "`"" + "`r`n"		
		}
		if (![String]::IsNullOrEmpty($Notes))
		{			
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"personalNotes`": `"" + $Notes + "`"" + "`r`n"	
		}
		if(![String]::IsNullOrEmpty($JobTitle)){
			if ($NewMessage.Length -gt 5) { $NewMessage += "," }
			$NewMessage += "`"jobTitle`": `"" + $JobTitle + "`"" + "`r`n"	
		}
		$NewMessage += "}"
		Write-Host "Contact Created"
        return Invoke-RestPOST -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -Content $NewMessage

		
	}
}

function Export-EXRContactToVcard {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 3, Mandatory = $true)] [string]$id,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$IncludePhoto,
        [Parameter(Position = 5, Mandatory = $true)]
        [string]
        $FileName
    )  
    Begin {
        if ($AccessToken -eq $null) {
            $AccessToken = Get-ProfiledToken -MailboxName $MailboxName  
            if ($AccessToken -eq $null) {
                $AccessToken = Get-EXRAccessToken -MailboxName $MailboxName       
            }                 
        }
        if ([String]::IsNullOrEmpty($MailboxName)) {
            $MailboxName = $AccessToken.mailbox
        }  
        $HttpClient = Get-HTTPClient -MailboxName $MailboxName
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "users" 
        $RequestURL = $EndPoint + "('" + $MailboxName + "')/Contacts('" + $id + "')"
        $Contact = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        $ufilename = Get-UniqueFileName -FileName $FileName
        Set-content -path $ufilename "BEGIN:VCARD"
        add-content -path $ufilename "VERSION:2.1"
        $givenName = ""
        if ($Contact.givenName -ne $null) {
            $givenName = $Contact.givenName
        }
        $surname = ""
        if ($Contact.surname -ne $null) {
            $surname = $Contact.surname
        }
        add-content -path $ufilename ("N:" + $surname + ";" + $givenName)
        add-content -path $ufilename ("FN:" + $Contact.displayName)
        $Department = "";
        if ($Contact.department -ne $null) {
            $Department = $Contact.department
        }
		
        $CompanyName = "";
        if ($Contact.companyName -ne $null) {
            $CompanyName = $Contact.companyName
        }
        add-content -path $ufilename ("ORG:" + $CompanyName + ";" + $Department)
        if ($Contact.jobTitle -ne $null) {
            add-content -path $ufilename ("TITLE:" + $Contact.jobTitle)
        }
        if ($Contact.mobilePhone) {
            add-content -path $ufilename ("TEL;CELL;VOICE:" + $Contact.mobilePhone)
        }
        if ($Contact.homePhones) {
            add-content -path $ufilename ("TEL;HOME;VOICE:" + $Contact.homePhones)
        }
        if ($Contact.homePhones) {
            add-content -path $ufilename ("TEL;WORK;VOICE:" + $Contact.homePhones)
        }
        if ($Contact.businessPhones) {
            add-content -path $ufilename ("TEL;WORK;FAX:" + $Contact.businessPhones)
        }
        if ($Contact.businessHomePage -ne $null) {
            add-content -path $ufilename ("URL;WORK:" + $Contact.businessHomePage)
        }
        if ($Contact.businessAddress -ne $null) {
            if ($Contact.businessAddress.countryOrRegion -ne $null) {
                $Country = $Contact.businessAddress.countryOrRegion.Replace("`n", "")
            }
            if ($Contact.businessAddress.city -ne $null) {
                $City = $Contact.businessAddress.city.Replace("`n", "")
            }
            if ($Contact.businessAddress.street -ne $null) {
                $Street = $Contact.businessAddress.street.Replace("`n", "")
            }
            if ($Contact.businessAddress.state -ne $null) {
                $State = $Contact.businessAddress.state.Replace("`n", "")
            }
            if ($Contact.businessAddress.postalCode -ne $null) {
                $PCode = $Contact.businessAddress.postalCode.Replace("`n", "")
            }
            $addr = "ADR;WORK;PREF:;" + $Country + ";" + $Street + ";" + $City + ";" + $State + ";" + $PCode + ";" + $Country
            add-content -path $ufilename $addr
        }
        if ($Contact.imAddresses -ne $null) {
            add-content -path $ufilename ("X-MS-IMADDRESS:" + $Contact.imAddresses)
		}
		$emCnt = 1;
		foreach($emailAddress in $Contact.emailAddresses){
			if($emCnt -eq 1){
				add-content -path $ufilename ("EMAIL;PREF;INTERNET:" + $emailAddress.address)
			}
			else{
				add-content -path $ufilename ("EMAIL;" + $emCnt + ";INTERNET:" + $emailAddress.address)
			}
			$emCnt++		
		}
       # add-content -path $ufilename ("EMAIL;PREF;INTERNET:" + $Contact.emailAddresses[0].address)		
		
        if ($IncludePhoto.IsPresent) {
            $photoBytes = Get-EXRContactPhoto -id $id
            add-content -path $ufilename "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
            $ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
            add-content -path $ufilename $ImageString
            add-content -path $ufilename "`r`n"
        }
        add-content -path $ufilename "END:VCARD"
        Write-Host "Contact exported to $ufilename"

    } 
}





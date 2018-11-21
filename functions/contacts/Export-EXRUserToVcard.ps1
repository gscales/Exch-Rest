function Export-EXRUserToVcard {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 3, Mandatory = $false)] [string]$id,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$IncludePhoto,
        [Parameter(Position = 5, Mandatory = $true)][string]$FileName,
        [Parameter(Position = 6, Mandatory = $false)] [string]$UPN
        

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
        if([String]::IsNullOrEmpty($UPN)){
            $RequestURL = $EndPoint + "('" + $id + "')"
        }else{
            $RequestURL = $EndPoint + "('" + $UPN + "')"
        }    
        $RequestURL += "?`$select=id,mail,displayName,userPrincipalName,givenName,surname,department,companyName,jobTitle,mobilePhone,homePhones,faxNumber,streetAddress,state,officeLocation,country,city,postalCode,proxyAddresses"
        $User = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        $ufilename = Get-UniqueFileName -FileName $FileName
        Set-content -path $ufilename "BEGIN:VCARD"
        add-content -path $ufilename "VERSION:2.1"
        $givenName = ""
        if ($User.givenName -ne $null) {
            $givenName = $User.givenName
        }
        $surname = ""
        if ($User.surname -ne $null) {
            $surname = $User.surname
        }
        add-content -path $ufilename ("N:" + $surname + ";" + $givenName)
        add-content -path $ufilename ("FN:" + $User.displayName)
        $Department = "";
        if ($User.department -ne $null) {
            $Department = $User.department
        }
		
        $CompanyName = "";
        if ($User.companyName -ne $null) {
            $CompanyName = $User.companyName
        }
        add-content -path $ufilename ("ORG:" + $CompanyName + ";" + $Department)
        if ($User.jobTitle -ne $null) {
            add-content -path $ufilename ("TITLE:" + $User.jobTitle)
        }
        if ($User.mobilePhone) {
            add-content -path $ufilename ("TEL;CELL;VOICE:" + $User.mobilePhone)
        }
        if ($User.homePhones) {
            add-content -path $ufilename ("TEL;HOME;VOICE:" + $User.homePhones)
        }
        if ($User.businessPhones) {
            add-content -path $ufilename ("TEL;WORK;VOICE:" + $User.businessPhones)
        }
        if ($User.faxNumber) {
            add-content -path $ufilename ("TEL;WORK;FAX:" + $User.faxNumber)
        }
        if ($User.businessHomePage -ne $null) {
            add-content -path $ufilename ("URL;WORK:" + $User.businessHomePage)
        }
        if ($User.streetAddress -ne $null) {
            if ($User.country -ne $null) {
                $Country = $User.country.Replace("`n", "")
            }
            if ($User.city -ne $null) {
                $City = $User.city.Replace("`n", "")
            }
            if ($User.streetAddress -ne $null) {
                $Street = $User.streetAddress.Replace("`n", "")
            }
            if ($User.state -ne $null) {
                $State = $User.state.Replace("`n", "")
            }
            if ($User.postalCode -ne $null) {
                $PCode = $User.postalCode.Replace("`n", "")
            }
            if($User.officeLocation -ne $null){
                $officeLocation =  $User.officeLocation.Replace("`n", "")
            }            
            $addr = "ADR;WORK;PREF:;" + $officeLocation + ";" + $Street + ";" + $City + ";" + $State + ";" + $PCode + ";" + $Country
            add-content -path $ufilename $addr
        }
        if ($User.imAddresses -ne $null) {
            add-content -path $ufilename ("X-MS-IMADDRESS:" + $User.imAddresses)
		}
        $emCnt = 2;
        add-content -path $ufilename ("EMAIL;PREF;INTERNET:" + $User.mail)
		foreach($emailAddress in $User.proxyAddresses){
            $proxy = $emailAddress.Replace("smtp:","").Replace("SMTP:","")
            if($proxy.tolower() -ne $user.mail){
                add-content -path $ufilename ("EMAIL;" + $emCnt + ";INTERNET:" + $proxy)
                $emCnt++	
            }
	
		}
       # add-content -path $ufilename ("EMAIL;PREF;INTERNET:" + $User.emailAddresses[0].address)		
		
        if ($IncludePhoto.IsPresent) {
            $photoBytes = Get-EXRUserPhoto -TargetUser $id
            add-content -path $ufilename "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
            $ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
            add-content -path $ufilename $ImageString
            add-content -path $ufilename "`r`n"
        }
        add-content -path $ufilename "END:VCARD"
        Write-Host "Contact exported to $ufilename"

    } 
}





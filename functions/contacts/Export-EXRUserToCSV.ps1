function Export-EXRUserToCSV {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 3, Mandatory = $false)] [string]$id,
        [Parameter(Position = 4, Mandatory = $false)] [switch]$IncludePhoto,
        [Parameter(Position = 5, Mandatory = $false)][string]$FileName,
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
        $RequestURL += "?`$select=id,mail,displayName,userPrincipalName,givenName,surname,department,companyName,jobTitle,mobilePhone,businessPhones,homePhones,faxNumber,streetAddress,state,officeLocation,country,city,postalCode,proxyAddresses"
        $User = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
		
        if ($IncludePhoto.IsPresent) {
           
            
            if([String]::IsNullOrEmpty($UPN)){
                $photoBytes = Get-EXRUserPhoto -TargetUser $id
            }else{
                $photoBytes = Get-EXRUserPhoto -TargetUser $UPN
            }   
            if($photoBytes){
                $ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
                Add-Member -InputObject $User -NotePropertyName UserPhoto -NotePropertyValue $ImageString
            }else{
                Add-Member -InputObject $User -NotePropertyName UserPhoto -NotePropertyValue ""
            }

        }
        if([String]::IsNullOrEmpty($FileName)){
            return $User
        }else{
            $User | Export-Csv -Path $FileName -NoTypeInformation
        }       
    } 
}





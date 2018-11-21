function Get-EXRGroupMembers {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [string]
        $MailboxName,
		
        [Parameter(Position = 1, Mandatory = $false)]
        [psobject]
        $AccessToken,
        
        [Parameter(Position = 2, Mandatory = $false)]
        [psobject]
        $GroupId,

        [Parameter(Position = 3, Mandatory = $false)]
        [switch]
        $ContactPropsOnly,

        [Parameter(Position = 4, Mandatory = $false)] [switch]$IncludePhoto



		
    )
    Process {
		
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
        $EndPoint = Get-EndPoint -AccessToken $AccessToken -Segment "groups" 
        $RequestURL = $EndPoint + "/" + $GroupId + "/members"
        if($ContactPropsOnly.IsPresent){
            $RequestURL += "?`$select=id,mail,displayName,userPrincipalName,givenName,surname,department,companyName,jobTitle,mobilePhone,businessPhones,homePhones,faxNumber,streetAddress,state,officeLocation,country,city,postalCode,proxyAddresses"
        }       
        $Result = Invoke-RestGET -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName
        foreach($Message in $Result.value){
            if ($IncludePhoto.IsPresent) {     
            
                $photoBytes = Get-EXRUserPhoto -TargetUser $Message.userPrincipalName
                if($photoBytes){
                    $ImageString = [System.Convert]::ToBase64String($photoBytes, [System.Base64FormattingOptions]::InsertLineBreaks)
                    Add-Member -InputObject $Message -NotePropertyName UserPhoto -NotePropertyValue $ImageString
                }else{
                    Add-Member -InputObject $Message -NotePropertyName UserPhoto -NotePropertyValue ""
                }
    
            }
            Write-Output $Message
        }
       
    }
}

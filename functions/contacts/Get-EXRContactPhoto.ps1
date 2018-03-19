function Get-EXRContactPhoto {

	   [CmdletBinding()] 
    param( 
        [Parameter(Position = 1, Mandatory = $false)] [psobject]$AccessToken,
        [Parameter(Position = 2, Mandatory = $false)] [string]$MailboxName,
        [Parameter(Position = 4, Mandatory = $true)] [string]$id,
        [Parameter(Position = 6, Mandatory = $false)] [string]$SaveASFileName
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
        $RequestURL = $EndPoint + "('" + $MailboxName + "')/Contacts('" + $id + "')/Photo/`$value"  
        $Result = Invoke-RestGet -RequestURL $RequestURL -HttpClient $HttpClient -AccessToken $AccessToken -MailboxName $MailboxName -NoJSON  
        if([String]::IsNullOrEmpty($SaveASFileName)){
            
            Write-Output $Result.ReadAsByteArrayAsync().Result
        }
        else{
            [System.IO.File]::WriteAllBytes($SaveASFileName,$Result.ReadAsByteArrayAsync().Result)
        }
		
		

    } 
}

